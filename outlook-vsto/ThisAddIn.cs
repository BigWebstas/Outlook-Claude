using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailSorter
{
    public partial class ThisAddIn
    {
        const int InboxSweepLimit = 200;
        const int StartupSweepDelayMs = 60 * 1000;
        const int FolderConfirmThreshold = 25;
        const string Title = "Claude Mail Sorter";
        static readonly string[] StarterFolders = { "Notifications", "OTP", "Receipts", "Newsletters", "Finance", "Travel" };

        // Outlook's object model only works on the main thread, so every task runs there, one at a time.
        // That keeps bursts of new mail under API rate limits and stops the processed list being written twice.
        readonly SemaphoreSlim gate = new SemaphoreSlim(1, 1);
        // Held so the COM event sources are not garbage collected.
        readonly List<Outlook.Items> watchedInboxes = new List<Outlook.Items>();
        Settings settings;
        ProcessedStore processed;
        Sorter sorter;
        Ribbon ribbon;
        int active;

        public bool Busy => active > 0;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon(this);
            return ribbon;
        }

        void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Lets `await` continue on the main thread after an API call.
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            settings = Settings.Load();
            processed = new ProcessedStore();
            sorter = new Sorter(settings);
            WatchInboxes();
            Application.OptionsPagesAdd += OnOptionsPagesAdd;

            if (settings.SortOnStartup)
            {
                // Give IMAP accounts time to sync before looking at the Inbox.
                var timer = new System.Windows.Forms.Timer { Interval = StartupSweepDelayMs };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    var _ = SortInboxes(false);
                };
                timer.Start();
            }
        }

        void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        // Adds the settings page to File > Options.
        void OnOptionsPagesAdd(Outlook.PropertyPages pages) => pages.Add(new SettingsPanel(settings, CreateStarterFolders, saveOnChange: true), Title);

        void WatchInboxes()
        {
            foreach (Outlook.Store store in Application.Session.Stores)
            {
                try
                {
                    var items = ((Outlook.Folder)store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)).Items;
                    items.ItemAdd += OnItemAdded;
                    watchedInboxes.Add(items);
                }
                catch (COMException) { } // This mailbox has no Inbox (for example a data file).
            }
        }

        void OnItemAdded(object item)
        {
            if (settings.AutoProcess && item is Outlook.MailItem mail)
            {
                var _ = Enqueue(new List<Outlook.MailItem> { mail }, false);
            }
        }

        // --- Entry points used by the ribbon ---

        public Task SortSelected()
        {
            var items = new List<Outlook.MailItem>();
            var selection = Application.ActiveExplorer()?.Selection;
            if (selection != null)
                foreach (var o in selection)
                    if (o is Outlook.MailItem mail) items.Add(mail);
            return Enqueue(items, true);
        }

        public Task SortFolder(Outlook.Folder folder)
        {
            folder = folder ?? Application.ActiveExplorer()?.CurrentFolder as Outlook.Folder;
            if (folder == null) return Task.CompletedTask;
            var items = folder.Items.OfType<Outlook.MailItem>().ToList();
            if (items.Count > FolderConfirmThreshold &&
                MessageBox.Show($"Sort {items.Count} emails in \"{folder.Name}\" with Claude? Each one is an API call.",
                    Title, MessageBoxButtons.OKCancel) != DialogResult.OK)
                return Task.CompletedTask;
            return Enqueue(items, true);
        }

        // `explicitRequest` shows errors in a dialog; the startup sweep stays quiet.
        public Task SortInboxes(bool explicitRequest)
        {
            var pending = new List<Outlook.MailItem>();
            foreach (var items in watchedInboxes)
                foreach (var mail in items.OfType<Outlook.MailItem>())
                    if (!processed.Contains(Sorter.MessageId(mail))) pending.Add(mail);
            pending = pending.OrderByDescending(m => m.ReceivedTime).Take(InboxSweepLimit).ToList();
            return Enqueue(pending, explicitRequest);
        }

        public void ShowSettings()
        {
            using (var form = new SettingsForm(settings, CreateStarterFolders)) form.ShowDialog();
        }

        // Creates the starter folders next to the Inbox of the default mailbox, skipping any that already exist.
        void CreateStarterFolders()
        {
            try
            {
                var folders = Application.Session.DefaultStore.GetRootFolder().Folders;
                var existing = folders.Cast<Outlook.MAPIFolder>().Select(f => f.Name).ToList();
                var added = StarterFolders.Where(n => !existing.Contains(n, StringComparer.OrdinalIgnoreCase)).ToList();
                foreach (var name in added) folders.Add(name, Outlook.OlDefaultFolders.olFolderInbox);
                MessageBox.Show(added.Count == 0
                    ? "All the starter folders already exist."
                    : "Created: " + string.Join(", ", added), Title);
            }
            catch (COMException e)
            {
                MessageBox.Show("Could not create the folders: " + e.Message, Title);
            }
        }

        // --- Work queue ---

        // `force` re-sorts emails that were already sorted once; only an explicit request sets it.
        Task Enqueue(List<Outlook.MailItem> items, bool force) => RunExclusive(async () =>
        {
            if (string.IsNullOrEmpty(settings.ApiKey))
            {
                if (force) MessageBox.Show("Add your Anthropic API key under Claude settings.", Title);
                return;
            }
            var ctx = new SortContext(Application);
            var failures = 0;
            Exception lastError = null;
            foreach (var item in items)
            {
                try
                {
                    var id = Sorter.MessageId(item);
                    if (!force && processed.Contains(id)) continue;
                    await sorter.SortAsync(item, ctx);
                    processed.Add(id);
                }
                catch (Exception e)
                {
                    failures++;
                    lastError = e;
                    System.Diagnostics.Debug.WriteLine($"[{Title}] {e}");
                }
            }
            if (failures > 0 && force)
                MessageBox.Show($"{failures} message(s) failed: {lastError.Message}", Title);
        });

        async Task RunExclusive(Func<Task> work)
        {
            active++;
            ribbon?.UpdateStatus();
            try
            {
                await gate.WaitAsync();
                try { await work(); }
                finally { gate.Release(); }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"[{Title}] {e}");
            }
            finally
            {
                active--;
                ribbon?.UpdateStatus();
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
