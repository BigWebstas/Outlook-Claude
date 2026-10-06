using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailSorter
{
    public class FolderEntry
    {
        public string Path;
        public Outlook.Folder Folder;
    }

    // Folders Claude may file into for one mailbox, and the Junk folder for spam.
    public class StoreFolders
    {
        public List<FolderEntry> Targets = new List<FolderEntry>();
        public Outlook.Folder Junk;
        // Mail already in Junk or Deleted Items is left alone.
        public HashSet<string> SkipIds = new HashSet<string>();
    }

    // Built once per batch so each email does not re-walk the folder tree.
    public class SortContext
    {
        // Folders Claude may never file mail into; spam goes to Junk through its own path.
        static readonly Outlook.OlDefaultFolders[] Excluded =
        {
            Outlook.OlDefaultFolders.olFolderDeletedItems,
            Outlook.OlDefaultFolders.olFolderJunk,
            Outlook.OlDefaultFolders.olFolderSentMail,
            Outlook.OlDefaultFolders.olFolderDrafts,
            Outlook.OlDefaultFolders.olFolderOutbox,
        };

        readonly Dictionary<string, StoreFolders> byStore = new Dictionary<string, StoreFolders>();
        readonly Outlook.Application app;
        List<string> categoryNames;

        public SortContext(Outlook.Application app) { this.app = app; }

        public Outlook.Categories Categories => app.Session.Categories;

        public List<string> CategoryNames
        {
            get
            {
                if (categoryNames == null)
                    categoryNames = Categories.Cast<Outlook.Category>().Select(c => c.Name).ToList();
                return categoryNames;
            }
        }

        public StoreFolders For(Outlook.Store store)
        {
            if (!byStore.TryGetValue(store.StoreID, out var folders))
                byStore[store.StoreID] = folders = Build(store);
            return folders;
        }

        static StoreFolders Build(Outlook.Store store)
        {
            var result = new StoreFolders();
            var excludedIds = new HashSet<string>();
            foreach (var kind in Excluded)
            {
                try
                {
                    var folder = (Outlook.Folder)store.GetDefaultFolder(kind);
                    excludedIds.Add(folder.EntryID);
                    if (kind == Outlook.OlDefaultFolders.olFolderJunk) result.Junk = folder;
                    if (kind == Outlook.OlDefaultFolders.olFolderJunk ||
                        kind == Outlook.OlDefaultFolders.olFolderDeletedItems)
                        result.SkipIds.Add(folder.EntryID);
                }
                catch (COMException) { } // This mailbox has no such folder.
            }

            var prefix = store.GetRootFolder().FolderPath + "\\";
            void Walk(Outlook.Folders folders)
            {
                foreach (Outlook.Folder f in folders)
                {
                    if (excludedIds.Contains(f.EntryID)) continue; // also skips its subfolders
                    if (f.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        var path = f.FolderPath.StartsWith(prefix) ? f.FolderPath.Substring(prefix.Length) : f.FolderPath;
                        result.Targets.Add(new FolderEntry { Path = path.Replace('\\', '/'), Folder = f });
                    }
                    Walk(f.Folders);
                }
            }
            Walk(store.GetRootFolder().Folders);
            return result;
        }
    }

    public class Sorter
    {
        const int MaxBodyChars = 20000;
        const string InternetMessageIdTag = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";

        readonly Settings settings;

        public Sorter(Settings settings) { this.settings = settings; }

        public static string MessageId(Outlook.MailItem item)
        {
            try { return item.PropertyAccessor.GetProperty(InternetMessageIdTag) as string; }
            catch (COMException) { return null; }
        }

        public async Task SortAsync(Outlook.MailItem item, SortContext ctx)
        {
            var folder = (Outlook.Folder)item.Parent;
            var store = folder.Store;
            var folders = ctx.For(store);
            if (folders.SkipIds.Contains(folder.EntryID)) return;

            var verdict = await ClaudeClient.AskAsync(
                settings, DescribeEmail(item), folders.Targets.Select(f => f.Path), ctx.CategoryNames);

            if (verdict.Spam)
            {
                item.FlagStatus = Outlook.OlFlagStatus.olNoFlag;
                item.Save();
                if (folders.Junk != null) item.Move(folders.Junk);
                return;
            }

            var changed = false;
            var newCategories = ResolveCategories(verdict.Tags, ctx)
                .Where(c => !SplitCategories(item.Categories).Contains(c, StringComparer.OrdinalIgnoreCase))
                .ToList();
            if (newCategories.Count > 0)
            {
                item.Categories = string.Join(", ", SplitCategories(item.Categories).Concat(newCategories));
                changed = true;
            }
            if (verdict.Important && item.FlagStatus != Outlook.OlFlagStatus.olFlagMarked)
            {
                item.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;
                changed = true;
            }
            if (changed) item.Save();

            // Only accept folders that exist, so an email cannot make Claude file mail anywhere else.
            var target = folders.Targets.FirstOrDefault(f => f.Path == verdict.Folder);
            if (target != null && target.Folder.EntryID != folder.EntryID) item.Move(target.Folder);
        }

        static string DescribeEmail(Outlook.MailItem item)
        {
            var body = (item.Body ?? "").Trim();
            if (body.Length > MaxBodyChars) body = body.Substring(0, MaxBodyChars) + "\n[truncated]";
            return $"From: {item.SenderName} <{SenderAddress(item)}>\nTo: {item.To}\nSubject: {item.Subject}\n" +
                   $"Date: {item.ReceivedTime:u}\n\n{body}";
        }

        static string SenderAddress(Outlook.MailItem item)
        {
            if (item.SenderEmailType == "EX")
            {
                try { return item.Sender?.GetExchangeUser()?.PrimarySmtpAddress ?? item.SenderEmailAddress; }
                catch (COMException) { }
            }
            return item.SenderEmailAddress;
        }

        static List<string> SplitCategories(string categories) =>
            (categories ?? "").Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(c => c.Trim()).Where(c => c.Length > 0).ToList();

        // Keeps categories that exist, and creates new ones only when the user allowed it.
        List<string> ResolveCategories(IEnumerable<string> names, SortContext ctx)
        {
            var resolved = new List<string>();
            var created = 0;
            foreach (var raw in names)
            {
                var name = raw.Trim();
                var existing = ctx.CategoryNames.FirstOrDefault(n => n.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    resolved.Add(existing);
                }
                else if (settings.CreateCategories && name.Length > 0 && created < ClaudeClient.MaxNewCategoriesPerEmail)
                {
                    var color = (Outlook.OlCategoryColor)(1 + ctx.CategoryNames.Count % 8);
                    ctx.Categories.Add(name, color);
                    ctx.CategoryNames.Add(name);
                    resolved.Add(name);
                    created++;
                }
            }
            return resolved.Distinct().ToList();
        }
    }
}
