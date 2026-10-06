using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailSorter
{
    // The settings fields, shared by the ribbon's settings dialog and the page under File > Options.
    // Implements Outlook.PropertyPage so Outlook can host it. Outlook only calls Apply after the page reports a
    // change, so on the Options page (saveOnChange) every edit is saved straight away instead.
    [ComVisible(true)]
    public class SettingsPanel : UserControl, Outlook.PropertyPage
    {
        readonly Settings settings;
        readonly TextBox apiKey = new TextBox { UseSystemPasswordChar = true, Width = 330 };
        readonly TextBox model = new TextBox { Width = 330 };
        readonly CheckBox auto = new CheckBox { Text = "Sort new mail automatically", AutoSize = true };
        readonly CheckBox categories = new CheckBox { Text = "Let Claude create new categories", AutoSize = true };
        readonly CheckBox startup = new CheckBox { Text = "Sort Inbox when Outlook starts", AutoSize = true };

        public SettingsPanel(Settings settings, Action createFolders, bool saveOnChange = false)
        {
            this.settings = settings;
            apiKey.Text = settings.ApiKey;
            model.Text = settings.Model;
            auto.Checked = settings.AutoProcess;
            categories.Checked = settings.CreateCategories;
            startup.Checked = settings.SortOnStartup;

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(12),
            };
            panel.Controls.Add(new Label { Text = "Anthropic API key", AutoSize = true });
            panel.Controls.Add(apiKey);
            panel.Controls.Add(new Label { Text = "Model (blank = " + Settings.DefaultModel + ")", AutoSize = true });
            panel.Controls.Add(model);
            panel.Controls.Add(auto);
            panel.Controls.Add(categories);
            panel.Controls.Add(startup);
            var folders = new Button { Text = "Create starter folders", AutoSize = true };
            folders.Click += (s, e) => createFolders();
            panel.Controls.Add(folders);
            Controls.Add(panel);

            if (saveOnChange)
            {
                apiKey.TextChanged += (s, e) => Apply();
                model.TextChanged += (s, e) => Apply();
                auto.CheckedChanged += (s, e) => Apply();
                categories.CheckedChanged += (s, e) => Apply();
                startup.CheckedChanged += (s, e) => Apply();
            }
        }

        // Called by the dialog's Save button, on every edit on the Options page, and by Outlook if it asks.
        public void Apply()
        {
            settings.ApiKey = apiKey.Text.Trim();
            settings.Model = model.Text.Trim();
            settings.AutoProcess = auto.Checked;
            settings.CreateCategories = categories.Checked;
            settings.SortOnStartup = startup.Checked;
            settings.Save();
        }

        public bool Dirty => false;

        public void GetPageInfo(ref string helpFile, ref int helpContext) { }
    }
}
