using System.Drawing;
using System.Windows.Forms;

namespace ClaudeMailSorter
{
    public class SettingsForm : Form
    {
        readonly TextBox apiKey = new TextBox { UseSystemPasswordChar = true, Dock = DockStyle.Top };
        readonly TextBox model = new TextBox { Dock = DockStyle.Top };
        readonly CheckBox auto = new CheckBox { Text = "Sort new mail automatically", AutoSize = true };
        readonly CheckBox categories = new CheckBox { Text = "Let Claude create new categories", AutoSize = true };
        readonly CheckBox startup = new CheckBox { Text = "Sort Inbox when Outlook starts", AutoSize = true };

        public SettingsForm(Settings settings)
        {
            Text = "Claude Mail Sorter settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 270);

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
            apiKey.Width = model.Width = 330;
            panel.Controls.Add(new Label { Text = "Anthropic API key", AutoSize = true });
            panel.Controls.Add(apiKey);
            panel.Controls.Add(new Label { Text = "Model (blank = " + Settings.DefaultModel + ")", AutoSize = true });
            panel.Controls.Add(model);
            panel.Controls.Add(auto);
            panel.Controls.Add(categories);
            panel.Controls.Add(startup);

            var ok = new Button { Text = "Save", DialogResult = DialogResult.OK, Left = 190, Top = 235 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 275, Top = 235 };
            Controls.Add(ok);
            Controls.Add(cancel);
            Controls.Add(panel);
            ok.BringToFront();
            cancel.BringToFront();
            AcceptButton = ok;
            CancelButton = cancel;

            ok.Click += (s, e) =>
            {
                settings.ApiKey = apiKey.Text.Trim();
                settings.Model = model.Text.Trim();
                settings.AutoProcess = auto.Checked;
                settings.CreateCategories = categories.Checked;
                settings.SortOnStartup = startup.Checked;
                settings.Save();
            };
        }
    }
}
