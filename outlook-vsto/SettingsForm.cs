using System;
using System.Drawing;
using System.Windows.Forms;

namespace ClaudeMailSorter
{
    public class SettingsForm : Form
    {
        public SettingsForm(Settings settings, Action createFolders)
        {
            Text = "Claude Mail Sorter settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 310);

            var panel = new SettingsPanel(settings, createFolders) { Dock = DockStyle.Fill };
            var ok = new Button { Text = "Save", DialogResult = DialogResult.OK, Left = 190, Top = 275 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 275, Top = 275 };
            Controls.Add(ok);
            Controls.Add(cancel);
            Controls.Add(panel);
            ok.BringToFront();
            cancel.BringToFront();
            AcceptButton = ok;
            CancelButton = cancel;

            ok.Click += (s, e) => panel.Apply();
        }
    }
}
