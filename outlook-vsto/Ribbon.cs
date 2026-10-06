using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailSorter
{
    // Ribbon button, message and folder right-click items, and the spinner shown while mail is processed.
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        const int FrameCount = 12;
        const int FrameMs = 100;
        const string ExplorerRibbonId = "Microsoft.Outlook.Explorer";

        const string Xml = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabMail"">
        <group id=""claudeGroup"" label=""Claude"">
          <button id=""sortSelected"" size=""large"" getLabel=""GetSortLabel"" getImage=""GetSortImage""
                  screentip=""Sort the selected emails with Claude"" onAction=""OnSortSelected""/>
          <button id=""sortInbox"" size=""normal"" label=""Sort Inbox now"" showImage=""false"" onAction=""OnSortInbox""/>
          <button id=""settings"" size=""normal"" label=""Claude settings"" showImage=""false"" onAction=""OnSettings""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuMailItem"">
      <button id=""ctxSortSelected"" label=""Sort with Claude"" onAction=""OnSortSelected""/>
    </contextMenu>
    <contextMenu idMso=""ContextMenuFolder"">
      <button id=""ctxSortFolder"" label=""Sort folder with Claude"" onAction=""OnSortFolder""/>
    </contextMenu>
  </contextMenus>
</customUI>";

        readonly ThisAddIn addIn;
        readonly Timer timer = new Timer { Interval = FrameMs };
        Office.IRibbonUI ui;
        Bitmap[] frames;
        Bitmap idle;
        int frame;

        public Ribbon(ThisAddIn addIn)
        {
            this.addIn = addIn;
            timer.Tick += (s, e) =>
            {
                frame = (frame + 1) % FrameCount;
                ui?.InvalidateControl("sortSelected");
            };
        }

        public string GetCustomUI(string ribbonId) => ribbonId == ExplorerRibbonId ? Xml : null;

        public void OnLoad(Office.IRibbonUI ribbonUI) { ui = ribbonUI; }

        // Called by the add-in whenever the number of running tasks changes.
        public void UpdateStatus()
        {
            if (addIn.Busy)
            {
                if (!timer.Enabled) timer.Start();
            }
            else
            {
                timer.Stop();
                frame = 0;
            }
            ui?.InvalidateControl("sortSelected");
        }

        public string GetSortLabel(Office.IRibbonControl control) =>
            addIn.Busy ? "Claude: working…" : "Sort with Claude";

        public Bitmap GetSortImage(Office.IRibbonControl control)
        {
            if (!addIn.Busy) return idle ?? (idle = Draw(null));
            if (frames == null)
            {
                frames = new Bitmap[FrameCount];
                for (var i = 0; i < FrameCount; i++) frames[i] = Draw(360f / FrameCount * i);
            }
            return frames[frame];
        }

        public void OnSortSelected(Office.IRibbonControl control) { var _ = addIn.SortSelected(); }

        public void OnSortFolder(Office.IRibbonControl control)
        {
            var _ = addIn.SortFolder(control.Context as Outlook.Folder);
        }

        public void OnSortInbox(Office.IRibbonControl control) { var _ = addIn.SortInboxes(true); }

        public void OnSettings(Office.IRibbonControl control) { addIn.ShowSettings(); }

        // The add-on icon: envelope and sparkle, plus a spinning arc when `spinnerAngle` is set.
        static Bitmap Draw(float? spinnerAngle)
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.ScaleTransform(0.5f, 0.5f);
                var orange = Color.FromArgb(0xd9, 0x77, 0x57);
                using (var brush = new SolidBrush(orange)) g.FillRectangle(brush, 4, 14, 46, 34);
                using (var pen = new Pen(Color.White, 4) { LineJoin = LineJoin.Round, StartCap = LineCap.Round, EndCap = LineCap.Round })
                    g.DrawLines(pen, new[] { new PointF(7, 18), new PointF(27, 33), new PointF(47, 18) });
                using (var gold = new SolidBrush(Color.FromArgb(0xe9, 0xa2, 0x3b)))
                using (var outline = new Pen(Color.White, 2))
                {
                    var sparkle = new[]
                    {
                        new PointF(50, 4), new PointF(54, 10), new PointF(60, 14), new PointF(54, 18),
                        new PointF(50, 24), new PointF(46, 18), new PointF(40, 14), new PointF(46, 10),
                    };
                    g.FillPolygon(gold, sparkle);
                    g.DrawPolygon(outline, sparkle);
                }
                if (spinnerAngle.HasValue)
                {
                    g.FillEllipse(Brushes.White, 34, 34, 28, 28);
                    using (var ring = new Pen(Color.Gainsboro, 4)) g.DrawEllipse(ring, 39, 39, 18, 18);
                    using (var arc = new Pen(orange, 4) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        g.DrawArc(arc, 39, 39, 18, 18, spinnerAngle.Value, 90);
                }
            }
            return bmp;
        }
    }
}
