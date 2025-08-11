using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    public partial class ColorPaletteControl : UserControl
    {
        private FlowLayoutPanel palettePanel;
        private Button btnEdit;
        private Core.ColorPaletteDefinition palette;

        public ColorPaletteControl()
        {
            InitializeComponent();
            System.Diagnostics.Debug.WriteLine("ColorPaletteControl initialized");
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.White;
            this.Name = "ColorPaletteControl";
            this.Dock = DockStyle.Fill;

            // Edit button at top-left
            btnEdit = new Button { Text = "Edit", Location = new Point(6, 6), Size = new Size(50, 22), FlatStyle = FlatStyle.Standard };
            btnEdit.Click += (s, e) => EditPalette();
            Controls.Add(btnEdit);
            
            // Removed hue slider from the task pane per request

            // Palette panel for color rows
            palettePanel = new FlowLayoutPanel
            {
                Location = new Point(6, 34),
                Size = new Size(128, 560),
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(0),
                BackColor = Color.White
            };
            Controls.Add(palettePanel);

            // Responsive layout hooks
            this.SizeChanged += (s, e) => LayoutControls();
            palettePanel.SizeChanged += (s, e) => ResizeRowWidths();

            // Load palette
            palette = Core.ColorPaletteStorage.LoadOrDefault();
            BuildPaletteRows();

            this.ResumeLayout(false);
            LayoutControls();
        }

        private void BuildPaletteRows()
        {
            palettePanel.SuspendLayout();
            palettePanel.Controls.Clear();

            // First special row: Clear Fill, Clear Outline (square + X)
            var firstRow = new Panel { Width = palettePanel.ClientSize.Width, Height = 22, Margin = new Padding(0, 0, 0, 4) };
            // Square checkerboard for "no fill"
            var noFill = new Panel { Width = 22, Height = 22, Location = new Point(0, 0), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Cursor = Cursors.Hand };
            noFill.Paint += (s, e) =>
            {
                var g = e.Graphics; var rect = new Rectangle(0, 0, 22, 22);
                using (var light = new SolidBrush(Color.FromArgb(220, 220, 220)))
                using (var dark = new SolidBrush(Color.FromArgb(180, 180, 180)))
                {
                    g.FillRectangle(light, new Rectangle(0, 0, 11, 11));
                    g.FillRectangle(dark, new Rectangle(11, 0, 11, 11));
                    g.FillRectangle(dark, new Rectangle(0, 11, 11, 11));
                    g.FillRectangle(light, new Rectangle(11, 11, 11, 11));
                }
            };
            noFill.Click += (s, e) => ClearFill();

            var noOutline = new Label { Width = 22, Height = 22, Location = new Point(26, 0), Text = "X", TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Red, Font = new Font("Segoe UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, BackColor = Color.White };
            noOutline.Click += (s, e) => ClearOutline();

            // Clear text color (A without color)
            var noText = new Button { Width = 22, Height = 22, Location = new Point(52, 0), Text = "A", FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Black, BackColor = Color.White };
            noText.FlatAppearance.BorderSize = 1;
            noText.Cursor = Cursors.Hand;
            noText.Click += (s, e) => ClearTextColor();

            firstRow.Controls.Add(noFill);
            firstRow.Controls.Add(noOutline);
            firstRow.Controls.Add(noText);
            palettePanel.Controls.Add(firstRow);

            foreach (var hex in palette.Colors)
            {
                var c = Core.ColorPaletteDefinition.FromHex(hex);
                var row = new Panel { Width = palettePanel.ClientSize.Width, Height = 22, Margin = new Padding(0, 2, 0, 2) };

                // Color sample (solid rectangle)
                var colorBox = new Panel { Width = 22, Height = 22, Location = new Point(0, 0), BackColor = c, BorderStyle = BorderStyle.FixedSingle };
                colorBox.Tag = c;
                colorBox.Cursor = Cursors.Hand;
                colorBox.Click += (s, e) => ApplyFill((Color)((Panel)s).Tag);
                
                // Border button (outlined square)
                var btnBorder = new Button { Width = 22, Height = 22, Location = new Point(26, 0), FlatStyle = FlatStyle.Flat, BackColor = Color.White };
                btnBorder.FlatAppearance.BorderColor = c;
                btnBorder.FlatAppearance.BorderSize = 2;
                btnBorder.Cursor = Cursors.Hand;
                btnBorder.Paint += (s, e) =>
                {
                    var rect = new Rectangle(5, 5, 12, 12);
                    using (var pen = new Pen(c, 2))
                        e.Graphics.DrawRectangle(pen, rect);
                };
                btnBorder.Tag = c;
                btnBorder.Click += (s, e) => ApplyBorder((Color)((Button)s).Tag);

                // Text button (A)
                var btnText = new Button { Width = 22, Height = 22, Location = new Point(52, 0), Text = "A", FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = c, BackColor = Color.White };
                btnText.FlatAppearance.BorderSize = 1;
                btnText.Cursor = Cursors.Hand;
                btnText.Tag = c;
                btnText.Click += (s, e) => ApplyText((Color)((Button)s).Tag);

                row.Controls.Add(colorBox);
                row.Controls.Add(btnBorder);
                row.Controls.Add(btnText);
                palettePanel.Controls.Add(row);
            }
            palettePanel.ResumeLayout(true);
        }

        private void ApplyBorder(Color color)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app?.ActiveWindow?.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    s.Line.Visible = Office.MsoTriState.msoTrue;
                    s.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
                    if (s.Line.Weight < 1) s.Line.Weight = 1;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ApplyBorder failed: {ex.Message}"); }
        }

        private void ApplyFill(Color color)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app?.ActiveWindow?.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    s.Fill.Visible = Office.MsoTriState.msoTrue;
                    s.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ApplyFill failed: {ex.Message}"); }
        }

        private void ApplyText(Color color)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app?.ActiveWindow?.Selection;
                if (sel == null || sel.ShapeRange == null || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    if (s.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        var tr = s.TextFrame.TextRange;
                        tr.Font.Color.RGB = ColorTranslator.ToOle(color);
                    }
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ApplyText failed: {ex.Message}"); }
        }

        private void ClearTextColor()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app?.ActiveWindow?.Selection;
                if (sel == null || sel.ShapeRange == null || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    if (s.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        var tr = s.TextFrame.TextRange;
                        try { tr.Font.Color.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorDark1; }
                        catch { tr.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); }
                    }
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ClearTextColor failed: {ex.Message}"); }
        }

        private void EditPalette()
        {
            using (var dlg = new ColorPaletteEditorDialog(palette))
            {
                if (dlg.ShowDialog(FindForm()) == DialogResult.OK)
                {
                    // Save per-PC
                    Core.ColorPaletteStorage.Save(dlg.Palette);
                    palette = dlg.Palette;
                    BuildPaletteRows();
                }
            }
        }

        private void ClearFill()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app?.ActiveWindow?.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    s.Fill.Visible = Office.MsoTriState.msoFalse;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ClearFill failed: {ex.Message}"); }
        }

        private void ClearOutline()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app?.ActiveWindow?.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    s.Line.Visible = Office.MsoTriState.msoFalse;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"ClearOutline failed: {ex.Message}"); }
        }

        private static int Clamp(int v, int min, int max) => v < min ? min : (v > max ? max : v);

        private static Color HslToColor(int h, int s, int l)
        {
            double H = h / 359.0, S = s / 100.0, L = l / 100.0;
            if (S == 0) { int v = (int)Math.Round(L * 255); return Color.FromArgb(v, v, v); }
            double q = L < 0.5 ? L * (1 + S) : L + S - L * S;
            double p = 2 * L - q;
            Func<double, double> hue2rgb = (t) =>
            {
                if (t < 0) t += 1; if (t > 1) t -= 1;
                if (t < 1.0 / 6) return p + (q - p) * 6 * t;
                if (t < 1.0 / 2) return q;
                if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
                return p;
            };
            double r = hue2rgb(H + 1.0 / 3), g = hue2rgb(H), b = hue2rgb(H - 1.0 / 3);
            return Color.FromArgb((int)Math.Round(r * 255), (int)Math.Round(g * 255), (int)Math.Round(b * 255));
        }

        // Rounded rainbow hue slider control
        private class HueSlider : Control
        {
            private int value; // 0..359
            public int Value
            {
                get => value;
                set { int v = Math.Max(0, Math.Min(359, value)); if (v != this.value) { this.value = v; Invalidate(); ValueChanged?.Invoke(this, EventArgs.Empty); } }
            }
            public event EventHandler ValueChanged;

            public HueSlider()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
                this.Height = 18; this.Width = 128;
            }

            protected override void OnMouseDown(MouseEventArgs e) { base.OnMouseDown(e); UpdateFromMouse(e); }
            protected override void OnMouseMove(MouseEventArgs e) { base.OnMouseMove(e); if (e.Button == MouseButtons.Left) UpdateFromMouse(e); }

            private void UpdateFromMouse(MouseEventArgs e)
            {
                int x = Math.Max(0, Math.Min(Width - 1, e.X));
                int v = (int)Math.Round((x / (double)(Width - 1)) * 359);
                Value = v;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
                var rect = new Rectangle(0, 0, Width - 1, Height - 1);

                // Rounded rect path for clipping and border
                using (var gp = new GraphicsPath())
                {
                    int r = rect.Height; // full rounded ends
                    gp.AddArc(rect.X, rect.Y, r, r, 90, 180);
                    gp.AddArc(rect.Right - r, rect.Y, r, r, 270, 180);
                    gp.CloseFigure();

                    // Draw rainbow gradient bitmap and clip to rounded rect
                    using (var bmp = new Bitmap(rect.Width, rect.Height))
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            int h = (int)Math.Round((x / (double)(bmp.Width - 1)) * 359);
                            var c = HslToColor(h, 100, 50);
                            using (var pen = new Pen(c))
                            {
                                using (var gBmp = Graphics.FromImage(bmp))
                                {
                                    gBmp.DrawLine(pen, x, 0, x, bmp.Height);
                                }
                            }
                        }
                        g.SetClip(gp);
                        g.DrawImageUnscaled(bmp, rect.X, rect.Y);
                        g.ResetClip();
                    }

                    // Border shadow
                    using (var penShadow = new Pen(Color.FromArgb(60, 0, 0, 0))) g.DrawPath(penShadow, gp);
                    using (var penBorder = new Pen(Color.FromArgb(160, 160, 160))) g.DrawPath(penBorder, gp);
                }

                // Circular indicator
                int cx = (int)Math.Round((value / 359.0) * (rect.Width - 1)) + rect.X;
                int radius = rect.Height + 4; // a bit larger than bar
                var thumbRect = new Rectangle(cx - radius / 2, rect.Y - (radius - rect.Height) / 2, radius, radius);
                using (var shadow = new SolidBrush(Color.FromArgb(60, 0, 0, 0))) g.FillEllipse(shadow, thumbRect.X + 1, thumbRect.Y + 1, thumbRect.Width, thumbRect.Height);
                using (var br = new SolidBrush(Color.White)) g.FillEllipse(br, thumbRect);
                using (var pen = new Pen(Color.FromArgb(140, 140, 140))) g.DrawEllipse(pen, thumbRect);
            }
        }

        private void LayoutControls()
        {
            const int padding = 6;
            if (btnEdit != null)
            {
                btnEdit.SetBounds(padding, padding, btnEdit.Width, btnEdit.Height);
            }
            if (palettePanel != null)
            {
                int top = (btnEdit?.Bottom ?? padding) + padding;
                int width = Math.Max(60, this.Width - (padding * 2));
                int height = Math.Max(50, this.Height - top - padding);
                palettePanel.SetBounds(padding, top, width, height);
            }
            ResizeRowWidths();
        }

        private void ResizeRowWidths()
        {
            if (palettePanel == null) return;
            int w = palettePanel.ClientSize.Width;
            foreach (Control c in palettePanel.Controls)
            {
                if (c is Panel row)
                {
                    row.Width = w;
                }
            }
        }
    }
}
