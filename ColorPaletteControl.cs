using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    /// <summary>
    /// Custom FlowLayoutPanel that enforces fixed width and scroll bar behavior
    /// </summary>
    public class CustomFlowLayoutPanel : FlowLayoutPanel
    {
        public CustomFlowLayoutPanel()
        {
            // Set initial properties with reduced width to account for scroll bar
            this.MinimumSize = new Size(120, 0);
            this.MaximumSize = new Size(120, 0);
            this.AutoSize = false;
            this.AutoSizeMode = AutoSizeMode.GrowOnly;
            this.Anchor = AnchorStyles.Top | AnchorStyles.Left; // Anchor to top-left only
        }
        
        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            
            // Force fixed width to prevent scroll bar width changes
            if (this.Width != 120)
            {
                this.Width = 120;
            }
            
            // Ensure horizontal scroll is completely disabled
            this.HorizontalScroll.Visible = false;
            this.HorizontalScroll.Enabled = false;
            this.HorizontalScroll.Maximum = 0;
            
            // Set minimum size to prevent layout changes
            if (this.MinimumSize.Width != 120)
            {
                this.MinimumSize = new Size(120, this.MinimumSize.Height);
            }
        }
        
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            
            // Re-enforce width if it changes
            if (this.Width != 120)
            {
                this.Width = 120;
            }
        }
        
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            
            // Force width to 120 pixels on any resize
            if (this.Width != 120)
            {
                this.Width = 120;
            }
        }
        
        // Note: SetBounds cannot be overridden as it's not virtual in Control class
        // Instead, we use event handlers and properties to enforce width
        
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            
            // Ensure scroll bar settings are applied after handle creation
            this.HorizontalScroll.Visible = false;
            this.HorizontalScroll.Enabled = false;
            this.HorizontalScroll.Maximum = 0;
            this.VerticalScroll.Visible = false;
  
            
            // Start timer to enforce width
            StartWidthEnforcementTimer();
        }
        
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                // Disable horizontal scroll bar completely
                cp.Style &= ~0x00100000; // WS_HSCROLL
                return cp;
            }
        }
        
        protected override void OnScroll(ScrollEventArgs se)
        {
            // Always allow vertical scrolling, horizontal is disabled by other means
            base.OnScroll(se);
        }
        
        // Add a timer to continuously enforce width
        private Timer widthEnforcementTimer;
        
        private void StartWidthEnforcementTimer()
        {
            if (widthEnforcementTimer == null)
            {
                widthEnforcementTimer = new Timer();
                widthEnforcementTimer.Interval = 100; // Check every 100ms
                widthEnforcementTimer.Tick += (s, e) => 
                {
                    if (this.Width != 120)
                    {
                        this.Width = 120;
                    }
                    // Ensure horizontal scroll is always disabled
                    if (this.HorizontalScroll.Visible || this.HorizontalScroll.Enabled)
                    {
                        this.HorizontalScroll.Visible = false;
                        this.HorizontalScroll.Enabled = false;
                        this.HorizontalScroll.Maximum = 0;
                    }
                };
                widthEnforcementTimer.Start();
            }
        }
        
        protected override void Dispose(bool disposing)
        {
            if (disposing && widthEnforcementTimer != null)
            {
                widthEnforcementTimer.Stop();
                widthEnforcementTimer.Dispose();
                widthEnforcementTimer = null;
            }
            base.Dispose(disposing);
        }
        
        protected override void OnParentChanged(EventArgs e)
        {
            base.OnParentChanged(e);
            
            // Ensure width is maintained when parent changes
            if (this.Width != 120)
            {
                this.Width = 120;
            }
        }
        
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            
            // Ensure width is maintained when visibility changes
            if (this.Width != 120)
            {
                this.Width = 120;
            }
        }
    }

    public partial class ColorPaletteControl : UserControl
    {
        private FlowLayoutPanel palettePanel;
        private Button btnEdit;
        private Core.ColorPaletteDefinition palette;

        // Override properties to prevent resizing
        public override bool AutoSize
        {
            get => false;
            set { } // Ignore attempts to change AutoSize
        }

        // Note: AutoSizeMode cannot be overridden as it's not virtual in UserControl
        // We'll handle this in the constructor and other methods

        public override Size MinimumSize
        {
            get => new Size(120, 0);
            set { } // Ignore attempts to change MinimumSize
        }

        public override Size MaximumSize
        {
            get => new Size(120, 0);
            set { } // Ignore attempts to change MaximumSize
        }

        // Method to enforce fixed width properties
        private void EnforceFixedWidth()
        {
            if (this.Width != 120)
            {
                this.Width = 120;
            }
            if (this.AutoSize)
            {
                this.AutoSize = false;
            }
            if (this.AutoSizeMode != AutoSizeMode.GrowOnly)
            {
                this.AutoSizeMode = AutoSizeMode.GrowOnly;
            }
        }

        // Timer to continuously enforce fixed width properties
        private Timer mainWidthEnforcementTimer;

        private void StartMainWidthEnforcementTimer()
        {
            if (mainWidthEnforcementTimer == null)
            {
                mainWidthEnforcementTimer = new Timer();
                mainWidthEnforcementTimer.Interval = 100; // Check every 100ms
                mainWidthEnforcementTimer.Tick += (s, e) => 
                {
                    EnforceFixedWidth();
                };
                mainWidthEnforcementTimer.Start();
            }
        }

        // Override SetBounds to enforce fixed width
        protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
        {
            // Always enforce fixed width of 132px
            if ((specified & BoundsSpecified.Width) != 0)
            {
                width = 132;
            }
            base.SetBoundsCore(x, y, width, height, specified);
        }

        // Override SetClientSizeCore to prevent client size changes
        protected override void SetClientSizeCore(int x, int y)
        {
            // Ensure client size maintains fixed width
            if (x != 120) // 120px content area
            {
                x = 120;
            }
            base.SetClientSizeCore(x, y);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (mainWidthEnforcementTimer != null)
                {
                    mainWidthEnforcementTimer.Stop();
                    mainWidthEnforcementTimer.Dispose();
                    mainWidthEnforcementTimer = null;
                }
            }
            base.Dispose(disposing);
        }

        public ColorPaletteControl()
        {
            InitializeComponent();
            
            // Set control properties to prevent resizing
            this.MinimumSize = new Size(132, 0); // 120px content + 12px padding
            this.MaximumSize = new Size(132, 0);
            this.AutoSize = false;
            this.AutoSizeMode = AutoSizeMode.GrowOnly; // Set this property directly
            
            // Start the width enforcement timer
            StartMainWidthEnforcementTimer();
            
            System.Diagnostics.Debug.WriteLine("ColorPaletteControl initialized");
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.White;
            this.Name = "ColorPaletteControl";
            this.Dock = DockStyle.Fill;
            this.Size = new Size(132, 600); // Fixed size
            this.MinimumSize = new Size(132, 0);
            this.MaximumSize = new Size(132, 0);

            // Edit button at top-left - Ensure text is visible with better styling
            btnEdit = new Button 
            { 
                Text = "Edit", 
                Location = new Point(6, 6), 
                Size = new Size(60, 24), // Increased size for better visibility
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.Black,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter
            };
            
            // Add border to make button more visible
            btnEdit.FlatAppearance.BorderSize = 1;
            btnEdit.FlatAppearance.BorderColor = Color.Gray;
            
            // Add hover effects to make button more interactive
            btnEdit.MouseEnter += (s, e) => btnEdit.BackColor = Color.FromArgb(240, 240, 240);
            btnEdit.MouseLeave += (s, e) => btnEdit.BackColor = Color.White;
            
            btnEdit.Click += (s, e) => EditPalette();
            Controls.Add(btnEdit);
            
            // Removed hue slider from the task pane per request

            // Create custom palette panel with fixed scroll bar behavior
            palettePanel = new CustomFlowLayoutPanel
            {
                Location = new Point(6, 36), // Adjusted for larger button
                Size = new Size(120, 560), // Reduced width to 120px
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
            var firstRow = new Panel { Width = 116, Height = 22, Margin = new Padding(0, 0, 0, 4) }; // Reduced width to 116px
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
                var row = new Panel { Width = 116, Height = 22, Margin = new Padding(0, 2, 0, 2) }; // Reduced width to 116px

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
                this.Height = 18; this.Width = 120; // Updated width to match palette
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
                // Fixed width to prevent scroll bar width changes
                int width = 120; // Fixed width instead of dynamic calculation
                int height = Math.Max(50, this.Height - top - padding);
                palettePanel.SetBounds(padding, top, width, height);
                
                // The custom panel will handle its own width enforcement
            }
            ResizeRowWidths();
        }

        private void ResizeRowWidths()
        {
            if (palettePanel == null) return;
            // Use fixed width to prevent scroll bar width changes
            int w = 116; // Reduced width to 116px
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
