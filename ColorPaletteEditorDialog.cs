using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using my_addin.Core;

namespace my_addin
{
    public class ColorPaletteEditorDialog : Form
    {
        // Left list: custom owner-drawn, draggable list
        private ReorderableColorList leftList;
        // Right: picker and numeric inputs
        private Panel rightPanel;
        private Panel colorPreview;
        private TextBox txtHex;
        private TextBox txtH, txtS, txtL, txtR, txtG, txtB;
        private HueSlider hueSlider;  // advanced hue slider
        private ColorCanvas colorCanvas; // 2D saturation/lightness canvas
        private Button btnAdd;
        private Button btnResetTemplate;
        private Button btnOK, btnCancel;

        public ColorPaletteDefinition Palette { get; private set; }
        private int selectedIndex = -1;
        private bool updatingControls = false;
        private int hVal, sVal, lVal, rVal, gVal, bVal;

        public ColorPaletteEditorDialog(ColorPaletteDefinition palette)
        {
            this.Palette = new ColorPaletteDefinition { Name = palette.Name, Colors = new List<string>(palette.Colors) };
            InitializeComponent();
            RebuildLeftList();
            if (Palette.Colors.Count > 0) SelectIndex(0);
        }

        private void InitializeComponent()
        {
            this.Text = "Edit Color Palette";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Size = new Size(620, 520);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false; this.MinimizeBox = false;

            leftList = new ReorderableColorList { Location = new Point(10, 10), Size = new Size(240, 430) };
            leftList.DeleteRequested += (idx) => { if (idx >= 0 && idx < Palette.Colors.Count) { Palette.Colors.RemoveAt(idx); RebuildLeftList(); if (Palette.Colors.Count > 0) SelectIndex(Math.Min(idx, Palette.Colors.Count - 1)); } };
            leftList.ReorderRequested += (from, to) => { if (from == to) return; var c = Palette.Colors[from]; Palette.Colors.RemoveAt(from); if (to < 0) to = 0; if (to > Palette.Colors.Count) to = Palette.Colors.Count; Palette.Colors.Insert(to, c); RebuildLeftList(); SelectIndex(to); };
            leftList.SelectedIndexChanged += (s, e) => { if (leftList.SelectedIndex >= 0) SelectIndex(leftList.SelectedIndex); };
            Controls.Add(leftList);

            rightPanel = new Panel { Location = new Point(260, 10), Size = new Size(340, 430), BorderStyle = BorderStyle.None };
            Controls.Add(rightPanel);

            // Color canvas (top) + hue slider below
            colorCanvas = new ColorCanvas { Location = new Point(0, 0), Size = new Size(320, 200) };
            colorCanvas.SelectionChanged += (s, e) => { if (!updatingControls) { sVal = colorCanvas.Saturation; lVal = colorCanvas.Lightness; UpdateFromHslFields(); } };
            rightPanel.Controls.Add(colorCanvas);
            hueSlider = new HueSlider { Location = new Point(0, 205), Size = new Size(320, 24) };
            hueSlider.ValueChanged += (s, e) => { if (!updatingControls) { hVal = hueSlider.Value; UpdateFromHslFields(); } };
            rightPanel.Controls.Add(hueSlider);

            // Hex + preview row with margin under picker area
            colorPreview = new Panel { Location = new Point(0, 236), Size = new Size(38, 24), BorderStyle = BorderStyle.FixedSingle };
            txtHex = new TextBox { Location = new Point(46, 236), Width = 100 };
            txtHex.TextChanged += (s, e) => { if (!updatingControls) UpdateFromHex(); };
            rightPanel.Controls.Add(colorPreview);
            rightPanel.Controls.Add(txtHex);

            // HSL and RGB as three rows of text inputs: H/R, S/G, L/B
            var hr = new TableLayoutPanel { Location = new Point(0, 268), Size = new Size(330, 28), ColumnCount = 4, RowCount = 1 };            
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            txtH = new TextBox { Width = 80, TextAlign = HorizontalAlignment.Center }; txtR = new TextBox { Width = 80, TextAlign = HorizontalAlignment.Center };
            txtH.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtH.Text, out var v)) { hVal = Clamp(v, 0, 359); UpdateFromHslFields(); } };
            txtR.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtR.Text, out var v)) { rVal = Clamp(v, 0, 255); UpdateFromRgbFields(); } };
            hr.Controls.Add(new Label { Text = "H", AutoSize = true }, 0, 0); hr.Controls.Add(txtH, 1, 0);
            hr.Controls.Add(new Label { Text = "R", AutoSize = true }, 2, 0); hr.Controls.Add(txtR, 3, 0);
            rightPanel.Controls.Add(hr);

            var sg = new TableLayoutPanel { Location = new Point(0, 300), Size = new Size(330, 28), ColumnCount = 4, RowCount = 1 };
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            txtS = new TextBox { Width = 80, TextAlign = HorizontalAlignment.Center }; txtG = new TextBox { Width = 80, TextAlign = HorizontalAlignment.Center };
            txtS.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtS.Text, out var v)) { sVal = Clamp(v, 0, 100); UpdateFromHslFields(); } };
            txtG.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtG.Text, out var v)) { gVal = Clamp(v, 0, 255); UpdateFromRgbFields(); } };
            sg.Controls.Add(new Label { Text = "S", AutoSize = true }, 0, 0); sg.Controls.Add(txtS, 1, 0);
            sg.Controls.Add(new Label { Text = "G", AutoSize = true }, 2, 0); sg.Controls.Add(txtG, 3, 0);
            rightPanel.Controls.Add(sg);

            var lb = new TableLayoutPanel { Location = new Point(0, 332), Size = new Size(330, 28), ColumnCount = 4, RowCount = 1 };
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            txtL = new TextBox { Width = 80, TextAlign = HorizontalAlignment.Center }; txtB = new TextBox { Width = 80, TextAlign = HorizontalAlignment.Center };
            txtL.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtL.Text, out var v)) { lVal = Clamp(v, 0, 100); UpdateFromHslFields(); } };
            txtB.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtB.Text, out var v)) { bVal = Clamp(v, 0, 255); UpdateFromRgbFields(); } };
            lb.Controls.Add(new Label { Text = "L", AutoSize = true }, 0, 0); lb.Controls.Add(txtL, 1, 0);
            lb.Controls.Add(new Label { Text = "B", AutoSize = true }, 2, 0); lb.Controls.Add(txtB, 3, 0);
            rightPanel.Controls.Add(lb);

            // Add + reset template
            btnAdd = new Button { Text = "Add", Location = new Point(0, 364), Width = 80 };
            btnAdd.Click += (s, e) => AddColor();
            btnResetTemplate = new Button { Text = "Add from template", Location = new Point(90, 364), Width = 140 };
            btnResetTemplate.Click += (s, e) => { this.Palette = ColorPaletteDefinition.Default(); RebuildLeftList(); if (Palette.Colors.Count > 0) SelectIndex(0); };
            rightPanel.Controls.Add(btnAdd);
            rightPanel.Controls.Add(btnResetTemplate);

            // Footer buttons
            btnOK = new Button { Text = "Apply", Location = new Point(380, 450), Width = 80, DialogResult = DialogResult.OK };
            btnCancel = new Button { Text = "Cancel", Location = new Point(470, 450), Width = 80, DialogResult = DialogResult.Cancel };
            this.AcceptButton = btnOK; this.CancelButton = btnCancel;
            Controls.Add(btnOK);
            Controls.Add(btnCancel);
        }

        // Advanced hue slider (rounded rainbow) control
        private class HueSlider : Control
        {
            private int value; // 0..359
            private Bitmap _cachedGradient;
            private GraphicsPath _cachedPath;
            
            public int Value
            {
                get => value;
                set { int v = Math.Max(0, Math.Min(359, value)); if (v != this.value) { this.value = v; Invalidate(); ValueChanged?.Invoke(this, EventArgs.Empty); } }
            }
            public event EventHandler ValueChanged;

            public HueSlider()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
                this.Height = 24; this.Width = 320;
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

                // Cache the gradient bitmap and path
                if (_cachedGradient == null || _cachedGradient.Width != rect.Width || _cachedGradient.Height != rect.Height)
                {
                    _cachedGradient?.Dispose();
                    _cachedPath?.Dispose();
                    
                    _cachedGradient = new Bitmap(rect.Width, rect.Height);
                    _cachedPath = new GraphicsPath();
                    
                    int r = rect.Height; // full rounded ends
                    _cachedPath.AddArc(rect.X, rect.Y, r, r, 90, 180);
                    _cachedPath.AddArc(rect.Right - r, rect.Y, r, r, 270, 180);
                    _cachedPath.CloseFigure();
                    
                    using (var gBmp = Graphics.FromImage(_cachedGradient))
                    {
                        double widthMinus1 = _cachedGradient.Width - 1;
                        for (int x = 0; x < _cachedGradient.Width; x++)
                        {
                            int h = (int)Math.Round((x / widthMinus1) * 359);
                            var c = HslToColor(h, 100, 50);
                            using (var pen = new Pen(c))
                            {
                                gBmp.DrawLine(pen, x, 0, x, _cachedGradient.Height);
                            }
                        }
                    }
                }

                g.SetClip(_cachedPath);
                g.DrawImageUnscaled(_cachedGradient, rect.X, rect.Y);
                g.ResetClip();

                using (var penShadow = new Pen(Color.FromArgb(60, 0, 0, 0))) g.DrawPath(penShadow, _cachedPath);
                using (var penBorder = new Pen(Color.FromArgb(160, 160, 160))) g.DrawPath(penBorder, _cachedPath);

                int cx = (int)Math.Round((value / 359.0) * (rect.Width - 1)) + rect.X;
                int radius = rect.Height + 6;
                var thumbRect = new Rectangle(cx - radius / 2, rect.Y - (radius - rect.Height) / 2, radius, radius);
                using (var shadow = new SolidBrush(Color.FromArgb(60, 0, 0, 0))) g.FillEllipse(shadow, thumbRect.X + 1, thumbRect.Y + 1, thumbRect.Width, thumbRect.Height);
                using (var br = new SolidBrush(Color.White)) g.FillEllipse(br, thumbRect);
                using (var pen = new Pen(Color.FromArgb(140, 140, 140))) g.DrawEllipse(pen, thumbRect);
            }
            
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _cachedGradient?.Dispose();
                    _cachedPath?.Dispose();
                }
                base.Dispose(disposing);
            }
        }

        // 2D color canvas for saturation/lightness selection
        private class ColorCanvas : Control
        {
            private int _hue = 0;
            private Bitmap _cachedBitmap;
            private int _cachedHue = -1;
            
            public int Hue 
            { 
                get => _hue;
                set 
                { 
                    if (_hue != value)
                    {
                        _hue = value;
                        Invalidate();
                    }
                }
            }
            public int Saturation { get; set; } = 100; // 0..100
            public int Lightness { get; set; } = 50;   // 0..100
            public event EventHandler SelectionChanged;

            public ColorCanvas()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
                Cursor = Cursors.Cross;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                
                // Cache the bitmap for the current hue
                if (_cachedBitmap == null || _cachedBitmap.Width != Width || _cachedBitmap.Height != Height || _cachedHue != Hue)
                {
                    _cachedBitmap?.Dispose();
                    _cachedBitmap = new Bitmap(Width, Height);
                    _cachedHue = Hue;
                    
                    var rect = new Rectangle(0, 0, Width, Height);
                    var bitmapData = _cachedBitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    
                    unsafe
                    {
                        byte* ptr = (byte*)bitmapData.Scan0;
                        int stride = bitmapData.Stride;
                        
                        double widthMinus1 = Width - 1;
                        double heightMinus1 = Height - 1;
                        
                        for (int y = 0; y < Height; y++)
                        {
                            byte* row = ptr + (y * stride);
                            double lightness = (1 - y / heightMinus1) * 100;
                            
                            for (int x = 0; x < Width; x++)
                            {
                                double saturation = (x / widthMinus1) * 100;
                                var col = HslToColor(Hue, (int)Math.Round(saturation), (int)Math.Round(lightness));
                                
                                int offset = x * 4;
                                row[offset] = col.B;
                                row[offset + 1] = col.G;
                                row[offset + 2] = col.R;
                                row[offset + 3] = col.A;
                            }
                        }
                    }
                    
                    _cachedBitmap.UnlockBits(bitmapData);
                }
                
                e.Graphics.DrawImageUnscaled(_cachedBitmap, 0, 0);

                // Draw indicator
                int ix = (int)Math.Round((Saturation / 100.0) * (Width - 1));
                int iy = (int)Math.Round((1 - Lightness / 100.0) * (Height - 1));
                Rectangle r = new Rectangle(ix - 6, iy - 6, 12, 12);
                using (var pen = new Pen(Color.White, 2)) e.Graphics.DrawEllipse(pen, r);
            }
            
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _cachedBitmap?.Dispose();
                }
                base.Dispose(disposing);
            }

            protected override void OnMouseDown(MouseEventArgs e) { base.OnMouseDown(e); UpdateFromMouse(e); }
            protected override void OnMouseMove(MouseEventArgs e) { base.OnMouseMove(e); if (e.Button == MouseButtons.Left) UpdateFromMouse(e); }
            private void UpdateFromMouse(MouseEventArgs e)
            {
                int x = Math.Max(0, Math.Min(Width - 1, e.X));
                int y = Math.Max(0, Math.Min(Height - 1, e.Y));
                Saturation = (int)Math.Round((x / (double)(Width - 1)) * 100);
                Lightness = (int)Math.Round((1 - y / (double)(Height - 1)) * 100);
                Invalidate();
                SelectionChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private void RebuildLeftList()
        {
            leftList.BeginUpdate();
            leftList.Items.Clear();
            foreach (var hex in Palette.Colors)
            {
                leftList.Items.Add(hex);
            }
            if (selectedIndex >= 0 && selectedIndex < leftList.Items.Count) leftList.SelectedIndex = selectedIndex; else if (leftList.Items.Count > 0) leftList.SelectedIndex = 0;
            leftList.EndUpdate();
        }

        private void SelectIndex(int index)
        {
            selectedIndex = index;
            var c = ColorPaletteDefinition.FromHex(Palette.Colors[index]);
            updatingControls = true;
            colorPreview.BackColor = c;
            txtHex.Text = ColorPaletteDefinition.ToHex(c);
            RgbToHsl(c, out hVal, out sVal, out lVal);
            rVal = c.R; gVal = c.G; bVal = c.B;
            txtH.Text = hVal.ToString(); txtS.Text = sVal.ToString(); txtL.Text = lVal.ToString();
            txtR.Text = rVal.ToString(); txtG.Text = gVal.ToString(); txtB.Text = bVal.ToString();
            hueSlider.Value = hVal;
            colorCanvas.Hue = hVal; colorCanvas.Saturation = sVal; colorCanvas.Lightness = lVal; colorCanvas.Invalidate();
            updatingControls = false;
            if (leftList.SelectedIndex != index) leftList.SelectedIndex = index;
        }

        private void UpdateFromHex()
        {
            var hex = txtHex.Text.Trim();
            if (!hex.StartsWith("#") || hex.Length != 7) return;
            var c = ColorPaletteDefinition.FromHex(hex);
            colorPreview.BackColor = c;
            updatingControls = true;
            RgbToHsl(c, out hVal, out sVal, out lVal);
            rVal = c.R; gVal = c.G; bVal = c.B;
            txtH.Text = hVal.ToString(); txtS.Text = sVal.ToString(); txtL.Text = lVal.ToString();
            txtR.Text = rVal.ToString(); txtG.Text = gVal.ToString(); txtB.Text = bVal.ToString();
            hueSlider.Value = hVal;
            colorCanvas.Hue = hVal; colorCanvas.Saturation = sVal; colorCanvas.Lightness = lVal; colorCanvas.Invalidate();
            updatingControls = false;
            if (selectedIndex >= 0) { Palette.Colors[selectedIndex] = hex; RebuildLeftList(); }
        }

        private void UpdateFromHslFields()
        {
            var c = HslToColor(hVal, sVal, lVal);
            updatingControls = true;
            rVal = c.R; gVal = c.G; bVal = c.B;
            txtR.Text = rVal.ToString(); txtG.Text = gVal.ToString(); txtB.Text = bVal.ToString();
            txtHex.Text = ColorPaletteDefinition.ToHex(c);
            colorPreview.BackColor = c;
            hueSlider.Value = hVal;
            colorCanvas.Hue = hVal; colorCanvas.Saturation = sVal; colorCanvas.Lightness = lVal; colorCanvas.Invalidate();
            updatingControls = false;
        }

        private void UpdateFromRgbFields()
        {
            var c = Color.FromArgb(Clamp(rVal, 0, 255), Clamp(gVal, 0, 255), Clamp(bVal, 0, 255));
            updatingControls = true;
            RgbToHsl(c, out hVal, out sVal, out lVal);
            txtH.Text = hVal.ToString(); txtS.Text = sVal.ToString(); txtL.Text = lVal.ToString();
            txtHex.Text = ColorPaletteDefinition.ToHex(c);
            colorPreview.BackColor = c;
            hueSlider.Value = hVal;
            colorCanvas.Hue = hVal; colorCanvas.Saturation = sVal; colorCanvas.Lightness = lVal; colorCanvas.Invalidate();
            updatingControls = false;
        }

        private void AddColor()
        {
            var hex = txtHex.Text.Trim();
            if (!hex.StartsWith("#") || (hex.Length != 7)) return;
            Palette.Colors.Add(hex);
            RebuildLeftList();
            SelectIndex(Palette.Colors.Count - 1);
        }

        // Simple HSL<->RGB helpers
        private static void RgbToHsl(Color c, out int h, out int s, out int l)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            double hD = 0, sD, lD = (max + min) / 2.0;
            if (Math.Abs(max - min) < 1e-6) { hD = 0; sD = 0; }
            else
            {
                double d = max - min;
                sD = lD > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                if (max == r) hD = (g - b) / d + (g < b ? 6 : 0);
                else if (max == g) hD = (b - r) / d + 2;
                else hD = (r - g) / d + 4;
                hD /= 6.0;
            }
            h = (int)Math.Round(hD * 359); s = (int)Math.Round(sD * 100); l = (int)Math.Round(lD * 100);
        }

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

        private static int Clamp(int value, int min, int max) => value < min ? min : (value > max ? max : value);
    }
}

// Custom owner-drawn reorderable list for palette colors
namespace my_addin
{
    using System;
    using System.Drawing;
    using System.Windows.Forms;

    internal class ReorderableColorList : ListBox
    {
        public event Action<int> DeleteRequested;
        public event Action<int, int> ReorderRequested; // from, to

        private int dragIndex = -1;
        private int insertionIndex = -1;
        private Point mouseDown;
        private bool dragging = false;
        
        // Cached drawing objects
        private static readonly Pen RedPen = new Pen(Color.Red, 2);
        private static readonly SolidBrush GrayBrush = new SolidBrush(Color.Gray);
        private static readonly Pen DarkGrayPen = new Pen(Color.DarkGray);
        private static readonly Pen InsertionPen = new Pen(Color.DeepSkyBlue, 2);
        private readonly SolidBrush _colorBrush = new SolidBrush(Color.White);

        public ReorderableColorList()
        {
            DrawMode = DrawMode.OwnerDrawFixed;
            ItemHeight = 24;
            BorderStyle = BorderStyle.FixedSingle;
            IntegralHeight = false;
            AllowDrop = true;
        }

        public new void BeginUpdate() => base.BeginUpdate();
        public new void EndUpdate() => base.EndUpdate();

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0 || e.Index >= Items.Count) return;

            string hex = Items[e.Index].ToString();
            Color c = Core.ColorPaletteDefinition.FromHex(hex);
            Rectangle bounds = e.Bounds;

            // Delete icon area
            Rectangle del = new Rectangle(bounds.X + 6, bounds.Y + 5, 12, 12);
            e.Graphics.DrawLine(RedPen, del.Left, del.Top, del.Right, del.Bottom);
            e.Graphics.DrawLine(RedPen, del.Right, del.Top, del.Left, del.Bottom);

            // Handle dots area
            Rectangle grip = new Rectangle(del.Right + 6, bounds.Y + 4, 10, bounds.Height - 8);
            for (int y = grip.Top; y < grip.Bottom; y += 4)
                for (int x = grip.Left; x < grip.Right; x += 4)
                    e.Graphics.FillEllipse(GrayBrush, x, y, 2, 2);

            // Swatch area fills remainder
            Rectangle swatch = new Rectangle(grip.Right + 6, bounds.Y + 2, bounds.Right - (grip.Right + 12), bounds.Height - 4);
            _colorBrush.Color = c;
            e.Graphics.FillRectangle(_colorBrush, swatch);
            e.Graphics.DrawRectangle(DarkGrayPen, swatch);

            // Hex text with contrast color
            double luminance = 0.2126 * c.R + 0.7152 * c.G + 0.0722 * c.B;
            Color textColor = luminance < 140 ? Color.White : Color.Black;
            TextRenderer.DrawText(e.Graphics, hex, SystemFonts.DefaultFont, new Point(swatch.X + 6, swatch.Y + 5), textColor);

            // Focus rectangle
            if ((e.State & DrawItemState.Focus) == DrawItemState.Focus)
                e.DrawFocusRectangle();

            // Insertion line
            if (dragging && insertionIndex >= 0)
            {
                int y = e.Bounds.Top - 1;
                if (e.Index == insertionIndex) 
                { 
                    e.Graphics.DrawLine(InsertionPen, bounds.Left + 2, y, bounds.Right - 2, y); 
                }
                if (insertionIndex == Items.Count && e.Index == Items.Count - 1) 
                { 
                    e.Graphics.DrawLine(InsertionPen, bounds.Left + 2, bounds.Bottom - 1, bounds.Right - 2, bounds.Bottom - 1); 
                }
            }
        }
        
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _colorBrush?.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            dragIndex = IndexFromPoint(e.Location);
            mouseDown = e.Location;
            dragging = false;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (!dragging && dragIndex >= 0)
            {
                // Check if clicked delete icon
                Rectangle itemBounds = GetItemRectangle(dragIndex);
                Rectangle del = new Rectangle(itemBounds.X + 6, itemBounds.Y + 5, 12, 12);
                if (del.Contains(e.Location)) DeleteRequested?.Invoke(dragIndex);
            }
            dragging = false; dragIndex = -1; insertionIndex = -1; Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.Button == MouseButtons.Left && dragIndex >= 0)
            {
                if (!dragging && (Math.Abs(e.X - mouseDown.X) + Math.Abs(e.Y - mouseDown.Y) > 4))
                {
                    dragging = true;
                    DoDragDrop(Items[dragIndex], DragDropEffects.Move);
                }
            }
        }

        protected override void OnDragOver(DragEventArgs drgevent)
        {
            base.OnDragOver(drgevent);
            drgevent.Effect = DragDropEffects.Move;
            Point pt = PointToClient(new Point(drgevent.X, drgevent.Y));
            int idx = IndexFromPoint(pt);
            if (idx < 0) idx = Items.Count; // after last
            insertionIndex = idx;
            Invalidate();
        }

        protected override void OnDragDrop(DragEventArgs drgevent)
        {
            base.OnDragDrop(drgevent);
            if (!dragging || dragIndex < 0) return;
            int to = insertionIndex;
            if (to > dragIndex) to--; // adjusting for removal
            if (to < 0) to = 0; if (to > Items.Count - 1) to = Items.Count - 1;
            ReorderRequested?.Invoke(dragIndex, to);
            dragging = false; dragIndex = -1; insertionIndex = -1; Invalidate();
        }
    }
}

