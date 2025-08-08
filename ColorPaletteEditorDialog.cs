using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using my_addin.Core;

namespace my_addin
{
    public class ColorPaletteEditorDialog : Form
    {
        // Left list: custom panel with rows (delete, drag, swatch)
        private FlowLayoutPanel leftList;
        // Right: picker and numeric inputs
        private Panel rightPanel;
        private Panel colorPreview;
        private TextBox txtHex;
        private TextBox txtH, txtS, txtL, txtR, txtG, txtB;
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

            leftList = new FlowLayoutPanel { Location = new Point(10, 10), Size = new Size(240, 430), AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(0) };
            Controls.Add(leftList);

            rightPanel = new Panel { Location = new Point(260, 10), Size = new Size(340, 430), BorderStyle = BorderStyle.None };
            Controls.Add(rightPanel);

            // Top row: preview swatch + hex entry (color pick area at top)
            colorPreview = new Panel { Location = new Point(0, 0), Size = new Size(38, 24), BorderStyle = BorderStyle.FixedSingle };
            txtHex = new TextBox { Location = new Point(46, 0), Width = 100 };
            txtHex.TextChanged += (s, e) => { if (!updatingControls) UpdateFromHex(); };
            rightPanel.Controls.Add(colorPreview);
            rightPanel.Controls.Add(txtHex);

            // HSL and RGB as three rows of text inputs: H/R, S/G, L/B
            var hr = new TableLayoutPanel { Location = new Point(0, 36), Size = new Size(330, 28), ColumnCount = 4, RowCount = 1 };
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            hr.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            txtH = new TextBox { Width = 120 }; txtR = new TextBox { Width = 120 };
            txtH.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtH.Text, out var v)) { hVal = Clamp(v, 0, 359); UpdateFromHslFields(); } };
            txtR.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtR.Text, out var v)) { rVal = Clamp(v, 0, 255); UpdateFromRgbFields(); } };
            hr.Controls.Add(new Label { Text = "H", AutoSize = true }, 0, 0); hr.Controls.Add(txtH, 1, 0);
            hr.Controls.Add(new Label { Text = "R", AutoSize = true }, 2, 0); hr.Controls.Add(txtR, 3, 0);
            rightPanel.Controls.Add(hr);

            var sg = new TableLayoutPanel { Location = new Point(0, 68), Size = new Size(330, 28), ColumnCount = 4, RowCount = 1 };
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            sg.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            txtS = new TextBox { Width = 120 }; txtG = new TextBox { Width = 120 };
            txtS.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtS.Text, out var v)) { sVal = Clamp(v, 0, 100); UpdateFromHslFields(); } };
            txtG.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtG.Text, out var v)) { gVal = Clamp(v, 0, 255); UpdateFromRgbFields(); } };
            sg.Controls.Add(new Label { Text = "S", AutoSize = true }, 0, 0); sg.Controls.Add(txtS, 1, 0);
            sg.Controls.Add(new Label { Text = "G", AutoSize = true }, 2, 0); sg.Controls.Add(txtG, 3, 0);
            rightPanel.Controls.Add(sg);

            var lb = new TableLayoutPanel { Location = new Point(0, 100), Size = new Size(330, 28), ColumnCount = 4, RowCount = 1 };
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            lb.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            txtL = new TextBox { Width = 120 }; txtB = new TextBox { Width = 120 };
            txtL.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtL.Text, out var v)) { lVal = Clamp(v, 0, 100); UpdateFromHslFields(); } };
            txtB.TextChanged += (s, e) => { if (!updatingControls && int.TryParse(txtB.Text, out var v)) { bVal = Clamp(v, 0, 255); UpdateFromRgbFields(); } };
            lb.Controls.Add(new Label { Text = "L", AutoSize = true }, 0, 0); lb.Controls.Add(txtL, 1, 0);
            lb.Controls.Add(new Label { Text = "B", AutoSize = true }, 2, 0); lb.Controls.Add(txtB, 3, 0);
            rightPanel.Controls.Add(lb);

            // Add + reset template
            btnAdd = new Button { Text = "Add", Location = new Point(0, 140), Width = 80 };
            btnAdd.Click += (s, e) => AddColor();
            btnResetTemplate = new Button { Text = "Add from template", Location = new Point(90, 140), Width = 140 };
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

        private void RebuildLeftList()
        {
            leftList.SuspendLayout();
            leftList.Controls.Clear();
            for (int i = 0; i < Palette.Colors.Count; i++)
            {
                int index = i;
                var row = new Panel { Width = 220, Height = 24, Margin = new Padding(0, 2, 0, 2), Tag = index, AllowDrop = true };

                // Delete button (X)
                var btnDel = new Label { Text = "X", ForeColor = Color.Red, TextAlign = ContentAlignment.MiddleCenter, Width = 20, Height = 20, Location = new Point(2, 2), Cursor = Cursors.Hand };
                btnDel.Click += (s, e) => { Palette.Colors.RemoveAt(index); RebuildLeftList(); if (Palette.Colors.Count > 0) SelectIndex(Math.Min(index, Palette.Colors.Count - 1)); };

                // Drag handle
                var handle = new Label { Text = "⋮⋮", ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleCenter, Width = 20, Height = 20, Location = new Point(24, 2), Cursor = Cursors.SizeAll };
                handle.MouseDown += (s, e) => row.DoDragDrop(index, DragDropEffects.Move);
                row.DragEnter += (s, e) => { if (e.Data.GetDataPresent(typeof(int))) e.Effect = DragDropEffects.Move; };
                row.DragDrop += (s, e) => { var from = (int)e.Data.GetData(typeof(int)); var to = index; if (from != to) { var c = Palette.Colors[from]; Palette.Colors.RemoveAt(from); Palette.Colors.Insert(to, c); RebuildLeftList(); SelectIndex(to); } };

                // Swatch with hex text
                var hex = Palette.Colors[i];
                var swatch = new Panel { Width = 160, Height = 20, Location = new Point(46, 2), BackColor = ColorPaletteDefinition.FromHex(hex), BorderStyle = BorderStyle.FixedSingle, Cursor = Cursors.Hand };
                var lbl = new Label { Text = hex, AutoSize = false, Width = 150, Height = 18, Location = new Point(4, 1), ForeColor = Color.White };
                swatch.Controls.Add(lbl);
                swatch.Click += (s, e) => SelectIndex(index);

                if (index == selectedIndex) row.BackColor = Color.FromArgb(230, 230, 230);

                row.Controls.AddRange(new Control[] { btnDel, handle, swatch });
                leftList.Controls.Add(row);
            }
            leftList.ResumeLayout(true);
        }

        private void SelectIndex(int index)
        {
            if (index < 0 || index >= Palette.Colors.Count) return;
            selectedIndex = index;
            RebuildLeftList();
            var c = ColorPaletteDefinition.FromHex(Palette.Colors[index]);
            updatingControls = true;
            colorPreview.BackColor = c;
            txtHex.Text = ColorPaletteDefinition.ToHex(c);
            RgbToHsl(c, out hVal, out sVal, out lVal);
            rVal = c.R; gVal = c.G; bVal = c.B;
            txtH.Text = hVal.ToString(); txtS.Text = sVal.ToString(); txtL.Text = lVal.ToString();
            txtR.Text = rVal.ToString(); txtG.Text = gVal.ToString(); txtB.Text = bVal.ToString();
            updatingControls = false;
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
            txtHex.Text = ColorPaletteDefinition.ToHex(c);
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

