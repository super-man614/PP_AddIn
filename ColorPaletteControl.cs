using System;
using System.Drawing;
using System.Windows.Forms;
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
            int headerHeight = this.Parent?.Height ?? 0;
            this.Size = new Size(100, headerHeight - 10);

            // Edit button at top-left
            btnEdit = new Button { Text = "Edit", Location = new Point(6, 6), Size = new Size(50, 22), FlatStyle = FlatStyle.Standard };
            btnEdit.Click += (s, e) => EditPalette();
            Controls.Add(btnEdit);

            // Palette panel for color rows
            palettePanel = new FlowLayoutPanel
            {
                Location = new Point(6, 34),
                Size = new Size(100, 600),
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(0),
                BackColor = Color.White
            };
            Controls.Add(palettePanel);

            // Load palette
            palette = Core.ColorPaletteStorage.LoadOrDefault();
            BuildPaletteRows();

            this.ResumeLayout(false);
        }

        private void BuildPaletteRows()
        {
            palettePanel.SuspendLayout();
            palettePanel.Controls.Clear();

            // First special row: Clear Fill, Clear Outline (square + X)
            var firstRow = new Panel { Width = 120, Height = 22, Margin = new Padding(0, 0, 0, 6) };
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

            firstRow.Controls.Add(noFill);
            firstRow.Controls.Add(noOutline);
            palettePanel.Controls.Add(firstRow);

            foreach (var hex in palette.Colors)
            {
                var c = Core.ColorPaletteDefinition.FromHex(hex);
                var row = new Panel { Width = 120, Height = 22, Margin = new Padding(0, 2, 0, 2) };

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
    }
}
