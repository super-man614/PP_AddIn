using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace my_addin
{
    public partial class ThisAddIn
    {
        private PowerPoint.EApplication_WindowSelectionChangeEventHandler selectionChangedHandler;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("=== PowerPoint Add-in Started ===");
            System.Diagnostics.Debug.WriteLine("Add-in is initializing...");
            System.Diagnostics.Debug.WriteLine("Ribbon should be created automatically by VSTO framework");
            
            // Auto-create taskpanes on startup
            try
            {
                System.Diagnostics.Debug.WriteLine("Auto-creating PowerPoint Tools taskpane...");
                var taskPane = new CustomTaskPane();
                Ribbon.TaskPaneInstance = taskPane; // Set shared instance
                System.Diagnostics.Debug.WriteLine("PowerPoint Tools taskpane created and shown automatically");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error auto-creating PowerPoint Tools taskpane: {ex.Message}");
            }

            try
            {
                System.Diagnostics.Debug.WriteLine("Auto-creating Color Palette taskpane...");
                var colorPaletteTaskPane = new ColorPaletteTaskPane();
                Ribbon.ColorPaletteInstance = colorPaletteTaskPane; // Set shared instance
                System.Diagnostics.Debug.WriteLine("Color Palette taskpane created and shown automatically");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error auto-creating Color Palette: {ex.Message}");
            }

            try
            {
                selectionChangedHandler = new PowerPoint.EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
                this.Application.WindowSelectionChange += selectionChangedHandler;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to attach selection change handler: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("=== PowerPoint Add-in Shutdown ===");
            try { this.Application.WindowSelectionChange -= selectionChangedHandler; } catch { }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            // No-op: placeholder for future feedback
        }

        // Paste matrix contents from clipboard when Ctrl+V pressed while matrix cells are selected
        public bool TryPasteIntoMatrix()
        {
            try
            {
                var app = this.Application;
                var wnd = app?.ActiveWindow; if (wnd == null) return false;
                var sel = wnd.Selection; if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return false;

                if (!Clipboard.ContainsText()) return false;
                string text = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(text)) return false;

                // Parse as TSV/CSV
                string[] rows = text.Replace("\r\n", "\n").Replace("\r", "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var data = new System.Collections.Generic.List<string[]>();
                foreach (var r in rows)
                {
                    var parts = r.Split('\t');
                    data.Add(parts);
                }

                PowerPoint.ShapeRange range = sel.ShapeRange;
                int n = range.Count;
                // Collect selected matrix cells and sort by (row,col)
                var cells = new System.Collections.Generic.List<(int row, int col, PowerPoint.Shape shape)>();
                foreach (PowerPoint.Shape s in range)
                {
                    try
                    {
                        if (s.Tags["MATRIX"] == "1")
                        {
                            int r = int.Parse(s.Tags["MATRIX_ROW"]);
                            int c = int.Parse(s.Tags["MATRIX_COL"]);
                            cells.Add((r, c, s));
                        }
                    }
                    catch { }
                }
                if (cells.Count == 0) return false;
                cells.Sort((a, b) => a.row != b.row ? a.row.CompareTo(b.row) : a.col.CompareTo(b.col));

                // Fill row-major until out of data or cells
                int dr = data.Count;
                int dc = 0; foreach (var arr in data) if (arr.Length > dc) dc = arr.Length;

                int idx = 0;
                for (int r = 0; r < dr; r++)
                {
                    for (int c = 0; c < data[r].Length; c++)
                    {
                        if (idx >= cells.Count) break;
                        var cell = cells[idx++];
                        string val = data[r][c];
                        try
                        {
                            cell.shape.TextFrame.TextRange.Text = val;
                        }
                        catch { }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PasteIntoMatrix failed: {ex.Message}");
                return false;
            }
        }
        
        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}