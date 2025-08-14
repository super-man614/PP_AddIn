using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using PowerPointAddIn.Services;

namespace my_addin
{
    public partial class ThisAddIn
    {
        private PowerPoint.EApplication_WindowSelectionChangeEventHandler selectionChangedHandler;
        private IErrorHandlerService _errorHandler;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // Initialize error handler
                _errorHandler = new ErrorHandlerService();
                _errorHandler.LogInfo("=== PowerPoint Add-in Started ===");
                _errorHandler.LogInfo("Add-in is initializing...");
                _errorHandler.LogInfo("Ribbon should be created automatically by VSTO framework");
                
                // Auto-create taskpanes on startup
                CreateTaskPanes();
                
                // Setup selection change handler
                SetupSelectionChangeHandler();
                
                _errorHandler.LogInfo("PowerPoint Add-in initialization completed successfully");
            }
            catch (Exception ex)
            {
                _errorHandler?.HandleError(ex, "Failed to initialize PowerPoint Add-in", "Initialization Error");
            }
        }

        private void CreateTaskPanes()
        {
            try
            {
                _errorHandler.LogInfo("Auto-creating PowerPoint Tools taskpane...");
                var taskPane = new CustomTaskPane();
                Ribbon.TaskPaneInstance = taskPane; // Set shared instance
                _errorHandler.LogInfo("PowerPoint Tools taskpane created and shown automatically");
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "Error auto-creating PowerPoint Tools taskpane");
            }

            try
            {
                _errorHandler.LogInfo("Auto-creating Color Palette taskpane...");
                var colorPaletteTaskPane = new ColorPaletteTaskPane();
                Ribbon.ColorPaletteInstance = colorPaletteTaskPane; // Set shared instance
                _errorHandler.LogInfo("Color Palette taskpane created and shown automatically");
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "Error auto-creating Color Palette");
            }
        }

        private void SetupSelectionChangeHandler()
        {
            try
            {
                selectionChangedHandler = new PowerPoint.EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
                this.Application.WindowSelectionChange += selectionChangedHandler;
                _errorHandler.LogInfo("Selection change handler attached successfully");
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "Failed to attach selection change handler");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                _errorHandler?.LogInfo("=== PowerPoint Add-in Shutdown ===");
                
                // Clean up selection change handler
                if (selectionChangedHandler != null)
                {
                    try 
                    { 
                        this.Application.WindowSelectionChange -= selectionChangedHandler; 
                    } 
                    catch (Exception ex)
                    {
                        _errorHandler?.LogError(ex, "Error removing selection change handler");
                    }
                }

                // Clean up old log files
                _errorHandler?.CleanupOldLogs();
            }
            catch (Exception ex)
            {
                // Use debug output as fallback if error handler fails
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                // Log selection changes for debugging
                if (Sel != null)
                {
                    int count = 0;
                    switch (Sel.Type)
                    {
                        case PowerPoint.PpSelectionType.ppSelectionShapes:
                            count = Sel.ShapeRange?.Count ?? 0;
                            break;
                        case PowerPoint.PpSelectionType.ppSelectionText:
                            count = Sel.TextRange2 != null ? 1 : 0;
                            break;
                        case PowerPoint.PpSelectionType.ppSelectionSlides:
                            count = Sel.SlideRange?.Count ?? 0;
                            break;
                        default:
                            count = 0;
                            break;
                    }

                    _errorHandler?.LogInfo($"Selection changed: Type={Sel.Type}, Count={count}");
                }
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "Error in selection change handler");
            }
        }

        // Paste matrix contents from clipboard when Ctrl+V pressed while matrix cells are selected
        public bool TryPasteIntoMatrix()
        {
            PowerPoint.Application app = null;
            PowerPoint.DocumentWindow wnd = null;
            PowerPoint.Selection sel = null;
            PowerPoint.ShapeRange range = null;

            try
            {
                app = this.Application;
                if (app == null)
                {
                    _errorHandler?.LogWarning("Application object is null");
                    return false;
                }

                wnd = app.ActiveWindow;
                if (wnd == null)
                {
                    _errorHandler?.LogWarning("Active window is null");
                    return false;
                }

                sel = wnd.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _errorHandler?.LogInfo("No shapes selected or selection type is not shapes");
                    return false;
                }

                if (!Clipboard.ContainsText())
                {
                    _errorHandler?.LogInfo("Clipboard does not contain text");
                    return false;
                }

                string text = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(text))
                {
                    _errorHandler?.LogInfo("Clipboard text is empty or whitespace");
                    return false;
                }

                // Parse as TSV/CSV
                string[] rows = text.Replace("\r\n", "\n").Replace("\r", "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var data = new System.Collections.Generic.List<string[]>();
                
                foreach (var r in rows)
                {
                    var parts = r.Split('\t');
                    data.Add(parts);
                }

                range = sel.ShapeRange;
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
                    catch (Exception ex)
                    {
                        _errorHandler?.LogWarning($"Failed to parse matrix cell tags for shape: {ex.Message}");
                    }
                }
                
                if (cells.Count == 0)
                {
                    _errorHandler?.LogInfo("No matrix cells found in selection");
                    return false;
                }
                
                cells.Sort((a, b) => a.row != b.row ? a.row.CompareTo(b.row) : a.col.CompareTo(b.col));

                // Fill row-major until out of data or cells
                int dr = data.Count;
                int dc = 0; 
                foreach (var arr in data) 
                {
                    if (arr.Length > dc) dc = arr.Length;
                }

                int idx = 0;
                int successCount = 0;
                
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
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            _errorHandler?.LogWarning($"Failed to set text for matrix cell ({cell.row},{cell.col}): {ex.Message}");
                        }
                    }
                }

                _errorHandler?.LogInfo($"Matrix paste completed: {successCount}/{cells.Count} cells updated");
                return successCount > 0;
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "PasteIntoMatrix operation failed");
                return false;
            }
            finally
            {
                // Clean up COM objects
                try
                {
                    if (range != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    if (sel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sel);
                    if (wnd != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wnd);
                    if (app != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                }
                catch (Exception ex)
                {
                    _errorHandler?.LogWarning($"Error releasing COM objects: {ex.Message}");
                }
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