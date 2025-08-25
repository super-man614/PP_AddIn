using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using my_addin.Core;
using System.Reflection; // For Missing.Value

namespace my_addin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private static CustomTaskPane _taskPaneInstance;
        private static ColorPaletteTaskPane _colorPaletteInstance;
        private static Ribbon _current;

        public static CustomTaskPane TaskPaneInstance
        {
            get { return _taskPaneInstance; }
            set { _taskPaneInstance = value; }
        }

        public static ColorPaletteTaskPane ColorPaletteInstance
        {
            get { return _colorPaletteInstance; }
            set { _colorPaletteInstance = value; }
        }

        public static Ribbon Current
        {
            get { return _current; }
            set { _current = value; }
        }

        public Ribbon()
        {
            _current = this;
            System.Diagnostics.Debug.WriteLine("Ribbon constructor called - Ribbon instance created");
        }

        
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("GetCustomUI called - loading Ribbon.xml embedded resource");
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                // Debug resource names to find the correct one
                System.Diagnostics.Debug.WriteLine("Available resources:");
                foreach (var resourceName in asm.GetManifestResourceNames())
                {
                    System.Diagnostics.Debug.WriteLine($"Resource: {resourceName}");
                }
                using (var stream = asm.GetManifestResourceStream("my_addin.Ribbon.xml"))
                using (var reader = new System.IO.StreamReader(stream ?? throw new InvalidOperationException("Ribbon.xml not found as embedded resource")))
                {
                    string xmlContent = reader.ReadToEnd();
                    // Define the output directory and file path
                    string directory = @"C:\RibbnTest";
                    string filePath = Path.Combine(directory, "Ribbon.xml");

                    // Ensure the directory exists
                    if(!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                    // Write the XML content to the file
                    File.WriteAllText(filePath, xmlContent);

                    // Log success
                    System.Diagnostics.Debug.WriteLine("Successfully loaded and saved Ribbon.xml to " + filePath);

                    return xmlContent;
                }
            }
            catch (Exception ex)
            {
                // As a fallback, return a minimal ribbon with a single Test button
                System.Diagnostics.Debug.WriteLine($"Failed to load Ribbon.xml: {ex.Message}");
                return @"<?xml version='1.0' encoding='UTF-8'?>
    <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
      <ribbon>
        <tabs>
          <tab id='PowerPointToolsTab' label='PowerPoint Tools'>
            <group id='FallbackGroup' label='Fallback'>
              <button id='TestRibbonButton' label='Test Ribbon' size='large' onAction='TestRibbon_Click' imageMso='HappyFace' screentip='Fallback ribbon loaded'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
            }
        }


        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine("Ribbon_Load called - ribbon is being initialized");
            ribbon = ribbonUI;
            System.Diagnostics.Debug.WriteLine("Ribbon UI stored successfully");
        }

        public void ToggleTaskPane_Click(Office.IRibbonControl control)
        {
            try
            {
                if (_taskPaneInstance == null || _taskPaneInstance.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Creating new task pane instance...");
                    _taskPaneInstance = new CustomTaskPane();
                }

                // Ensure our main Tools pane docks to the right
                try
                {
                    _taskPaneInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                }
                catch { }

                // Make sure Color Palette exists and is also on the right
                if (_colorPaletteInstance == null || _colorPaletteInstance.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Ensuring Color Palette task pane exists...");
                    _colorPaletteInstance = new ColorPaletteTaskPane();
                }
                try
                {
                    _colorPaletteInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                }
                catch { }

                // Show both panes
                _taskPaneInstance.Visible = true;
                _colorPaletteInstance.Visible = true;

                // Ensure Color Palette sits left-most among right-docked panes
                try { Core.PaneOrdering.EnsureColorPaletteLeftMost(_colorPaletteInstance); } catch { }
                try { Core.PaneManager.OnPaneVisibilityChanged(); } catch { }

                // Update the toggle button state
                if (ribbon != null)
                    ribbon.Invalidate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ToggleTaskPane_Click: {ex.Message}");
            }
        }

        public void TestRibbon_Click(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("TestRibbon_Click executed successfully");
                // Removed unnecessary popup message
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TestRibbon_Click: {ex.Message}");
            }
        }

        public void ColorPaletteToggle_Click(Office.IRibbonControl control)
        {
            try
            {
                if (_colorPaletteInstance == null || _colorPaletteInstance.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Creating new Color Palette task pane...");
                    _colorPaletteInstance = new ColorPaletteTaskPane();
                }

                // Force docking to the right
                try
                {
                    _colorPaletteInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                }
                catch { }

                // If our Tools pane exists, ensure it's on the right as well
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    try { _taskPaneInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }
                }

                // Toggle visibility
                _colorPaletteInstance.Toggle();

                // If both are visible, ensure Color Palette sits left-most
                if (_colorPaletteInstance.Visible && _taskPaneInstance != null && !_taskPaneInstance.IsDisposed && _taskPaneInstance.Visible)
                {
                    try { Core.PaneOrdering.EnsureColorPaletteLeftMost(_colorPaletteInstance); } catch { }
                }
                try { Core.PaneManager.OnPaneVisibilityChanged(); } catch { }

                // Update the toggle button state
                if (ribbon != null)
                    ribbon.Invalidate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ColorPaletteToggle_Click: {ex.Message}");
            }
        }

        public bool GetColorPalettePressed(Office.IRibbonControl control)
        {
            return _colorPaletteInstance != null && !_colorPaletteInstance.IsDisposed && _colorPaletteInstance.Visible;
        }

        // Preset callbacks
        public void Preset1_Apply(Office.IRibbonControl control) { ShapePresetStorage.ApplyPreset(1); }
        public void Preset1_Save(Office.IRibbonControl control) { ShapePresetStorage.SavePreset(1); }
        public void Preset1_Clear(Office.IRibbonControl control) { ShapePresetStorage.ClearPreset(1); }
        public void Preset2_Apply(Office.IRibbonControl control) { ShapePresetStorage.ApplyPreset(2); }
        public void Preset2_Save(Office.IRibbonControl control) { ShapePresetStorage.SavePreset(2); }
        public void Preset2_Clear(Office.IRibbonControl control) { ShapePresetStorage.ClearPreset(2); }
        public void Preset3_Apply(Office.IRibbonControl control) { ShapePresetStorage.ApplyPreset(3); }
        public void Preset3_Save(Office.IRibbonControl control) { ShapePresetStorage.SavePreset(3); }
        public void Preset3_Clear(Office.IRibbonControl control) { ShapePresetStorage.ClearPreset(3); }

        // Format Tools callbacks
        public void UniformSizes_Click(Office.IRibbonControl control) { FormatTools.MatchSize(); }
        public void MatchWidth_Click(Office.IRibbonControl control) { FormatTools.MatchWidth(); }
        public void MatchHeight_Click(Office.IRibbonControl control) { FormatTools.MatchHeight(); }
        public void MatchSize_Click(Office.IRibbonControl control) { FormatTools.MatchSize(); }
        
        public void MatchColors_Click(Office.IRibbonControl control) { FormatTools.MatchFill(); }
        public void MatchFill_Click(Office.IRibbonControl control) { FormatTools.MatchFill(); }
        public void MatchFont_Click(Office.IRibbonControl control) { FormatTools.MatchFontColor(); }
        public void MatchOutline_Click(Office.IRibbonControl control) { FormatTools.MatchOutline(); }
        
        public void AlignFonts_Click(Office.IRibbonControl control) { FormatTools.MatchFontSize(); }
        public void MatchFontSize_Click(Office.IRibbonControl control) { FormatTools.MatchFontSize(); }
        public void MatchFontFamily_Click(Office.IRibbonControl control) { FormatTools.MatchFontFamily(); }
        public void MatchFontColor_Click(Office.IRibbonControl control) { FormatTools.MatchFontColor(); }

        // Wrap toggle
        public void WrapToggle_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app?.ActiveWindow?.Selection;
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange == null || selection.ShapeRange.Count < 1)
                    return;

                // Determine current wrap setting from first shape with text
                Office.MsoTriState? current = null;
                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    var sh = selection.ShapeRange[i];
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        current = sh.TextFrame2.WordWrap;
                        break;
                    }
                }
                var newVal = (current == Office.MsoTriState.msoTrue) ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;

                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    var sh = selection.ShapeRange[i];
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        sh.TextFrame2.WordWrap = newVal;
                    }
                }
            }
            catch { }
        }

        // File operation callbacks
        public void New_Click(Office.IRibbonControl control)
    {
        try
        {
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteNewFile();
            }
            else
            {
                var app = Globals.ThisAddIn.Application;
                app.Presentations.Add(Office.MsoTriState.msoTrue);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in New_Click: {ex.Message}");
        }
    }

        public void Open_Click(Office.IRibbonControl control)
    {
        try
        {
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteOpenFile();
            }
            else
            {
                var app = Globals.ThisAddIn.Application;
                    //app.Presentations.Open(Missing.Value, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
                    app.Presentations.Open(Missing.Value.ToString(), Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Open_Click: {ex.Message}");
        }
    }

        public void Save_Click(Office.IRibbonControl control)
    {
        try
        {
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteSaveFile();
            }
            else
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActivePresentation.Save();
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Save_Click: {ex.Message}");
        }
    }

        public void SaveAs_Click(Office.IRibbonControl control)
    {
        try
        {
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteSaveAsFile();
            }
            else
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActivePresentation.Save();
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in SaveAs_Click: {ex.Message}");
        }
    }

        public void Print_Click(Office.IRibbonControl control)
    {
        try
        {
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecutePrint();
            }
            else
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActivePresentation.PrintOut();
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Print_Click: {ex.Message}");
        }
    }

        public void Share_Click(Office.IRibbonControl control)
    {
        try
        {
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteShare();
            }
            else
            {
                // Fallback to message if task pane is not available
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Share Presentation", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Share_Click: {ex.Message}");
        }
    }

        // Wizards callbacks
        public void Agenda_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteAgendaWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Agenda Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Agenda_Click: {ex.Message}");
            }
        }

        public void Master_Click(Office.IRibbonControl control)
    {
        try
        {
            // Delegate to task pane implementation
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteMasterWizard();
            }
            else
            {
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Master Wizard",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Master_Click: {ex.Message}");
        }
    }

        public void Element_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteElementWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Element Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Element_Click: {ex.Message}");
            }
        }

        public void TextWizard_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteTextWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Text Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TextWizard_Click: {ex.Message}");
            }
        }

        public void Format_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteFormatWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Format Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Format_Click: {ex.Message}");
            }
        }

        public void Map_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteMapWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Map Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Map_Click: {ex.Message}");
            }
        }

        // Smart Elements callbacks
        public void Chart_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteChartWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Chart Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Chart_Click: {ex.Message}");
            }
        }

        public void Diagram_Click(Office.IRibbonControl control)
        {
            try
            {
                // Delegate to task pane implementation
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    _taskPaneInstance.ExecuteDiagramWizard();
                }
                else
                {
                    MessageBox.Show("Task pane is not available. Please open the task pane first.", "Diagram Wizard",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Diagram_Click: {ex.Message}");
            }
        }

        public void Table_Click(Office.IRibbonControl control)
    {
        try
        {
            // Delegate to task pane implementation
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteTableWizard();
            }
            else
            {
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Table Wizard",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Table_Click: {ex.Message}");
        }
    }

        public void MatrixTable_Click(Office.IRibbonControl control)
    {
        try
        {
            // Delegate to task pane implementation
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteMatrixTableWizard();
            }
            else
            {
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Matrix Table Wizard",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in MatrixTable_Click: {ex.Message}");
        }
    }

        public void ExcelPaste_Click(Office.IRibbonControl control)
        {
            try
            {
                // Get the active slide
                var app = Globals.ThisAddIn.Application;
                var slide = GetActiveSlideOrNull();
                if (slide == null)
                {
                    MessageBox.Show("Please select a slide to paste Excel data.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Find the Matrix table (table with cells as objects) in the main layout
                PowerPoint.Shape matrixTableShape = null;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        // You may want to add more checks here to identify your specific Matrix table
                        // For example, by name, tag, or size
                        if (shape.Name.Contains("Matrix") || shape.Tags["MatrixTable"] == "1")
                        {
                            matrixTableShape = shape;
                            break;
                        }
                    }
                }

                if (matrixTableShape == null)
                {
                    MessageBox.Show("No Matrix table found on the current slide.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Get clipboard data (CSV or Text)
                string clipboardText = "";
                if (System.Windows.Forms.Clipboard.ContainsData(DataFormats.CommaSeparatedValue))
                {
                    clipboardText = System.Windows.Forms.Clipboard.GetData(DataFormats.CommaSeparatedValue)?.ToString();
                }
                else if (System.Windows.Forms.Clipboard.ContainsData(DataFormats.Text))
                {
                    clipboardText = System.Windows.Forms.Clipboard.GetText();
                }

                if (string.IsNullOrWhiteSpace(clipboardText))
                {
                    MessageBox.Show("No valid data found in clipboard.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Parse clipboard text into rows and columns
                var rows = clipboardText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                var data = rows.Select(r => r.Split('\t', ',', ';')).ToArray();

                var table = matrixTableShape.Table;
                int rowCount = Math.Min(table.Rows.Count, data.Length);

                for (int i = 1; i <= rowCount; i++)
                {
                    var rowData = data[i - 1];
                    int colCount = Math.Min(table.Columns.Count, rowData.Length);
                    for (int j = 1; j <= colCount; j++)
                    {
                        // Place the correct cell value into the corresponding cell
                        table.Cell(i, j).Shape.TextFrame.TextRange.Text = rowData[j - 1];
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ExcelPaste_Click: {ex.Message}");
                MessageBox.Show($"Error executing Excel Paste: {ex.Message}", "Excel Paste Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Executes Excel Paste functionality directly from Ribbon without requiring task pane
        /// </summary>
        private void ExecuteExcelPasteDirectly()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var slide = GetActiveSlideOrNull();
                if (slide == null)
                {
                    MessageBox.Show("Please select a slide to paste Excel data.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if PowerPoint can access clipboard
                if (!System.Windows.Forms.Clipboard.ContainsData(DataFormats.CommaSeparatedValue) && 
                    !System.Windows.Forms.Clipboard.ContainsData(DataFormats.Text))
                {
                    MessageBox.Show("No Excel data found in clipboard. Please copy data from Excel first.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Get clipboard data
                string clipboardText = "";
                if (System.Windows.Forms.Clipboard.ContainsData(DataFormats.CommaSeparatedValue))
                {
                    clipboardText = System.Windows.Forms.Clipboard.GetData(DataFormats.CommaSeparatedValue).ToString();
                }
                else if (System.Windows.Forms.Clipboard.ContainsData(DataFormats.Text))
                {
                    clipboardText = System.Windows.Forms.Clipboard.GetText();
                }

                if (string.IsNullOrWhiteSpace(clipboardText))
                {
                    MessageBox.Show("No valid data found in clipboard.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Parse the clipboard data
                var rows = ParseClipboardData(clipboardText);
                if (rows == null || rows.Count == 0)
                {
                    MessageBox.Show("Could not parse clipboard data. Please ensure data is properly formatted.", "Excel Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Create table on the slide
                CreateTableFromData(slide, rows);

                MessageBox.Show($"Successfully created table with {rows.Count} rows and {rows.Max(row => row.Count)} columns from Excel data.", 
                    "Excel Paste Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to execute Excel Paste: {ex.Message}");
            }
        }

        /// <summary>
        /// Parses clipboard data into rows and columns
        /// </summary>
        private List<List<string>> ParseClipboardData(string clipboardText)
        {
            try
            {
                var rows = new List<List<string>>();
                var lines = clipboardText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var columns = new List<string>();
                    var cells = line.Split('\t'); // Tab-separated values

                    foreach (var cell in cells)
                    {
                        // Remove quotes if present
                        var cleanCell = cell.Trim('"', ' ');
                        columns.Add(cleanCell);
                    }

                    if (columns.Count > 0)
                    {
                        rows.Add(columns);
                    }
                }

                return rows;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing clipboard data: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Creates a table on the slide from the parsed data
        /// </summary>
        private void CreateTableFromData(PowerPoint.Slide slide, List<List<string>> rows)
        {
            try
            {
                if (rows == null || rows.Count == 0) return;

                // Determine table dimensions
                int rowCount = rows.Count;
                int colCount = rows.Max(row => row.Count);

                if (colCount == 0) return;

                // Table dimensions (matching existing implementation)
                float cellWidth = 80.0f;
                float cellHeight = 25.0f;
                float tableWidth = colCount * cellWidth;
                float tableHeight = rowCount * cellHeight;
                
                // Center the table on the slide (using CustomLayout like existing implementation)
                float left = (slide.CustomLayout.Width - tableWidth) / 2;
                float top = (slide.CustomLayout.Height - tableHeight) / 2;

                // Create the table
                PowerPoint.Shape tableShape = slide.Shapes.AddTable(
                    rowCount,
                    colCount,
                    left,
                    top,
                    tableWidth,
                    tableHeight);

                // Name the table like existing implementation
                tableShape.Name = "ExcelPastedTable";

                // Get the table object
                PowerPoint.Table table = tableShape.Table;

                // Populate the table with data
                for (int row = 0; row < rowCount; row++)
                {
                    var rowData = rows[row];
                    for (int col = 0; col < colCount; col++)
                    {
                        string cellText = (col < rowData.Count) ? rowData[col] : "";
                        
                        try
                        {
                            var cell = table.Cell(row + 1, col + 1);
                            if (cell != null && cell.Shape != null)
                            {
                                // Set cell text
                                cell.Shape.TextFrame.TextRange.Text = cellText;
                                
                                // Apply basic formatting (matching existing implementation)
                                cell.Shape.TextFrame.TextRange.Font.Size = 10;
                                cell.Shape.TextFrame.TextRange.Font.Name = "Calibri";
                                
                                // Add borders (matching existing implementation)
                                cell.Shape.Line.Visible = Office.MsoTriState.msoTrue;
                                cell.Shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error setting cell ({row},{col}): {ex.Message}");
                        }
                    }
                }

                // Apply table styling (matching existing implementation)
                try
                {
                    // Header row styling
                    for (int col = 1; col <= colCount; col++)
                    {
                        var headerCell = table.Cell(1, col);
                        if (headerCell?.Shape != null)
                        {
                            headerCell.Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                            headerCell.Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error applying table styling: {ex.Message}");
                }

                // Select the table
                tableShape.Select();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create table from data: {ex.Message}");
            }
        }

        public void StickyNote_Click(Office.IRibbonControl control)
    {
        try
        {
            // Delegate to task pane implementation
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteStickyNoteWizard();
            }
            else
            {
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Sticky Note Wizard",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in StickyNote_Click: {ex.Message}");
        }
    }

        public void Citation_Click(Office.IRibbonControl control)
    {
        try
        {
            // Delegate to task pane implementation
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteCitationWizard();
            }
            else
            {
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Citation Wizard",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in Citation_Click: {ex.Message}");
        }
    }

        public void StandardObjects_Click(Office.IRibbonControl control)
    {
        try
        {
            // Delegate to task pane implementation
            if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
            {
                _taskPaneInstance.ExecuteStandardObjectsWizard();
            }
            else
            {
                MessageBox.Show("Task pane is not available. Please open the task pane first.", "Standard Objects Wizard",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error in StandardObjects_Click: {ex.Message}");
        }
    }

        public void Matrix_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    // Get user input for rows and columns
                    var matrixDialog = new MatrixTableDialog();
                    if (matrixDialog.ShowDialog() == DialogResult.OK)
                    {
                        int rows = matrixDialog.Rows;
                        int columns = matrixDialog.Columns;
                        bool hasHeader = matrixDialog.HasHeader;
                        
                        var targetSlide = GetActiveSlideOrNull();
                        if (targetSlide != null)
                        {
                            CreateCustomMatrix(targetSlide, rows, columns);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating matrix table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private PowerPoint.Slide GetActiveSlideOrNull()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow != null && app.ActiveWindow.Selection != null)
                {
                    return app.ActiveWindow.View.Slide;
                }
            }
            catch
            {
                // Ignore errors
            }
            return null;
        }

        private void CreateCustomMatrix(PowerPoint.Slide slide, int rows, int columns)
        {
            try
            {
                // Slide size
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;

                // Define consistent spacing between cells
                float cellSpacing = 8f;

                // Grid size with margin
                float availableWidth = slideWidth * 0.85f;
                float availableHeight = slideHeight * 0.75f;

                // Total spacing needed
                float totalHorizontalSpacing = (columns - 1) * cellSpacing;
                float totalVerticalSpacing = (rows - 1) * cellSpacing;

                // Available space for actual cells
                float cellAreaWidth = availableWidth - totalHorizontalSpacing;
                float cellAreaHeight = availableHeight - totalVerticalSpacing;

                // Cell size
                float maxCellWidth = cellAreaWidth / columns;
                float maxCellHeight = cellAreaHeight / rows;
                float cellSize = Math.Min(maxCellWidth, maxCellHeight);
                cellSize = Math.Max(cellSize, 30f); // Enforce minimum readable size

                // Actual grid size
                float actualGridWidth = (columns * cellSize) + totalHorizontalSpacing;
                float actualGridHeight = (rows * cellSize) + totalVerticalSpacing;

                // Grid start position
                float startLeft = (slideWidth - actualGridWidth) / 2;
                float startTop = (slideHeight - actualGridHeight) / 2;

                // Track all cells in row-major order for later selection
                List<PowerPoint.Shape> allCells = new List<PowerPoint.Shape>();

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < columns; col++)
                    {
                        float left = startLeft + (col * (cellSize + cellSpacing));
                        float top = startTop + (row * (cellSize + cellSpacing));

                        // Create individual shape (text-enabled rectangle)
                        PowerPoint.Shape cell = slide.Shapes.AddShape(
                            Type: Office.MsoAutoShapeType.msoShapeRectangle,
                            Left: left,
                            Top: top,
                            Width: cellSize,
                            Height: cellSize);

                        // Default text
                        cell.TextFrame.TextRange.Text = "XXXX";
                        cell.TextFrame.TextRange.Font.Size = Math.Max(8f, cellSize / 4f);
                        cell.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        cell.TextFrame.TextRange.Font.Name = "Segoe UI";

                        // Align text
                        cell.TextFrame.HorizontalAnchor = Office.MsoHorizontalAnchor.msoAnchorCenter;
                        cell.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                        cell.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                        // Remove text margins for tight layout
                        cell.TextFrame.MarginBottom = 0f;
                        cell.TextFrame.MarginTop = 0f;
                        cell.TextFrame.MarginLeft = 0f;
                        cell.TextFrame.MarginRight = 0f;

                        // Border
                        cell.Line.Visible = Office.MsoTriState.msoTrue;
                        cell.Line.Weight = 1.0f;
                        cell.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                        // Background
                        cell.Fill.Visible = Office.MsoTriState.msoTrue;
                        cell.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                        // Track cell
                        allCells.Add(cell);

                        try
                        {
                            // Tag as matrix cell with row/col for paste handling
                            cell.Tags.Add("MATRIX", "1");
                            cell.Tags.Add("MATRIX_ROW", row.ToString());
                            cell.Tags.Add("MATRIX_COL", col.ToString());
                        }
                        catch { /* ignore tag failures */ }
                    }
                }

                // Select all shapes in row-major order (important for Excel paste)
                string[] shapeNames = allCells.Select(s => s.Name).ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeNames);
                shapeRange.Select();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create matrix table: {ex.Message}");
            }
        }
    }
}