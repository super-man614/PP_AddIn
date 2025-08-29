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

        // Element: Object Template
        public void ObjectTemplate_Click(Office.IRibbonControl control)
        {
            try
            {
                var slide = GetActiveSlideOrNull();
                if (slide == null)
                {
                    MessageBox.Show("Please select a slide.", "Object Template", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var dialog = new ObjectTemplateDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.SelectedItem != null)
                    {
                        dialog.SelectedItem.InsertAction(slide);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to insert object template: {ex.Message}", "Object Template", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Element: Slide Template
        public void SlideTemplate_Click(Office.IRibbonControl control)
        {
            try
            {
                var presentation = GetActivePresentationOrNull();
                if (presentation == null)
                {
                    MessageBox.Show("Please open a presentation.", "Slide Template", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var dialog = new SlideTemplateDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.SelectedItem != null)
                    {
                        dialog.SelectedItem.InsertAction(presentation);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to insert slide template: {ex.Message}", "Slide Template", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        public void Matrix_Click_TableVersion(Office.IRibbonControl control)
        {
            try
            {
                using (var dialog = new MatrixTableDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        int rows = dialog.Rows;
                        int columns = dialog.Columns;
                        
                        // Create matrix table in PowerPoint
                        PowerPoint.Application app = Globals.ThisAddIn.Application;
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        
                        // Add a table to the slide
                        PowerPoint.Shape tableShape = slide.Shapes.AddTable(rows, columns, 100, 100, 400, 200);
                        PowerPoint.Table table = tableShape.Table;
                        
                        // Configure the matrix table based on the dialog image pattern
                        for (int row = 1; row <= rows; row++)
                        {
                            for (int col = 1; col <= columns; col++)
                            {
                                PowerPoint.Cell cell = table.Cell(row, col);
                                
                                // Apply blue fill to the top-left 3x3 matrix as shown in the image
                                if (row <= 3 && col <= 3)
                                {
                                    cell.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(70, 130, 180)); // Steel blue
                                    cell.Shape.Fill.Solid();
                                }
                                else
                                {
                                    // Keep other cells with default white background
                                    cell.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                                    cell.Shape.Fill.Solid();
                                }
                                
                                // Add border to all cells
                                cell.Shape.Line.Weight = 1;
                                cell.Shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Matrix_Click: {ex.Message}");
                MessageBox.Show($"Error creating matrix table: {ex.Message}", "Matrix Table Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExcelPaste_Click(Office.IRibbonControl control)
        {
            try
            {
                // Get data from clipboard
                IDataObject dataObj = Clipboard.GetDataObject();
                if (dataObj != null && dataObj.GetDataPresent(DataFormats.Text))
                {
                    string clipboardText = (string)dataObj.GetData(DataFormats.Text);
                    // Split clipboard into rows and then into cells
                    string[] rows = clipboardText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> cellsList = new List<string>();
                    foreach (string row in rows)
                    {
                        if (string.IsNullOrWhiteSpace(row)) continue;
                        string[] cells = row.Split('\t');
                        foreach (string cell in cells)
                        {
                            cellsList.Add(cell);
                        }
                    }

                    // Get selected shapes in PowerPoint
                    PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    if (selection != null && selection.ShapeRange != null && selection.ShapeRange.Count > 0)
                    {
                        int shapeCount = selection.ShapeRange.Count;
                        int cellCount = cellsList.Count;
                        int minCount = Math.Min(shapeCount, cellCount);

                        for (int i = 0; i < minCount; i++)
                        {
                            PowerPoint.Shape currentShape = selection.ShapeRange[i + 1];
                            if (currentShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                currentShape.TextFrame.TextRange.Text = cellsList[i].Trim();
                            }
                        }
                        if (cellCount < shapeCount)
                        {
                            // Optionally clear remaining shapes if there are more shapes than cells
                            for (int i = cellCount; i < shapeCount; i++)
                            {
                                PowerPoint.Shape currentShape = selection.ShapeRange[i + 1];
                                if (currentShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    currentShape.TextFrame.TextRange.Text = "";
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select at least one shape in PowerPoint.", "Excel Paste", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Clipboard doesn't contain text data. Please copy from Excel first.", "Excel Paste", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error pasting data: " + ex.Message, "Excel Paste Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private PowerPoint.Presentation GetActivePresentationOrNull()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    return app.ActivePresentation;
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

        public void AlignLeft_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to align.", "Align Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to align.", "Align Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // The last selected object is the master object (reference)
                float masterLeft = shapes[shapes.Count].Left;

                                 // Align all other objects to the left edge of the master object
                 for (int i = 1; i < shapes.Count; i++)
                 {
                     shapes[i].Left = masterLeft;
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error aligning objects: {ex.Message}", "Align Left", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void AlignRight_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to align.", "Align Right", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 2)
                 {
                     MessageBox.Show("Please select at least 2 objects to align.", "Align Right", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // The last selected object is the master object (reference)
                 float masterRight = shapes[shapes.Count].Left + shapes[shapes.Count].Width;

                 // Align all other objects to the right edge of the master object
                 for (int i = 1; i < shapes.Count; i++)
                 {
                     shapes[i].Left = masterRight - shapes[i].Width;
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error aligning objects: {ex.Message}", "Align Right", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void AlignCenter_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to align.", "Align Center", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 2)
                 {
                     MessageBox.Show("Please select at least 2 objects to align.", "Align Center", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // The last selected object is the master object (reference)
                 var masterShape = shapes[shapes.Count];
                 float masterCenterX = masterShape.Left + (masterShape.Width / 2);

                 // Align all other objects to the horizontal center of the master object
                 for (int i = 1; i < shapes.Count; i++)
                 {
                     shapes[i].Left = masterCenterX - (shapes[i].Width / 2);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error aligning objects: {ex.Message}", "Align Center", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void AlignTop_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to align.", "Align Top", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 2)
                 {
                     MessageBox.Show("Please select at least 2 objects to align.", "Align Top", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // The last selected object is the master object (reference)
                 float masterTop = shapes[shapes.Count].Top;

                 // Align all other objects to the top edge of the master object
                 for (int i = 1; i < shapes.Count; i++)
                 {
                     shapes[i].Top = masterTop;
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error aligning objects: {ex.Message}", "Align Top", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void AlignBottom_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to align.", "Align Bottom", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 2)
                 {
                     MessageBox.Show("Please select at least 2 objects to align.", "Align Bottom", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // The last selected object is the master object (reference)
                 float masterBottom = shapes[shapes.Count].Top + shapes[shapes.Count].Height;

                 // Align all other objects to the bottom edge of the master object
                 for (int i = 1; i < shapes.Count; i++)
                 {
                     shapes[i].Top = masterBottom - shapes[i].Height;
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error aligning objects: {ex.Message}", "Align Bottom", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void AlignMiddle_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to align.", "Align Middle", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 2)
                 {
                     MessageBox.Show("Please select at least 2 objects to align.", "Align Middle", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // The last selected object is the master object (reference)
                 var masterShape = shapes[shapes.Count];
                 float masterCenterY = masterShape.Top + (masterShape.Height / 2);

                 // Align all other objects to the vertical center of the master object
                 for (int i = 1; i < shapes.Count; i++)
                 {
                     shapes[i].Top = masterCenterY - (shapes[i].Height / 2);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error aligning objects: {ex.Message}", "Align Middle", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void RotateLeft_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to rotate.", "Rotate Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to rotate.", "Rotate Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Rotate all selected objects by -90 degrees (counter-clockwise)
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].Rotation = shapes[i].Rotation - 90;
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error rotating objects: {ex.Message}", "Rotate Left", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void RotateRight_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to rotate.", "Rotate Right", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to rotate.", "Rotate Right", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Rotate all selected objects by +90 degrees (clockwise)
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].Rotation = shapes[i].Rotation + 90;
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error rotating objects: {ex.Message}", "Rotate Right", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void FlipHorizontal_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to flip.", "Flip Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to flip.", "Flip Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Flip all selected objects horizontally
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipHorizontal);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error flipping objects horizontally: {ex.Message}", "Flip Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void FlipVertical_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to flip.", "Flip Vertical", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to flip.", "Flip Vertical", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Flip all selected objects vertically
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error flipping objects vertically: {ex.Message}", "Flip Vertical", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void BringForward_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to bring forward.", "Bring Forward", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to bring forward.", "Bring Forward", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Bring all selected objects forward one layer
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error bringing objects forward: {ex.Message}", "Bring Forward", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void SendBackward_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to send backward.", "Send Backward", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to send backward.", "Send Backward", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Send all selected objects backward one layer
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error sending objects backward: {ex.Message}", "Send Backward", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void BringToFront_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to bring to front.", "Bring to Front", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to bring to front.", "Bring to Front", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Bring all selected objects to the very front
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error bringing objects to front: {ex.Message}", "Bring to Front", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         public void SendToBack_Click(Office.IRibbonControl control)
         {
             try
             {
                 var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                 
                 if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                 {
                     MessageBox.Show("Please select objects to send to back.", "Send to Back", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 var shapes = selection.ShapeRange;
                 
                 if (shapes.Count < 1)
                 {
                     MessageBox.Show("Please select at least 1 object to send to back.", "Send to Back", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     return;
                 }

                 // Send all selected objects to the very back
                 for (int i = 1; i <= shapes.Count; i++)
                 {
                     shapes[i].ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error sending objects to back: {ex.Message}", "Send to Back", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
                 }

        public void SwapPosition_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to swap positions.", "Swap Position", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count != 2)
                {
                    MessageBox.Show("Please select exactly 2 objects to swap their positions.", "Swap Position", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // Get the two shapes
                var shape1 = shapes[1];
                var shape2 = shapes[2];
                
                // Calculate center positions of both shapes
                float shape1CenterX = shape1.Left + (shape1.Width / 2);
                float shape1CenterY = shape1.Top + (shape1.Height / 2);
                
                float shape2CenterX = shape2.Left + (shape2.Width / 2);
                float shape2CenterY = shape2.Top + (shape2.Height / 2);
                
                // Calculate new positions to center each shape at the other's center
                float newShape1Left = shape2CenterX - (shape1.Width / 2);
                float newShape1Top = shape2CenterY - (shape1.Height / 2);
                
                float newShape2Left = shape1CenterX - (shape2.Width / 2);
                float newShape2Top = shape1CenterY - (shape2.Height / 2);
                
                // Swap the positions
                shape1.Left = newShape1Left;
                shape1.Top = newShape1Top;
                
                shape2.Left = newShape2Left;
                shape2.Top = newShape2Top;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error swapping positions: {ex.Message}", "Swap Position", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RemoveMarginObjects_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var slide = GetActiveSlideOrNull();
                
                if (slide == null)
                {
                    MessageBox.Show("Please select a slide.", "Remove Margin Objects", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // Get slide dimensions
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                
                var shapesToRemove = new List<PowerPoint.Shape>();
                
                // Check all shapes on the slide
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    
                    // Check if shape is completely outside slide boundaries
                    bool isOutsideLeft = shape.Left + shape.Width < 0;
                    bool isOutsideRight = shape.Left > slideWidth;
                    bool isOutsideTop = shape.Top + shape.Height < 0;
                    bool isOutsideBottom = shape.Top > slideHeight;
                    
                    // If shape is completely outside any boundary, mark for removal
                    if (isOutsideLeft || isOutsideRight || isOutsideTop || isOutsideBottom)
                    {
                        shapesToRemove.Add(shape);
                    }
                }
                
                // Remove the shapes (in reverse order to avoid index issues)
                for (int i = shapesToRemove.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        shapesToRemove[i].Delete();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to delete shape: {ex.Message}");
                    }
                }
                
                // Objects removed silently - user can see the result directly
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error removing margin objects: {ex.Message}", "Remove Margin Objects", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Group_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to group.", "Group", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to group.", "Group", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check for shapes that cannot be grouped
                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shape = shapes[i];
                    
                    // Check if shape is locked
                    try
                    {
                        if (shape.LockAspectRatio == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            // This is just a test to see if we can access shape properties
                        }
                    }
                    catch
                    {
                        MessageBox.Show($"One or more shapes cannot be grouped. They may be locked, protected, or part of the slide master.", "Group", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    // Skip certain shape types that cannot be grouped
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder ||
                        shape.Type == Microsoft.Office.Core.MsoShapeType.msoComment)
                    {
                        MessageBox.Show($"Cannot group placeholder or comment shapes. Please select only regular shapes.", "Group", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Group the selected shapes - this creates a new grouped shape
                var groupedShape = shapes.Group();
                
                // Select the newly created group
                groupedShape.Select();
                
                MessageBox.Show($"Successfully grouped {shapes.Count} objects.", "Group", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error grouping objects: {ex.Message}\n\nThis may occur if shapes are locked, protected, or incompatible for grouping.", "Group", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"Group error: {ex}");
            }
        }

        public void Ungroup_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select grouped objects to ungroup.", "Ungroup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 1)
                {
                    MessageBox.Show("Please select at least 1 grouped object to ungroup.", "Ungroup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int ungroupedCount = 0;
                
                // Ungroup each selected shape if it's a group
                for (int i = 1; i <= shapes.Count; i++)
                {
                    try
                    {
                        var shape = shapes[i];
                        
                        // Check if the shape is a group before trying to ungroup
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            var ungroupedShapes = shape.Ungroup();
                            ungroupedCount += ungroupedShapes.Count;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Shape {i} is not a group (Type: {shape.Type})");
                        }
                    }
                    catch (Exception innerEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to ungroup shape {i}: {innerEx.Message}");
                    }
                }
                
                if (ungroupedCount > 0)
                {
                    MessageBox.Show($"Successfully ungrouped {ungroupedCount} objects.", "Ungroup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No grouped objects found to ungroup.", "Ungroup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error ungrouping objects: {ex.Message}", "Ungroup", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"Ungroup error: {ex}");
            }
        }

        // =================== FORMAT FUNCTIONS ===================
        
        #region Resize Functions
        
        public void Format_Resize_Click(Office.IRibbonControl control)
        {
            Format_ResizeSize_Click(control);
        }
        
        public void Format_ResizeWidth_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to resize.", "Match Width", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to match width.", "Match Width", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                float masterWidth = shapes[shapes.Count].Width;
                
                // Apply width to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    shapes[i].Width = masterWidth;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching width: {ex.Message}", "Match Width", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void Format_ResizeHeight_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to resize.", "Match Height", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to match height.", "Match Height", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                float masterHeight = shapes[shapes.Count].Height;
                
                // Apply height to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    shapes[i].Height = masterHeight;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching height: {ex.Message}", "Match Height", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void Format_ResizeSize_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to resize.", "Match Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to match size.", "Match Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                float masterWidth = shapes[shapes.Count].Width;
                float masterHeight = shapes[shapes.Count].Height;
                
                // Apply both width and height to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    shapes[i].Width = masterWidth;
                    shapes[i].Height = masterHeight;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching size: {ex.Message}", "Match Size", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        #endregion
        
        #region Color Matching Functions
        
        public void Format_MatchColors_Click(Office.IRibbonControl control)
        {
            Format_MatchFill_Click(control);
        }
        
        public void Format_MatchFill_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to match fill color.", "Match Fill", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to match fill color.", "Match Fill", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                var masterShape = shapes[shapes.Count];
                var masterFill = masterShape.Fill;
                
                // Apply fill to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    var targetShape = shapes[i];
                    
                    // Copy fill properties
                    targetShape.Fill.ForeColor.RGB = masterFill.ForeColor.RGB;
                    targetShape.Fill.Visible = masterFill.Visible;
                    targetShape.Fill.Transparency = masterFill.Transparency;
                    
                    if (masterFill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                    {
                        targetShape.Fill.Solid();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching fill color: {ex.Message}", "Match Fill", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void Format_MatchOutline_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select objects to match outline color.", "Match Outline", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 objects to match outline color.", "Match Outline", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                var masterShape = shapes[shapes.Count];
                var masterLine = masterShape.Line;
                
                // Apply outline to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    var targetShape = shapes[i];
                    
                    // Copy line properties
                    targetShape.Line.ForeColor.RGB = masterLine.ForeColor.RGB;
                    targetShape.Line.Visible = masterLine.Visible;
                    targetShape.Line.Weight = masterLine.Weight;
                    targetShape.Line.Transparency = masterLine.Transparency;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching outline color: {ex.Message}", "Match Outline", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void Format_MatchFont_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select text objects to match font color.", "Match Font Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 text objects to match font color.", "Match Font Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                var masterShape = shapes[shapes.Count];
                
                if (masterShape.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue || masterShape.TextFrame.TextRange.Text.Trim() == "")
                {
                    MessageBox.Show("Master object must contain text.", "Match Font Color", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                var masterFont = masterShape.TextFrame.TextRange.Font;
                
                // Apply font color to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    var targetShape = shapes[i];
                    
                    if (targetShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && targetShape.TextFrame.TextRange.Text.Trim() != "")
                    {
                        targetShape.TextFrame.TextRange.Font.Color.RGB = masterFont.Color.RGB;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching font color: {ex.Message}", "Match Font Color", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        #endregion
        
        #region Font Matching Functions
        
        public void Format_AlignFonts_Click(Office.IRibbonControl control)
        {
            Format_MatchFontFamily_Click(control);
        }
        
        public void Format_MatchFontColor_Click(Office.IRibbonControl control)
        {
            Format_MatchFont_Click(control);
        }
        
        public void Format_MatchFontSize_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select text objects to match font size.", "Match Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 text objects to match font size.", "Match Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                var masterShape = shapes[shapes.Count];
                
                if (masterShape.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue || masterShape.TextFrame.TextRange.Text.Trim() == "")
                {
                    MessageBox.Show("Master object must contain text.", "Match Font Size", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                var masterFontSize = masterShape.TextFrame.TextRange.Font.Size;
                
                // Apply font size to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    var targetShape = shapes[i];
                    
                    if (targetShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && targetShape.TextFrame.TextRange.Text.Trim() != "")
                    {
                        targetShape.TextFrame.TextRange.Font.Size = masterFontSize;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching font size: {ex.Message}", "Match Font Size", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void Format_MatchFontFamily_Click(Office.IRibbonControl control)
        {
            try
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select text objects to match font family.", "Match Font Family", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var shapes = selection.ShapeRange;
                
                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select at least 2 text objects to match font family.", "Match Font Family", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // The last selected object is the master object (reference)
                var masterShape = shapes[shapes.Count];
                
                if (masterShape.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoTrue || masterShape.TextFrame.TextRange.Text.Trim() == "")
                {
                    MessageBox.Show("Master object must contain text.", "Match Font Family", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                var masterFontName = masterShape.TextFrame.TextRange.Font.Name;
                
                // Apply font family to all other objects
                for (int i = 1; i < shapes.Count; i++)
                {
                    var targetShape = shapes[i];
                    
                    if (targetShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && targetShape.TextFrame.TextRange.Text.Trim() != "")
                    {
                        targetShape.TextFrame.TextRange.Font.Name = masterFontName;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching font family: {ex.Message}", "Match Font Family", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        #endregion
        
    }
}