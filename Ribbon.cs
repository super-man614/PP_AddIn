using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace my_addin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("my_addin.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// Called when the ribbon is loaded
        /// </summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// Toggle the task pane visibility
        /// </summary>
        public void ToggleTaskPane_Click(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.ToggleTaskPane();
                
                // Update the button label based on visibility
                if (ribbon != null)
                {
                    ribbon.InvalidateControl("ToggleTaskPaneButton");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error toggling task pane: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Create a new presentation
        /// </summary>
        public void NewPresentation_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                app.Presentations.Add();
                MessageBox.Show("New presentation created!", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating presentation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Export current presentation to PDF
        /// </summary>
        public void ExportPDF_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "PDF Files (*.pdf)|*.pdf";
                    saveDialog.Title = "Export to PDF";
                    saveDialog.FileName = app.ActivePresentation.Name.Replace(".pptx", "").Replace(".ppt", "") + ".pdf";
                    
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        app.ActivePresentation.ExportAsFixedFormat(saveDialog.FileName, 
                            PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                        MessageBox.Show("PDF exported successfully!", "Success", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("No active presentation to export.", "Warning", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting PDF: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Add a new slide to the current presentation
        /// </summary>
        public void AddSlide_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActivePresentation.Slides.Add(app.ActivePresentation.Slides.Count + 1, 
                        PowerPoint.PpSlideLayout.ppLayoutBlank);
                    MessageBox.Show("Slide added successfully!", "Success", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No active presentation to add slide to.", "Warning", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding slide: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Duplicate the current slide
        /// </summary>
        public void DuplicateSlide_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow != null)
                {
                    var slideIndex = app.ActiveWindow.Selection.SlideRange[1].SlideIndex;
                    app.ActivePresentation.Slides[slideIndex].Duplicate();
                    MessageBox.Show("Slide duplicated successfully!", "Success", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No active slide to duplicate.", "Warning", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error duplicating slide: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Get the label for the toggle task pane button
        /// </summary>
        public string GetToggleTaskPaneLabel(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.TaskPane != null && Globals.ThisAddIn.TaskPane.Visible)
                {
                    return "Hide Tools";
                }
                else
                {
                    return "Show Tools";
                }
            }
            catch
            {
                return "Show Tools";
            }
        }

        #endregion

        #region File Operations

        /// <summary>
        /// Create a new presentation
        /// </summary>
        public void New_Click(Office.IRibbonControl control)
        {
            try
            {
                // Find the task pane control to reuse its functionality
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    // Use reflection to call the existing BtnNew_Click method
                    var method = taskPaneControl.GetType().GetMethod("BtnNew_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    // Fallback implementation
                    var app = Globals.ThisAddIn.Application;
                    app.Presentations.Add();
                    MessageBox.Show("New presentation created!", "Success", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating new presentation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Open an existing presentation
        /// </summary>
        public void Open_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnOpen_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    // Fallback implementation
                    var openDialog = new OpenFileDialog();
                    openDialog.Filter = "PowerPoint Files (*.pptx;*.ppt)|*.pptx;*.ppt";
                    openDialog.Title = "Open Presentation";
                    
                    if (openDialog.ShowDialog() == DialogResult.OK)
                    {
                        var app = Globals.ThisAddIn.Application;
                        app.Presentations.Open(openDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening presentation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Save the current presentation
        /// </summary>
        public void Save_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnSave_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    // Fallback implementation
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActivePresentation != null)
                    {
                        app.ActivePresentation.Save();
                        MessageBox.Show("Presentation saved!", "Success", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("No active presentation to save.", "Warning", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving presentation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Save the presentation with a new name
        /// </summary>
        public void SaveAs_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnSaveAs_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    // Fallback implementation
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActivePresentation != null)
                    {
                        var saveDialog = new SaveFileDialog();
                        saveDialog.Filter = "PowerPoint Files (*.pptx)|*.pptx";
                        saveDialog.Title = "Save Presentation As";
                        
                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            app.ActivePresentation.SaveAs(saveDialog.FileName);
                            MessageBox.Show("Presentation saved!", "Success", 
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No active presentation to save.", "Warning", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving presentation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Print the current presentation
        /// </summary>
        public void Print_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnPrint_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    // Fallback implementation
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActivePresentation != null)
                    {
                        app.ActivePresentation.PrintOut();
                    }
                    else
                    {
                        MessageBox.Show("No active presentation to print.", "Warning", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error printing presentation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Share the current presentation
        /// </summary>
        public void Share_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnShare_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    MessageBox.Show("Share functionality accessed via ribbon.", "Share", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing share function: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Wizard Functions

        /// <summary>
        /// Create an agenda slide
        /// </summary>
        public void Agenda_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAgenda_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating agenda: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Access slide master
        /// </summary>
        public void Master_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMaster_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing master: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert design elements
        /// </summary>
        public void Element_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnElement_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting element: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Text formatting wizard
        /// </summary>
        public void TextWizard_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnTextWizard_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing text wizard: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Format shapes and objects
        /// </summary>
        public void Format_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnFormat_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing format function: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert maps and diagrams
        /// </summary>
        public void Map_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMap_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting map: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Smart Elements

        /// <summary>
        /// Insert charts and graphs
        /// </summary>
        public void Chart_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnChart_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting chart: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert diagrams and flowcharts
        /// </summary>
        public void Diagram_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnDiagram_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting diagram: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert tables
        /// </summary>
        public void Table_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnTable_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting table: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert matrix tables
        /// </summary>
        public void MatrixTable_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMatrixTable_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting matrix table: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Add sticky notes
        /// </summary>
        public void StickyNote_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnStickyNote_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting sticky note: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert citations
        /// </summary>
        public void Citation_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnCitation_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting citation: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Insert standard objects
        /// </summary>
        public void StandardObjects_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnStandardObjects_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting standard objects: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Position Operations

        /// <summary>
        /// Align objects to the left
        /// </summary>
        public void AlignLeft_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignLeft_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning left: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Center align objects horizontally
        /// </summary>
        public void AlignCenter_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignCenter_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning center: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Align objects to the right
        /// </summary>
        public void AlignRight_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignRight_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning right: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Distribute objects evenly
        /// </summary>
        public void Distribute_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnDistribute_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error distributing objects: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Match width and height
        /// </summary>
        public void MatchBoth_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMatchBoth_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching both dimensions: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Match object heights
        /// </summary>
        public void MatchHeight_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMatchHeight_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching height: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Match object widths
        /// </summary>
        public void MatchWidth_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMatchWidth_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching width: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Transform Operations

        /// <summary>
        /// Rotate objects to vertical
        /// </summary>
        public void MakeVertical_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMakeVertical_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error making vertical: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Rotate objects to horizontal
        /// </summary>
        public void MakeHorizontal_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnMakeHorizontal_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error making horizontal: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Swap object locations
        /// </summary>
        public void SwapLocations_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnSwapLocations_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error swapping locations: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Size Operations

        private string[] slideSizes = { "4:3 Standard", "16:9 Widescreen", "16:10 Widescreen", "A3", "A4", "B4 (JIS)", "B5 (JIS)", "Banner", "Ledger", "Letter", "Overhead", "35mm Slides", "Custom" };
        private int selectedSizeIndex = 1; // Default to 16:9

        /// <summary>
        /// Handle slide size selection change
        /// </summary>
        public void SlideSize_Change(Office.IRibbonControl control, string text)
        {
            try
            {
                // Find the index of the selected item
                for (int i = 0; i < slideSizes.Length; i++)
                {
                    if (slideSizes[i] == text)
                    {
                        selectedSizeIndex = i;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting slide size: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Get the current slide size text
        /// </summary>
        public string GetSlideSizeText(Office.IRibbonControl control)
        {
            return slideSizes[selectedSizeIndex];
        }

        /// <summary>
        /// Get the number of slide size items
        /// </summary>
        public int GetSlideSizeItemCount(Office.IRibbonControl control)
        {
            return slideSizes.Length;
        }

        /// <summary>
        /// Get the label for a slide size item
        /// </summary>
        public string GetSlideSizeItemLabel(Office.IRibbonControl control, int index)
        {
            return index < slideSizes.Length ? slideSizes[index] : "";
        }

        /// <summary>
        /// Apply the selected slide size
        /// </summary>
        public void ApplySize_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnApplySize_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
                else
                {
                    // Fallback implementation
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActivePresentation != null)
                    {
                        var presentation = app.ActivePresentation;
                        var pageSetup = presentation.PageSetup;
                        
                        switch (selectedSizeIndex)
                        {
                            case 0: // 4:3 Standard
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                                break;
                            case 1: // 16:9 Widescreen
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen16x9;
                                break;
                            case 2: // 16:10 Widescreen
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen16x10;
                                break;
                            case 3: // A3
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA3Paper;
                                break;
                            case 4: // A4
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA4Paper;
                                break;
                            case 5: // B4 (JIS)
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeB4JISPaper;
                                break;
                            case 6: // B5 (JIS)
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeB5JISPaper;
                                break;
                            case 7: // Banner
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeBanner;
                                break;
                            case 8: // Ledger
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeLedgerPaper;
                                break;
                            case 9: // Letter
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeLetterPaper;
                                break;
                            case 10: // Overhead
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOverhead;
                                break;
                            case 11: // 35mm Slides
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSize35MM;
                                break;
                            default: // Custom or fallback
                                pageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen16x9;
                                break;
                        }
                        
                        MessageBox.Show($"Slide size changed to {slideSizes[selectedSizeIndex]}!", "Success", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("No active presentation to resize.", "Warning", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying slide size: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Shape Operations

        /// <summary>
        /// Align process chain
        /// </summary>
        public void AlignProcessChain_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignProcessChain_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning process chain: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Align angles
        /// </summary>
        public void AlignAngles_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignAngles_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning angles: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Align to process arrow
        /// </summary>
        public void AlignToProcessArrow_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignToProcessArrow_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning to process arrow: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Adjust pentagon header
        /// </summary>
        public void AdjustPentagonHeader_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAdjustPentagonHeader_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adjusting pentagon header: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Align block arrows
        /// </summary>
        public void AlignBlockArrows_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignBlockArrows_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning block arrows: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Align rounded rectangle arrows
        /// </summary>
        public void AlignRoundedRectangleArrows_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnAlignRoundedRectangleArrows_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning rounded rectangle arrows: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Color Operations

        /// <summary>
        /// Change fill color
        /// </summary>
        public void FillColor_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnFillColor_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing fill color: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Change text color
        /// </summary>
        public void TextColor_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnTextColor_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing text color: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Change outline color
        /// </summary>
        public void OutlineColor_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnOutlineColor_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing outline color: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Text Formatting

        /// <summary>
        /// Make text bold
        /// </summary>
        public void Bold_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnBold_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying bold: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Make text italic
        /// </summary>
        public void Italic_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnItalic_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying italic: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Underline text
        /// </summary>
        public void Underline_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnUnderline_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying underline: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Add bullet points
        /// </summary>
        public void Bullets_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnBullets_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding bullets: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Wrap text in shape
        /// </summary>
        public void WrapText_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnWrapText_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error wrapping text: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Disable text wrapping
        /// </summary>
        public void NoWrapText_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnNoWrapText_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error disabling text wrap: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Navigation

        /// <summary>
        /// Zoom in on the slide
        /// </summary>
        public void ZoomIn_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnZoomIn_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error zooming in: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Zoom out on the slide
        /// </summary>
        public void ZoomOut_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnZoomOut_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error zooming out: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Fit slide to window
        /// </summary>
        public void FitToWindow_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnFitToWindow_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fitting to window: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Expert Tools

        /// <summary>
        /// Access free webinar resources
        /// </summary>
        public void FreeWebinar_Click(Office.IRibbonControl control)
        {
            try
            {
                var taskPaneControl = Globals.ThisAddIn.TaskPane?.TaskPaneControl;
                if (taskPaneControl != null)
                {
                    var method = taskPaneControl.GetType().GetMethod("BtnFreeWebinar_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    method?.Invoke(taskPaneControl, new object[] { null, EventArgs.Empty });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing webinar: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
} 