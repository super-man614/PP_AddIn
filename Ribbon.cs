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