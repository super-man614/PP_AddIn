using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _customTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Initialize the custom task pane
                InitializeTaskPane();
                
                // Add ribbon controls
                AddRibbonControls();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error during startup: {ex.Message}", 
                    "Startup Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up the task pane
            if (_customTaskPane != null)
            {
                _customTaskPane.Dispose();
                _customTaskPane = null;
            }
        }

        /// <summary>
        /// Initializes the custom task pane
        /// </summary>
        private void InitializeTaskPane()
        {
            try
            {
                _customTaskPane = new CustomTaskPane();
                
                // Show the task pane on startup
                _customTaskPane.Show();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error initializing task pane: {ex.Message}", 
                    "Initialization Error", System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Adds ribbon controls for the add-in
        /// </summary>
        private void AddRibbonControls()
        {
            // The ribbon is now handled by the Ribbon.cs class and Ribbon.xml
            // No additional setup needed here as the ribbon is automatically loaded
        }

        /// <summary>
        /// Public method to toggle the task pane (can be called from ribbon or other UI)
        /// </summary>
        public void ToggleTaskPane()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Toggle();
            }
        }

        /// <summary>
        /// Public method to show the task pane
        /// </summary>
        public void ShowTaskPane()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Show();
            }
        }

        /// <summary>
        /// Public method to hide the task pane
        /// </summary>
        public void HideTaskPane()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Hide();
            }
        }

        /// <summary>
        /// Gets the custom task pane instance
        /// </summary>
        public CustomTaskPane TaskPane
        {
            get { return _customTaskPane; }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
