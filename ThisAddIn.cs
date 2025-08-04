using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace my_addin
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _customTaskPane;
        private Ribbon _ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup called");
                
                // Add-in is loading - debug message removed for production
                
                // Initialize the custom task pane
                InitializeTaskPane();
                
                // Test ribbon registration
                TestRibbonRegistration();
                
                // Ribbon will be automatically loaded via IRibbonExtensibility interface
                System.Diagnostics.Debug.WriteLine("Task pane initialized, ribbon should load automatically");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error during startup: {ex.Message}", 
                    "Startup Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"Startup error: {ex}");
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
        
        /// <summary>
        /// Test method to verify ribbon registration
        /// </summary>
        public void TestRibbonRegistration()
        {
            try
            {
                // Test if ribbon interface is available
                var ribbon = new Ribbon();
                string customUI = ribbon.GetCustomUI("Microsoft.PowerPoint");
                System.Diagnostics.Debug.WriteLine($"Ribbon test - CustomUI available: {!string.IsNullOrEmpty(customUI)}");
                if (!string.IsNullOrEmpty(customUI))
                {
                    System.Diagnostics.Debug.WriteLine($"CustomUI length: {customUI.Length}");
                    System.Diagnostics.Debug.WriteLine($"CustomUI preview: {customUI.Substring(0, Math.Min(200, customUI.Length))}...");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ribbon test failed: {ex.Message}");
            }
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
