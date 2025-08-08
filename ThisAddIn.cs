using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("=== PowerPoint Add-in Started ===");
            System.Diagnostics.Debug.WriteLine("Add-in is initializing...");
            System.Diagnostics.Debug.WriteLine("Ribbon should be created automatically by VSTO framework");
            
            // Auto-create Color Palette taskpane on startup
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
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("=== PowerPoint Add-in Shutdown ===");
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