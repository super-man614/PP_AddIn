using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using PowerPointAddIn.Services;

namespace my_addin
{
    public partial class ThisAddIn
    {
        // Removed unused field to fix CS0169 warning
        // private PowerPoint.EApplication_WindowSelectionChangeEventHandler selectionChangedHandler;
        private IErrorHandlerService _errorHandler;
        private CustomTaskPane _taskPaneInstance;
        private ColorPaletteTaskPane _colorPaletteInstance;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // Initialize error handler first
                _errorHandler = new ErrorHandlerService();
                
                // Global exception handling
                AppDomain.CurrentDomain.UnhandledException += (s, exArgs) =>
                {
                    try { (_errorHandler ?? new PowerPointAddIn.Services.ErrorHandlerService()).HandleError(exArgs.ExceptionObject as Exception ?? new Exception("Unknown unhandled exception"), "An unexpected error occurred"); } catch {}
                };
                System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (s, exArgs) =>
                {
                    try { (_errorHandler ?? new PowerPointAddIn.Services.ErrorHandlerService()).HandleError(exArgs.Exception, "An unexpected task error occurred"); exArgs.SetObserved(); } catch {}
                };

                _errorHandler.LogInfo("=== PowerPoint Add-in Started ===");
                
                // Auto-create taskpanes on startup
                CreateTaskPanes();
                
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
                _taskPaneInstance = new CustomTaskPane();
                if (Ribbon.Current != null)
                {
                    Ribbon.TaskPaneInstance = _taskPaneInstance;
                }
                _errorHandler.LogInfo("PowerPoint Tools taskpane created successfully");
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "Error auto-creating PowerPoint Tools taskpane");
            }

            try
            {
                _errorHandler.LogInfo("Auto-creating Color Palette taskpane...");
                _colorPaletteInstance = new ColorPaletteTaskPane();
                
                if (Ribbon.Current != null)
                {
                    Ribbon.ColorPaletteInstance = _colorPaletteInstance;
                }
                _errorHandler.LogInfo("Color Palette taskpane created successfully");
            }
            catch (Exception ex)
            {
                _errorHandler?.LogError(ex, "Error auto-creating Color Palette");
            }
        }

        public bool TryPasteIntoMatrix()
        {
            try
            {
                // Simple method to return false - placeholder for matrix paste functionality
                // This can be enhanced later if needed
                return false;
            }
            catch
            {
                return false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                _errorHandler?.LogInfo("=== PowerPoint Add-in Shutdown ===");
                
                // Clean up task panes
                _taskPaneInstance?.Dispose();
                _colorPaletteInstance?.Dispose();
                
                // Clean up old log files
                _errorHandler?.CleanupOldLogs();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
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