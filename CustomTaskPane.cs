using System;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace my_addin
{
    public class CustomTaskPane
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;
        private bool _isDisposed = false;

        // File operation methods
        public void ExecuteNewFile()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteNewFile();
            }
        }

        public void ExecuteOpenFile()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteOpenFile();
            }
        }

        public void ExecuteSaveFile()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteSaveFile();
            }
        }

        public void ExecuteSaveAsFile()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteSaveAsFile();
            }
        }

        public void ExecutePrint()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecutePrint();
            }
        }

        public void ExecuteShare()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteShare();
            }
        }

        public CustomTaskPane()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Creating TaskPaneControl...");
                
                // Validate that ThisAddIn is available
                if (Globals.ThisAddIn == null)
                {
                    throw new InvalidOperationException("Globals.ThisAddIn is null - add-in not properly initialized");
                }
                
                if (Globals.ThisAddIn.CustomTaskPanes == null)
                {
                    throw new InvalidOperationException("CustomTaskPanes collection is null");
                }
                
                // Create the user control for the task pane
                _taskPaneControl = new TaskPaneControl();
                System.Diagnostics.Debug.WriteLine("TaskPaneControl created successfully");
                
                System.Diagnostics.Debug.WriteLine("Adding task pane to collection...");
                // Create the custom task pane
                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_taskPaneControl, "PowerPoint Tools");
                System.Diagnostics.Debug.WriteLine("Task pane added to collection");
                Core.PaneManager.Register(_taskPane);
                
                // Set task pane properties
                _taskPane.Width = 320;
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true; // Start visible at startup
                System.Diagnostics.Debug.WriteLine($"Task pane properties set - Width: {_taskPane.Width}, Dock: {_taskPane.DockPosition}, Visible: {_taskPane.Visible}");
                
                // Handle visibility events
                _taskPane.VisibleChanged += TaskPane_VisibleChanged;
                System.Diagnostics.Debug.WriteLine("Task pane event handlers attached");
                System.Diagnostics.Debug.WriteLine("CustomTaskPane creation completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating task pane: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw to let the caller handle it
            }
        }

        /// <summary>
        /// Gets or sets the visibility of the task pane
        /// </summary>
        public bool Visible
        {
            get { return _taskPane.Visible; }
            set { _taskPane.Visible = value; }
        }

        /// <summary>
        /// Gets the width of the task pane
        /// </summary>
        public int Width
        {
            get { return _taskPane.Width; }
            set { _taskPane.Width = value; }
        }

        /// <summary>
        /// Gets or sets the dock position of the task pane
        /// </summary>
        public Microsoft.Office.Core.MsoCTPDockPosition DockPosition
        {
            get { return _taskPane.DockPosition; }
            set { _taskPane.DockPosition = value; }
        }

        /// <summary>
        /// Toggles the visibility of the task pane
        /// </summary>
        public void Toggle()
        {
            _taskPane.Visible = !_taskPane.Visible;
        }

        /// <summary>
        /// Shows the task pane
        /// </summary>
        public void Show()
        {
            _taskPane.Visible = true;
        }

        /// <summary>
        /// Hides the task pane
        /// </summary>
        public void Hide()
        {
            _taskPane.Visible = false;
        }

        /// <summary>
        /// Event handler for when the task pane visibility changes
        /// </summary>
        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            // You can add custom logic here when the task pane is shown or hidden
            // For example, refresh data when the pane becomes visible
            if (_taskPane.Visible && _taskPaneControl != null)
            {
                // Refresh slide list when task pane becomes visible
                // This is handled in the TaskPaneControl itself
            }
            try { Core.PaneManager.OnPaneVisibilityChanged(); } catch { }
        }

        /// <summary>
        /// Gets the underlying task pane control
        /// </summary>
        public TaskPaneControl TaskPaneControl
        {
            get { return _taskPaneControl; }
        }

        /// <summary>
        /// Gets the underlying Microsoft task pane object
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get { return _taskPane; }
        }

        /// <summary>
        /// Gets a value indicating whether this instance has been disposed
        /// </summary>
        public bool IsDisposed
        {
            get { return _isDisposed; }
        }

        // Wizard execution methods
        public void ExecuteAgendaWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteAgendaWizard();
            }
        }

        public void ExecuteElementWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteElementWizard();
            }
        }

        public void ExecuteMasterWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteMasterWizard();
            }
        }

        public void ExecuteTextWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteTextWizard();
            }
        }

        public void ExecuteFormatWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteFormatWizard();
            }
        }

        public void ExecuteMapWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteMapWizard();
            }
        }

        public void ExecuteChartWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteChartWizard();
            }
        }

        public void ExecuteDiagramWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteDiagramWizard();
            }
        }

        public void ExecuteTableWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteTableWizard();
            }
        }

        public void ExecuteMatrixTableWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteMatrixTableWizard();
            }
        }

        public void ExecuteExcelPaste()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteExcelPaste();
            }
        }

        public void ExecuteStickyNoteWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteStickyNoteWizard();
            }
        }

        public void ExecuteCitationWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteCitationWizard();
            }
        }

        public void ExecuteStandardObjectsWizard()
        {
            if (_taskPaneControl != null && !_isDisposed)
            {
                _taskPaneControl.ExecuteStandardObjectsWizard();
            }
        }

        /// <summary>
        /// Disposes of the task pane resources
        /// </summary>
        public void Dispose()
        {
            if (_isDisposed)
                return;

            if (_taskPane != null)
            {
                _taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                Core.PaneManager.Unregister(_taskPane);
                
                // Remove the task pane from the collection
                if (Globals.ThisAddIn != null && Globals.ThisAddIn.CustomTaskPanes != null)
                {
                    try
                    {
                        Globals.ThisAddIn.CustomTaskPanes.Remove(_taskPane);
                    }
                    catch
                    {
                        // Ignore errors during cleanup
                    }
                }
                
                _taskPane = null;
            }
            
            if (_taskPaneControl != null)
            {
                _taskPaneControl.Dispose();
                _taskPaneControl = null;
            }

            _isDisposed = true;
        }
    }
}