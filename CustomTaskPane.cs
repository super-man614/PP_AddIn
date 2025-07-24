using System;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace my_addin
{
    public class CustomTaskPane
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;

        public CustomTaskPane()
        {
            try
            {
                // Create the user control for the task pane
                _taskPaneControl = new TaskPaneControl();
                
                // Create the custom task pane
                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_taskPaneControl, "PowerPoint Tools");
                
                // Set task pane properties
                _taskPane.Width = 320;
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = false; // Initially hidden
                
                // Handle visibility events
                _taskPane.VisibleChanged += TaskPane_VisibleChanged;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error creating task pane: {ex.Message}", 
                    "Task Pane Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
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
        /// Disposes of the task pane resources
        /// </summary>
        public void Dispose()
        {
            if (_taskPane != null)
            {
                _taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                
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
        }
    }
} 