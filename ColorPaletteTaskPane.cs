using System;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using my_addin.Core;

namespace my_addin
{
    public class ColorPaletteTaskPane
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private ColorPaletteControl _colorPaletteControl;
        private bool _isDisposed = false;

        public ColorPaletteTaskPane()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Creating ColorPaletteControl...");
                
                // Validate that ThisAddIn is available
                if (Globals.ThisAddIn == null)
                {
                    throw new InvalidOperationException("Globals.ThisAddIn is null - add-in not properly initialized");
                }
                
                if (Globals.ThisAddIn.CustomTaskPanes == null)
                {
                    throw new InvalidOperationException("CustomTaskPanes collection is null");
                }
                
                // Create the user control for the color palette
                _colorPaletteControl = new ColorPaletteControl();
                System.Diagnostics.Debug.WriteLine("ColorPaletteControl created successfully");
                
                System.Diagnostics.Debug.WriteLine("Adding Color Palette task pane to collection...");
                // Create the custom task pane
                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_colorPaletteControl, "Color Palette");
                System.Diagnostics.Debug.WriteLine("Color Palette task pane added to collection");
                Core.PaneManager.Register(_taskPane);
                
                // Set task pane properties
                _taskPane.Width = 140;
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true; // Auto-show Color Palette on startup
                System.Diagnostics.Debug.WriteLine($"Color Palette properties set - Width: {_taskPane.Width}, Dock: {_taskPane.DockPosition}, Visible: {_taskPane.Visible}");
                
                // Handle visibility events
                _taskPane.VisibleChanged += TaskPane_VisibleChanged;
                System.Diagnostics.Debug.WriteLine("Color Palette event handlers attached");
                System.Diagnostics.Debug.WriteLine("ColorPaletteTaskPane creation completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating Color Palette task pane: {ex.Message}");
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
            // Keep Color Palette at left-most among right-docked panes
            try
            {
                Core.PaneOrdering.EnsureColorPaletteLeftMost(this);
                Core.PaneManager.OnPaneVisibilityChanged();
            }
            catch { }

            System.Diagnostics.Debug.WriteLine($"Color Palette visibility changed: {_taskPane.Visible}");
        }

        /// <summary>
        /// Gets the underlying task pane control
        /// </summary>
        public ColorPaletteControl ColorPaletteControl
        {
            get { return _colorPaletteControl; }
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
            
            if (_colorPaletteControl != null)
            {
                _colorPaletteControl.Dispose();
                _colorPaletteControl = null;
            }

            _isDisposed = true;
        }
    }
} 