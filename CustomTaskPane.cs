using System;
using Microsoft.Office.Tools;

namespace my_addin
{
    public class CustomTaskPane : IDisposable
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;
        private bool _isDisposed;

        // Unnecessary: Repeated null and disposed checks in every method.
        // Optimization: Use a single helper method to execute actions if not disposed.

        public CustomTaskPane()
        {
            if (Globals.ThisAddIn?.CustomTaskPanes == null)
                throw new InvalidOperationException("Add-in not properly initialized");

            _taskPaneControl = new TaskPaneControl();
            _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_taskPaneControl, "PowerPoint Tools");
            Core.PaneManager.Register(_taskPane);

            _taskPane.Width = 320;
            _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Visible = true;
            _taskPane.VisibleChanged += TaskPane_VisibleChanged;
        }

        // Helper to execute actions if not disposed
        private void ExecuteIfAvailable(Action<TaskPaneControl> action)
        {
            if (_taskPaneControl != null && !_isDisposed)
                action(_taskPaneControl);
        }

        // File operation methods
        public void ExecuteNewFile() => ExecuteIfAvailable(c => c.ExecuteNewFile());
        public void ExecuteOpenFile() => ExecuteIfAvailable(c => c.ExecuteOpenFile());
        public void ExecuteSaveFile() => ExecuteIfAvailable(c => c.ExecuteSaveFile());
        public void ExecuteSaveAsFile() => ExecuteIfAvailable(c => c.ExecuteSaveAsFile());
        public void ExecutePrint() => ExecuteIfAvailable(c => c.ExecutePrint());
        public void ExecuteShare() => ExecuteIfAvailable(c => c.ExecuteShare());

        // Wizard execution methods
        public void ExecuteAgendaWizard() => ExecuteIfAvailable(c => c.ExecuteAgendaWizard());
        public void ExecuteElementWizard() => ExecuteIfAvailable(c => c.ExecuteElementWizard());
        public void ExecuteMasterWizard() => ExecuteIfAvailable(c => c.ExecuteMasterWizard());
        public void ExecuteTextWizard() => ExecuteIfAvailable(c => c.ExecuteTextWizard());
        public void ExecuteFormatWizard() => ExecuteIfAvailable(c => c.ExecuteFormatWizard());
        public void ExecuteMapWizard() => ExecuteIfAvailable(c => c.ExecuteMapWizard());
        public void ExecuteChartWizard() => ExecuteIfAvailable(c => c.ExecuteChartWizard());
        public void ExecuteDiagramWizard() => ExecuteIfAvailable(c => c.ExecuteDiagramWizard());
        public void ExecuteTableWizard() => ExecuteIfAvailable(c => c.ExecuteTableWizard());
        public void ExecuteMatrixTableWizard() => ExecuteIfAvailable(c => c.ExecuteMatrixTableWizard());
        public void ExecuteExcelPaste() => ExecuteIfAvailable(c => c.ExecuteExcelPaste());
        public void ExecuteStickyNoteWizard() => ExecuteIfAvailable(c => c.ExecuteStickyNoteWizard());
        public void ExecuteCitationWizard() => ExecuteIfAvailable(c => c.ExecuteCitationWizard());
        public void ExecuteStandardObjectsWizard() => ExecuteIfAvailable(c => c.ExecuteStandardObjectsWizard());

        public bool Visible
        {
            get => _taskPane != null && _taskPane.Visible;
            set { if (_taskPane != null) _taskPane.Visible = value; }
        }

        public int Width
        {
            get => _taskPane != null ? _taskPane.Width : 0;
            set { if (_taskPane != null) _taskPane.Width = value; }
        }

        public Microsoft.Office.Core.MsoCTPDockPosition DockPosition
        {
            get => _taskPane != null ? _taskPane.DockPosition : Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            set { if (_taskPane != null) _taskPane.DockPosition = value; }
        }

        public void Toggle() { if (_taskPane != null) _taskPane.Visible = !_taskPane.Visible; }
        public void Show() { if (_taskPane != null) _taskPane.Visible = true; }
        public void Hide() { if (_taskPane != null) _taskPane.Visible = false; }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            try { Core.PaneManager.OnPaneVisibilityChanged(); } catch { }
        }

        public TaskPaneControl TaskPaneControl => _taskPaneControl;
        public Microsoft.Office.Tools.CustomTaskPane TaskPane => _taskPane;
        public bool IsDisposed => _isDisposed;

        public void Dispose()
        {
            if (_isDisposed) return;

            if (_taskPane != null)
            {
                _taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                Core.PaneManager.Unregister(_taskPane);

                try
                {
                    Globals.ThisAddIn?.CustomTaskPanes?.Remove(_taskPane);
                }
                catch { /* Ignore errors during cleanup */ }

                _taskPane = null;
            }

            _taskPaneControl?.Dispose();
            _taskPaneControl = null;

            _isDisposed = true;
        }
    }
}
