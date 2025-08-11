using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;

namespace my_addin.Core
{
    public static class PaneManager
    {
        private static readonly List<Microsoft.Office.Tools.CustomTaskPane> registered = new List<Microsoft.Office.Tools.CustomTaskPane>();

        public static void Register(Microsoft.Office.Tools.CustomTaskPane pane)
        {
            if (pane == null) return;
            if (!registered.Contains(pane)) registered.Add(pane);
        }

        public static void Unregister(Microsoft.Office.Tools.CustomTaskPane pane)
        {
            if (pane == null) return;
            registered.Remove(pane);
        }

        public static void OnPaneVisibilityChanged()
        {
            try
            {
                // If Color Palette exists, ensure it is left-most among right-docked panes
                var cp = Ribbon.ColorPaletteInstance;
                if (cp != null && !cp.IsDisposed)
                {
                    PaneOrdering.EnsureColorPaletteLeftMost(cp);
                }
            }
            catch { }
        }

        public static IReadOnlyList<Microsoft.Office.Tools.CustomTaskPane> CurrentPanes => registered.AsReadOnly();
    }
} 