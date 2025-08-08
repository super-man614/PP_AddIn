using System;
using Office = Microsoft.Office.Core;

namespace my_addin.Core
{
    public static class PaneOrdering
    {
        // Ensure the Color Palette task pane is docked right and appears as the left-most among right-docked panes
        public static void EnsureColorPaletteLeftMost(ColorPaletteTaskPane colorPalette)
        {
            try
            {
                if (colorPalette == null || colorPalette.IsDisposed)
                    return;

                // Force right dock
                try { colorPalette.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }

                var panes = Globals.ThisAddIn?.CustomTaskPanes;
                if (panes == null) return;

                // Collect other right-docked visible panes
                var others = new System.Collections.Generic.List<Microsoft.Office.Tools.CustomTaskPane>();
                foreach (Microsoft.Office.Tools.CustomTaskPane p in panes)
                {
                    try
                    {
                        bool isPalette = object.ReferenceEquals(p, colorPalette.TaskPane);
                        if (!isPalette && p.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight && p.Visible)
                        {
                            others.Add(p);
                        }
                    }
                    catch { }
                }

                // Temporarily hide others to ensure palette is recreated closest to slide
                foreach (var p in others)
                {
                    try { p.Visible = false; } catch { }
                }

                // Make sure the palette is visible now (becomes left-most)
                try { colorPalette.Visible = true; } catch { }

                // Restore other panes
                foreach (var p in others)
                {
                    try { p.Visible = true; p.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PaneOrdering.EnsureColorPaletteLeftMost error: {ex.Message}");
            }
        }
    }
}
