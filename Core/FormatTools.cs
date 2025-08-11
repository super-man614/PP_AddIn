using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace my_addin.Core
{
    public static class FormatTools
    {
        /// <summary>
        /// Matches the width of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchWidth()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master
                float targetWidth = masterShape.Width;

                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    shapes[i].Width = targetWidth;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching width: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches the height of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchHeight()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master
                float targetHeight = masterShape.Height;

                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    shapes[i].Height = targetHeight;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching height: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches both width and height of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchSize()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master
                float targetWidth = masterShape.Width;
                float targetHeight = masterShape.Height;

                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    shapes[i].Width = targetWidth;
                    shapes[i].Height = targetHeight;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching size: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches the fill color of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchFill()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master

                // Get master fill color
                int? masterFillColor = null;
                try
                {
                    if (masterShape.Fill?.ForeColor != null)
                    {
                        masterFillColor = masterShape.Fill.ForeColor.RGB;
                    }
                }
                catch { }

                if (!masterFillColor.HasValue)
                    return;

                // Apply to all other shapes
                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    try
                    {
                        var shape = shapes[i];
                        if (shape.Fill != null)
                        {
                            shape.Fill.Visible = Office.MsoTriState.msoTrue;
                            shape.Fill.ForeColor.RGB = masterFillColor.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error applying fill to shape {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching fill: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches the font color of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchFontColor()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master

                // Get master font color
                int? masterFontColor = null;
                try
                {
                    if (masterShape.HasTextFrame == Office.MsoTriState.msoTrue && 
                        masterShape.TextFrame?.TextRange?.Font?.Color != null)
                    {
                        masterFontColor = masterShape.TextFrame.TextRange.Font.Color.RGB;
                    }
                }
                catch { }

                if (!masterFontColor.HasValue)
                    return;

                // Apply to all other shapes
                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    try
                    {
                        var shape = shapes[i];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && 
                            shape.TextFrame?.TextRange?.Font != null)
                        {
                            shape.TextFrame.TextRange.Font.Color.RGB = masterFontColor.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error applying font color to shape {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching font color: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches the outline color of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchOutline()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master

                // Get master outline color
                int? masterOutlineColor = null;
                try
                {
                    if (masterShape.Line?.ForeColor != null)
                    {
                        masterOutlineColor = masterShape.Line.ForeColor.RGB;
                    }
                }
                catch { }

                if (!masterOutlineColor.HasValue)
                    return;

                // Apply to all other shapes
                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    try
                    {
                        var shape = shapes[i];
                        if (shape.Line != null)
                        {
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.ForeColor.RGB = masterOutlineColor.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error applying outline to shape {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching outline: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches the font size of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchFontSize()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master

                // Get master font size
                float? masterFontSize = null;
                try
                {
                    if (masterShape.HasTextFrame == Office.MsoTriState.msoTrue && 
                        masterShape.TextFrame?.TextRange?.Font != null)
                    {
                        masterFontSize = masterShape.TextFrame.TextRange.Font.Size;
                    }
                }
                catch { }

                if (!masterFontSize.HasValue)
                    return;

                // Apply to all other shapes
                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    try
                    {
                        var shape = shapes[i];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && 
                            shape.TextFrame?.TextRange?.Font != null)
                        {
                            shape.TextFrame.TextRange.Font.Size = masterFontSize.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error applying font size to shape {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching font size: {ex.Message}");
            }
        }

        /// <summary>
        /// Matches the font family of all selected shapes to the last selected shape
        /// </summary>
        public static void MatchFontFamily()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange?.Count < 2)
                    return;

                var shapes = selection.ShapeRange;
                var masterShape = shapes[shapes.Count]; // Last selected is master

                // Get master font family
                string masterFontFamily = null;
                try
                {
                    if (masterShape.HasTextFrame == Office.MsoTriState.msoTrue && 
                        masterShape.TextFrame?.TextRange?.Font != null)
                    {
                        masterFontFamily = masterShape.TextFrame.TextRange.Font.Name;
                    }
                }
                catch { }

                if (string.IsNullOrEmpty(masterFontFamily))
                    return;

                // Apply to all other shapes
                for (int i = 1; i <= shapes.Count - 1; i++) // Skip the last one (master)
                {
                    try
                    {
                        var shape = shapes[i];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && 
                            shape.TextFrame?.TextRange?.Font != null)
                        {
                            shape.TextFrame.TextRange.Font.Name = masterFontFamily;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error applying font family to shape {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error matching font family: {ex.Message}");
            }
        }
    }
} 