using System;
using System.Collections.Specialized;
using System.Globalization;
using System.Text;
using my_addin.Properties;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace my_addin.Core
{
    public static class ShapePresetStorage
    {
        public static bool SavePreset(int presetIndex)
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange == null || selection.ShapeRange.Count < 1)
                {
                    return false;
                }

                var shape = selection.ShapeRange[1];
                string serialized = SerializeShapeFormat(shape);
                SetPresetString(presetIndex, serialized);
                Settings.Default.Save();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool ApplyPreset(int presetIndex)
        {
            try
            {
                string data = GetPresetString(presetIndex);
                if (string.IsNullOrWhiteSpace(data))
                    return false;

                var app = Globals.ThisAddIn?.Application;
                var selection = app?.ActiveWindow?.Selection;
                if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange == null || selection.ShapeRange.Count < 1)
                {
                    return false;
                }

                var preset = ShapeFormatPreset.Deserialize(data);
                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    ApplyFormat(selection.ShapeRange[i], preset);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void ClearPreset(int presetIndex)
        {
            SetPresetString(presetIndex, string.Empty);
            Settings.Default.Save();
        }

        private static string GetPresetString(int presetIndex)
        {
            switch (presetIndex)
            {
                case 1: return Settings.Default.Preset1;
                case 2: return Settings.Default.Preset2;
                case 3: return Settings.Default.Preset3;
                default: return string.Empty;
            }
        }

        private static void SetPresetString(int presetIndex, string value)
        {
            switch (presetIndex)
            {
                case 1: Settings.Default.Preset1 = value; break;
                case 2: Settings.Default.Preset2 = value; break;
                case 3: Settings.Default.Preset3 = value; break;
            }
        }

        private static string SerializeShapeFormat(PowerPoint.Shape shape)
        {
            var preset = new ShapeFormatPreset();

            try
            {
                // Text formatting only (not content)
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var tr = shape.TextFrame.TextRange;
                    // Don't save text content: preset.Text = tr?.Text ?? string.Empty;
                    if (tr?.Font != null)
                    {
                        preset.FontName = tr.Font.Name;
                        preset.FontSize = (float?)tr.Font.Size;
                        preset.FontColorRGB = (int?)tr.Font.Color?.RGB;
                        preset.Bold = TryTriToBool(tr.Font.Bold);
                        preset.Italic = TryTriToBool(tr.Font.Italic);
                        preset.Underline = TryTriToBool(tr.Font.Underline);
                    }

                    // Word wrap (TextFrame2 for wrap)
                    preset.WordWrap = shape.TextFrame2.WordWrap == Office.MsoTriState.msoTrue;
                }

                // Fill/Line
                if (shape.Fill != null && shape.Fill.ForeColor != null)
                {
                    preset.FillRGB = shape.Fill.ForeColor.RGB;
                }
                if (shape.Line != null)
                {
                    if (shape.Line.ForeColor != null)
                        preset.LineRGB = shape.Line.ForeColor.RGB;
                    try { preset.LineWeight = (float?)shape.Line.Weight; } catch { }
                }
            }
            catch { }

            return preset.Serialize();
        }

        private static void ApplyFormat(PowerPoint.Shape shape, ShapeFormatPreset preset)
        {
            try
            {
                // Text formatting only (not content)
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var tr = shape.TextFrame.TextRange;
                    // Don't apply text content: tr.Text = preset.Text;
                    if (!string.IsNullOrEmpty(preset.FontName))
                        tr.Font.Name = preset.FontName;
                    if (preset.FontSize.HasValue)
                        tr.Font.Size = preset.FontSize.Value;
                    if (preset.FontColorRGB.HasValue)
                        tr.Font.Color.RGB = preset.FontColorRGB.Value;
                    if (preset.Bold.HasValue)
                        tr.Font.Bold = preset.Bold.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    if (preset.Italic.HasValue)
                        tr.Font.Italic = preset.Italic.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    if (preset.Underline.HasValue)
                        tr.Font.Underline = preset.Underline.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;

                    if (preset.WordWrap.HasValue)
                        shape.TextFrame2.WordWrap = preset.WordWrap.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                }

                // Fill/Line
                if (preset.FillRGB.HasValue)
                {
                    shape.Fill.Visible = Office.MsoTriState.msoTrue;
                    shape.Fill.ForeColor.RGB = preset.FillRGB.Value;
                }
                if (preset.LineRGB.HasValue)
                {
                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = preset.LineRGB.Value;
                }
                if (preset.LineWeight.HasValue)
                {
                    shape.Line.Weight = preset.LineWeight.Value;
                }
            }
            catch { }
        }

        private static bool? TryTriToBool(Office.MsoTriState tri)
        {
            if (tri == Office.MsoTriState.msoTrue) return true;
            if (tri == Office.MsoTriState.msoFalse) return false;
            return null;
        }

        private class ShapeFormatPreset
        {
            // Removed Text property - we only save formatting, not content
            public int? FillRGB { get; set; }
            public int? LineRGB { get; set; }
            public float? LineWeight { get; set; }
            public string FontName { get; set; }
            public float? FontSize { get; set; }
            public int? FontColorRGB { get; set; }
            public bool? Bold { get; set; }
            public bool? Italic { get; set; }
            public bool? Underline { get; set; }
            public bool? WordWrap { get; set; }

            public string Serialize()
            {
                var sb = new StringBuilder();
                // Removed Text serialization - we only save formatting
                sb.Append("Fill=").Append(FillRGB?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append(';');
                sb.Append("Line=").Append(LineRGB?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append(';');
                sb.Append("LineW=").Append(LineWeight?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append(';');
                sb.Append("Font=").Append(FontName ?? string.Empty).Append(';');
                sb.Append("FontSize=").Append(FontSize?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append(';');
                sb.Append("FontColor=").Append(FontColorRGB?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append(';');
                sb.Append("Bold=").Append(BoolToStr(Bold)).Append(';');
                sb.Append("Italic=").Append(BoolToStr(Italic)).Append(';');
                sb.Append("Underline=").Append(BoolToStr(Underline)).Append(';');
                sb.Append("Wrap=").Append(BoolToStr(WordWrap)).Append(';');
                return sb.ToString();
            }

            public static ShapeFormatPreset Deserialize(string s)
            {
                var preset = new ShapeFormatPreset();
                if (string.IsNullOrEmpty(s)) return preset;
                var parts = s.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    var kv = part.Split(new[] { '=' }, 2);
                    if (kv.Length != 2) continue;
                    var key = kv[0];
                    var val = kv[1];
                    switch (key)
                    {
                        // Removed Text case - we only save formatting
                        case "Fill": preset.FillRGB = TryParseInt(val); break;
                        case "Line": preset.LineRGB = TryParseInt(val); break;
                        case "LineW": preset.LineWeight = TryParseFloat(val); break;
                        case "Font": preset.FontName = val; break;
                        case "FontSize": preset.FontSize = TryParseFloat(val); break;
                        case "FontColor": preset.FontColorRGB = TryParseInt(val); break;
                        case "Bold": preset.Bold = TryParseBool(val); break;
                        case "Italic": preset.Italic = TryParseBool(val); break;
                        case "Underline": preset.Underline = TryParseBool(val); break;
                        case "Wrap": preset.WordWrap = TryParseBool(val); break;
                    }
                }
                return preset;
            }

            private static string ToB64(string s)
            {
                if (string.IsNullOrEmpty(s)) return string.Empty;
                return Convert.ToBase64String(Encoding.UTF8.GetBytes(s));
            }

            private static string FromB64(string s)
            {
                if (string.IsNullOrEmpty(s)) return string.Empty;
                try { return Encoding.UTF8.GetString(Convert.FromBase64String(s)); } catch { return string.Empty; }
            }

            private static int? TryParseInt(string s)
            {
                if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)) return v;
                return null;
            }

            private static float? TryParseFloat(string s)
            {
                if (float.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return v;
                return null;
            }

            private static bool? TryParseBool(string s)
            {
                if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)) return true;
                if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)) return false;
                return null;
            }

            private static string BoolToStr(bool? b)
            {
                return b.HasValue ? (b.Value ? "true" : "false") : string.Empty;
            }
        }
    }
} 