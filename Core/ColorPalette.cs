using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace my_addin.Core
{
    public class ColorPaletteDefinition
    {
        public string Name { get; set; } = "Default";
        public List<string> Colors { get; set; } = new List<string>(); // hex strings like #RRGGBB

        public static ColorPaletteDefinition Default()
        {
            return new ColorPaletteDefinition
            {
                Name = "Template Default",
                Colors = new List<string>
                {
                    "#07396F","#235490","#4A6C9E","#6C86AE","#95A7C5",
                    "#0B6C6C","#129696","#17BEBE","#7AD1CF","#CDEDEA",
                    "#333333","#666666","#999999","#C0C0C0","#E6E6E6",
                    "#A61C00","#E84C3D","#F1948A","#F5B7B1","#FADBD8",
                    "#E67E22","#F39C12","#F5CBA7","#F8C471","#FDEBD0",
                    "#27AE60","#2ECC71","#58D68D","#82E0AA","#ABEBC6"
                }
            };
        }

        public static Color FromHex(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return Color.Black;
            if (hex.StartsWith("#")) hex = hex.Substring(1);
            if (hex.Length == 6)
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                return Color.FromArgb(r, g, b);
            }
            return Color.Black;
        }

        public static string ToHex(Color c)
        {
            return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
        }
    }

    public static class ColorPaletteStorage
    {
        private const string FileName = "color_palette.json";

        public static string GetStoragePath()
        {
            string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string addinDir = Path.Combine(baseDir, "my-addin");
            try { if (!Directory.Exists(addinDir)) Directory.CreateDirectory(addinDir); } catch { }
            return Path.Combine(addinDir, FileName);
        }

        public static ColorPaletteDefinition LoadOrDefault()
        {
            try
            {
                var path = GetStoragePath();
                if (File.Exists(path))
                {
                    var lines = File.ReadAllLines(path)
                                   .Select(l => l.Trim())
                                   .Where(l => !string.IsNullOrWhiteSpace(l) && !l.StartsWith("# "))
                                   .ToList();
                    if (lines.Count > 0)
                    {
                        return new ColorPaletteDefinition { Name = "User", Colors = lines };
                    }
                }
            }
            catch { }
            return ColorPaletteDefinition.Default();
        }

        public static void Save(ColorPaletteDefinition palette)
        {
            try
            {
                var lines = new List<string>();
                lines.Add("# Simple palette file: one hex color per line");
                lines.AddRange(palette.Colors);
                File.WriteAllLines(GetStoragePath(), lines);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save color palette: {ex.Message}");
            }
        }

        public static void ResetToTemplate()
        {
            Save(ColorPaletteDefinition.Default());
        }
    }
}

