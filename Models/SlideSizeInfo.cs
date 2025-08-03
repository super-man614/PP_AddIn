using System;

namespace PowerPointAddIn.Models
{
    /// <summary>
    /// Represents slide size information
    /// </summary>
    public class SlideSizeInfo
    {
        public decimal Width { get; set; }
        public decimal Height { get; set; }
        public string Name { get; set; }
        public SlideSizeType Type { get; set; }
        public bool ScaleContent { get; set; } = true;

        public SlideSizeInfo(decimal width, decimal height, string name, SlideSizeType type = SlideSizeType.Custom)
        {
            Width = width;
            Height = height;
            Name = name;
            Type = type;
        }

        public override string ToString() => $"{Name} ({Width}\" × {Height}\")";
    }

    /// <summary>
    /// Slide size suggestion with reasoning
    /// </summary>
    public class SlideSizeSuggestion
    {
        public SlideSizeInfo RecommendedSize { get; set; }
        public string Reasoning { get; set; }
        public double ConfidenceScore { get; set; }
        public ContentAnalysisResult AnalysisResult { get; set; }

        public SlideSizeSuggestion(SlideSizeInfo recommendedSize, string reasoning, double confidenceScore = 0.8)
        {
            RecommendedSize = recommendedSize;
            Reasoning = reasoning;
            ConfidenceScore = confidenceScore;
        }
    }

    /// <summary>
    /// Content analysis result for size suggestions
    /// </summary>
    public class ContentAnalysisResult
    {
        public int TotalSlides { get; set; }
        public int SlidesWithImages { get; set; }
        public int SlidesWithCharts { get; set; }
        public int SlidesWithTables { get; set; }
        public int SlidesWithText { get; set; }
        public int SlidesWithVideos { get; set; }
        public int SlidesWithSmartArt { get; set; }

        public double ImageRatio => TotalSlides > 0 ? (double)SlidesWithImages / TotalSlides : 0;
        public double ChartRatio => TotalSlides > 0 ? (double)SlidesWithCharts / TotalSlides : 0;
        public double TableRatio => TotalSlides > 0 ? (double)SlidesWithTables / TotalSlides : 0;
        public double TextRatio => TotalSlides > 0 ? (double)SlidesWithText / TotalSlides : 0;
        public double VideoRatio => TotalSlides > 0 ? (double)SlidesWithVideos / TotalSlides : 0;
        public double SmartArtRatio => TotalSlides > 0 ? (double)SlidesWithSmartArt / TotalSlides : 0;
    }

    /// <summary>
    /// Enumeration of slide size types
    /// </summary>
    public enum SlideSizeType
    {
        Standard,       // 4:3
        Widescreen16x9, // 16:9
        Widescreen16x10,// 16:10
        A4Portrait,
        A4Landscape,
        A3Portrait,
        A3Landscape,
        LetterPortrait,
        LetterLandscape,
        Banner,
        SocialMedia,
        InstagramPost,
        InstagramStory,
        YouTubeThumbnail,
        LinkedInPost,
        TwitterPost,
        FacebookPost,
        Custom
    }
} 