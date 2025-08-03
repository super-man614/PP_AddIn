using System.Collections.Generic;
using PowerPointAddIn.Models;

namespace PowerPointAddIn.Constants
{
    /// <summary>
    /// Application-wide constants
    /// </summary>
    public static class AppConstants
    {
        /// <summary>
        /// Application information
        /// </summary>
        public static class App
        {
            public const string Name = "PowerPoint Tools";
            public const string Version = "1.0.0";
            public const string Namespace = "PowerPointAddIn";
            public const string RibbonId = "PowerPointToolsTab";
        }

        /// <summary>
        /// Task pane configuration
        /// </summary>
        public static class TaskPane
        {
            public const string Title = "PowerPoint Tools";
            public const int DefaultWidth = 320;
            public const int MinWidth = 250;
            public const int MaxWidth = 500;
        }

        /// <summary>
        /// Error messages and titles
        /// </summary>
        public static class Messages
        {
            public const string ErrorTitle = "Error";
            public const string WarningTitle = "Warning";
            public const string InfoTitle = "Information";
            public const string SuccessTitle = "Success";

            public const string NoActivePresentation = "No active presentation found. Please open or create a presentation first.";
            public const string NoSelectedShapes = "Please select shapes to perform this operation.";
            public const string InsufficientShapes = "Please select at least {0} shapes for this operation.";
            public const string OperationSuccessful = "Operation completed successfully!";
            public const string OperationFailed = "Operation failed. Please try again.";
        }

        /// <summary>
        /// File operations
        /// </summary>
        public static class Files
        {
            public const string PowerPointFilter = "PowerPoint Files (*.pptx;*.ppt)|*.pptx;*.ppt";
            public const string PdfFilter = "PDF Files (*.pdf)|*.pdf";
            public const string DefaultExtension = ".pptx";
        }

        /// <summary>
        /// Image and icon paths
        /// </summary>
        public static class Icons
        {
            public const string BasePath = "icons";
            public const string PositionPath = "icons/position";
            public const string WizardPath = "icons/wizzards";
            public const string ElementsPath = "icons/elements";
            public const string FilePath = "icons/file";

            public const string ApplyIcon = "icons8-apply-64.png";
            public const string EnlargeIcon = "icons8-enlarge-50.png";
            public const string OpenFileIcon = "icons8-open-file-48.png";
        }

        /// <summary>
        /// Predefined slide size presets
        /// </summary>
        public static readonly Dictionary<SlideSizeType, SlideSizeInfo> SlideSizePresets = new Dictionary<SlideSizeType, SlideSizeInfo>
        {
            { SlideSizeType.Standard, new SlideSizeInfo(10m, 7.5m, "4:3 Standard", SlideSizeType.Standard) },
            { SlideSizeType.Widescreen16x9, new SlideSizeInfo(13.3m, 7.5m, "16:9 Widescreen", SlideSizeType.Widescreen16x9) },
            { SlideSizeType.Widescreen16x10, new SlideSizeInfo(12.8m, 8m, "16:10 Widescreen", SlideSizeType.Widescreen16x10) },
            { SlideSizeType.A4Portrait, new SlideSizeInfo(8.27m, 11.69m, "A4 Portrait", SlideSizeType.A4Portrait) },
            { SlideSizeType.A4Landscape, new SlideSizeInfo(11.69m, 8.27m, "A4 Landscape", SlideSizeType.A4Landscape) },
            { SlideSizeType.LetterPortrait, new SlideSizeInfo(8.5m, 11m, "Letter Portrait", SlideSizeType.LetterPortrait) },
            { SlideSizeType.LetterLandscape, new SlideSizeInfo(11m, 8.5m, "Letter Landscape", SlideSizeType.LetterLandscape) },
            { SlideSizeType.A3Portrait, new SlideSizeInfo(11.69m, 16.54m, "A3 Portrait", SlideSizeType.A3Portrait) },
            { SlideSizeType.A3Landscape, new SlideSizeInfo(16.54m, 11.69m, "A3 Landscape", SlideSizeType.A3Landscape) },
            { SlideSizeType.Banner, new SlideSizeInfo(8m, 1m, "Banner", SlideSizeType.Banner) },
            { SlideSizeType.SocialMedia, new SlideSizeInfo(1.91m, 1m, "Social Media", SlideSizeType.SocialMedia) },
            { SlideSizeType.InstagramPost, new SlideSizeInfo(1m, 1m, "Instagram Post", SlideSizeType.InstagramPost) },
            { SlideSizeType.InstagramStory, new SlideSizeInfo(1m, 1.78m, "Instagram Story", SlideSizeType.InstagramStory) },
            { SlideSizeType.YouTubeThumbnail, new SlideSizeInfo(1.78m, 1m, "YouTube Thumbnail", SlideSizeType.YouTubeThumbnail) },
            { SlideSizeType.LinkedInPost, new SlideSizeInfo(1.91m, 1m, "LinkedIn Post", SlideSizeType.LinkedInPost) },
            { SlideSizeType.TwitterPost, new SlideSizeInfo(1.91m, 1m, "Twitter Post", SlideSizeType.TwitterPost) },
            { SlideSizeType.FacebookPost, new SlideSizeInfo(1.91m, 1m, "Facebook Post", SlideSizeType.FacebookPost) }
        };

        /// <summary>
        /// Measurement constants
        /// </summary>
        public static class Measurements
        {
            public const float PointsPerInch = 72f;
            public const double ClosestPresetTolerance = 0.5;
            public const float DefaultScaleFactor = 1.0f;
            public const int DefaultButtonSize = 20;
            public const int LargeButtonSize = 25;
        }

        /// <summary>
        /// UI Configuration
        /// </summary>
        public static class UI
        {
            public const int TooltipDelay = 1000;
            public const int TooltipAutoPopDelay = 5000;
            public const int TooltipReshowDelay = 500;
            
            public static class Colors
            {
                public const string BorderColor = "Gray";
                public const string DividerColor = "#E1E1E1";
                public const string SectionHeaderColor = "Gray";
            }
        }
    }
} 