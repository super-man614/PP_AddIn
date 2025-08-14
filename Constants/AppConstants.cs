using System;
using System.IO;
using System.Collections.Generic;
using PowerPointAddIn.Models;

namespace PowerPointAddIn.Constants
{
    /// <summary>
    /// Application-wide constants and configuration
    /// </summary>
    public static class AppConstants
    {
        // Application Information
        public const string APP_NAME = "PowerPoint Add-in Tools";
        public const string APP_VERSION = "1.0.0.3";
        public const string APP_PUBLISHER = "PowerPoint Add-in Tools";
        
        // File Paths
        public static readonly string USER_APPDATA_PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        public static readonly string ADDIN_DATA_PATH = Path.Combine(USER_APPDATA_PATH, "PowerPointAddIn");
        public static readonly string LOGS_PATH = Path.Combine(ADDIN_DATA_PATH, "Logs");
        public static readonly string CONFIG_PATH = Path.Combine(ADDIN_DATA_PATH, "Config");
        public static readonly string TEMP_PATH = Path.GetTempPath();
        
        // Logging Configuration
        public const int LOG_RETENTION_DAYS = 30;
        public const int MAX_LOG_FILE_SIZE_MB = 10;
        public const bool ENABLE_DEBUG_LOGGING = true;
        
        // Office Integration
        public const string OFFICE_APP_NAME = "POWERPNT.EXE";
        public const string OFFICE_ROOT_PATH = @"C:\\Program Files\\Microsoft Office\\root\\Office16";
        public const string OFFICE_ALTERNATE_PATH = @"C:\\Program Files\\Microsoft Office\\Office16";
        
        // UI Configuration
        public const int DEFAULT_TASKPANE_WIDTH = 300;
        public const int DEFAULT_TASKPANE_HEIGHT = 600;
        public const bool AUTO_SHOW_TASKPANES = true;
        public const bool REMEMBER_TASKPANE_POSITIONS = true;
        
        // Matrix Configuration
        public const string MATRIX_TAG_PREFIX = "MATRIX";
        public const string MATRIX_ROW_TAG = "MATRIX_ROW";
        public const string MATRIX_COL_TAG = "MATRIX_COL";
        public const string MATRIX_DEFAULT_TEXT = "XXXX";
        public const int MATRIX_MAX_ROWS = 100;
        public const int MATRIX_MAX_COLS = 100;
        
        // Color Palette Configuration
        public const int COLOR_PALETTE_MAX_COLORS = 50;
        public const int COLOR_PALETTE_DEFAULT_SIZE = 24;
        public const bool COLOR_PALETTE_AUTO_SAVE = true;
        
        // Error Handling
        public const int MAX_ERROR_RETRY_ATTEMPTS = 3;
        public const int ERROR_DIALOG_TIMEOUT_MS = 5000;
        public const bool SHOW_DETAILED_ERRORS = false;
        
        // Performance
        public const int MAX_OBJECTS_PER_OPERATION = 1000;
        public const bool ENABLE_OPERATION_CANCELLATION = true;
        public const int OPERATION_TIMEOUT_MS = 30000;
        
        // File Operations
        public static readonly string[] SUPPORTED_IMAGE_FORMATS = { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico" };
        public static readonly string[] SUPPORTED_DOCUMENT_FORMATS = { ".pptx", ".ppt", ".pdf", ".html" };
        public const int MAX_FILE_SIZE_MB = 100;
        
        // Network and Updates
        public const bool CHECK_FOR_UPDATES = true;
        public const int UPDATE_CHECK_INTERVAL_DAYS = 7;
        public const string UPDATE_CHECK_URL = "https://api.example.com/updates";

        // User messages used across services
        public static class Messages
        {
            public const string NoActivePresentation = "No active presentation. Please open or create a presentation first.";
            public const string NoSelectedShapes = "Please select one or more shapes to perform this action.";
            public const string InsufficientShapes = "Please select at least {0} shapes to perform this action.";
            public const string SuccessTitle = "Success";
        }

        // Measurement-related constants
        public static class Measurements
        {
            public const float PointsPerInch = 72f;
            public const double ClosestPresetTolerance = 0.5; // inches
        }

        // Common slide size presets mapped by type
        public static readonly Dictionary<SlideSizeType, SlideSizeInfo> SlideSizePresets = new Dictionary<SlideSizeType, SlideSizeInfo>
        {
            { SlideSizeType.Standard,         new SlideSizeInfo(10m,    7.5m,  "4:3 Standard",       SlideSizeType.Standard) },
            { SlideSizeType.Widescreen16x9,    new SlideSizeInfo(13.3m,  7.5m,  "16:9 Widescreen",     SlideSizeType.Widescreen16x9) },
            { SlideSizeType.Widescreen16x10,   new SlideSizeInfo(12.8m,  8m,    "16:10 Widescreen",    SlideSizeType.Widescreen16x10) },
            { SlideSizeType.A4Portrait,        new SlideSizeInfo(8.27m, 11.69m, "A4 Portrait",         SlideSizeType.A4Portrait) },
            { SlideSizeType.A4Landscape,       new SlideSizeInfo(11.69m, 8.27m, "A4 Landscape",        SlideSizeType.A4Landscape) },
            { SlideSizeType.A3Portrait,        new SlideSizeInfo(11.69m, 16.54m, "A3 Portrait",        SlideSizeType.A3Portrait) },
            { SlideSizeType.A3Landscape,       new SlideSizeInfo(16.54m, 11.69m, "A3 Landscape",       SlideSizeType.A3Landscape) },
            { SlideSizeType.LetterPortrait,    new SlideSizeInfo(8.5m,  11m,    "Letter Portrait",     SlideSizeType.LetterPortrait) },
            { SlideSizeType.LetterLandscape,   new SlideSizeInfo(11m,    8.5m,  "Letter Landscape",    SlideSizeType.LetterLandscape) },
            { SlideSizeType.Banner,            new SlideSizeInfo(8m,     1m,    "Banner",              SlideSizeType.Banner) },
            { SlideSizeType.SocialMedia,       new SlideSizeInfo(1.91m,  1m,    "Social Media",        SlideSizeType.SocialMedia) },
            { SlideSizeType.InstagramPost,     new SlideSizeInfo(1m,     1m,    "Instagram Post",      SlideSizeType.InstagramPost) },
            { SlideSizeType.InstagramStory,    new SlideSizeInfo(1m,     1.78m, "Instagram Story",     SlideSizeType.InstagramStory) },
            { SlideSizeType.YouTubeThumbnail,  new SlideSizeInfo(1.78m,  1m,    "YouTube Thumbnail",   SlideSizeType.YouTubeThumbnail) },
            { SlideSizeType.LinkedInPost,      new SlideSizeInfo(1.91m,  1m,    "LinkedIn Post",       SlideSizeType.LinkedInPost) },
            { SlideSizeType.TwitterPost,       new SlideSizeInfo(1.91m,  1m,    "Twitter Post",        SlideSizeType.TwitterPost) },
            { SlideSizeType.FacebookPost,      new SlideSizeInfo(1.91m,  1m,    "Facebook Post",       SlideSizeType.FacebookPost) }
        };
        
        /// <summary>
        /// Gets the Office executable path, trying multiple possible locations
        /// </summary>
        public static string GetOfficeExecutablePath()
        {
            // Try the configured path first
            string primaryPath = Path.Combine(OFFICE_ROOT_PATH, OFFICE_APP_NAME);
            if (File.Exists(primaryPath))
                return primaryPath;
            
            // Try alternate path
            string alternatePath = Path.Combine(OFFICE_ALTERNATE_PATH, OFFICE_APP_NAME);
            if (File.Exists(alternatePath))
                return alternatePath;
            
            // Try to find in Program Files
            string[] possiblePaths = {
                @"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
                @"C:\\Program Files\\Microsoft Office\\Office16\\POWERPNT.EXE",
                @"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
                @"C:\\Program Files (x86)\\Microsoft Office\\Office16\\POWERPNT.EXE",
                @"C:\\Program Files\\Microsoft Office\\root\\Office15\\POWERPNT.EXE",
                @"C:\\Program Files\\Microsoft Office\\Office15\\POWERPNT.EXE"
            };
            
            foreach (string path in possiblePaths)
            {
                if (File.Exists(path))
                    return path;
            }
            
            // Return default if none found
            return primaryPath;
        }
        
        /// <summary>
        /// Ensures all required directories exist
        /// </summary>
        public static void EnsureDirectoriesExist()
        {
            try
            {
                Directory.CreateDirectory(ADDIN_DATA_PATH);
                Directory.CreateDirectory(LOGS_PATH);
                Directory.CreateDirectory(CONFIG_PATH);
            }
            catch (Exception ex)
            {
                // Log error but don't throw - this is called during startup
                System.Diagnostics.Debug.WriteLine($"Failed to create directories: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Gets the current Office version string
        /// </summary>
        public static string GetOfficeVersionString()
        {
            try
            {
                string exePath = GetOfficeExecutablePath();
                if (File.Exists(exePath))
                {
                    var versionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(exePath);
                    return versionInfo.FileVersion ?? "Unknown";
                }
            }
            catch
            {
                // Ignore errors
            }
            return "Unknown";
        }
    }
} 