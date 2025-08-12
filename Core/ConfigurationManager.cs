using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using PowerPointAddIn.Constants;

namespace PowerPointAddIn.Core
{
    /// <summary>
    /// Manages application configuration and settings
    /// </summary>
    public class ConfigurationManager
    {
        private static ConfigurationManager _instance;
        private static readonly object _lock = new object();
        private readonly string _configFilePath;
        private AppConfiguration _currentConfig;

        public static ConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ConfigurationManager();
                        }
                    }
                }
                return _instance;
            }
        }

        private ConfigurationManager()
        {
            AppConstants.EnsureDirectoriesExist();
            _configFilePath = Path.Combine(AppConstants.CONFIG_PATH, "appsettings.json");
            _currentConfig = LoadConfiguration();
        }

        /// <summary>
        /// Gets the current configuration
        /// </summary>
        public AppConfiguration Configuration => _currentConfig;

        /// <summary>
        /// Loads configuration from file or creates default
        /// </summary>
        private AppConfiguration LoadConfiguration()
        {
            try
            {
                if (File.Exists(_configFilePath))
                {
                    string jsonContent = File.ReadAllText(_configFilePath);
                    var config = JsonSerializer.Deserialize<AppConfiguration>(jsonContent);
                    return config ?? CreateDefaultConfiguration();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load configuration: {ex.Message}");
            }

            return CreateDefaultConfiguration();
        }

        /// <summary>
        /// Creates default configuration
        /// </summary>
        private AppConfiguration CreateDefaultConfiguration()
        {
            return new AppConfiguration
            {
                Logging = new LoggingConfig
                {
                    EnableFileLogging = true,
                    EnableDebugLogging = AppConstants.ENABLE_DEBUG_LOGGING,
                    LogRetentionDays = AppConstants.LOG_RETENTION_DAYS,
                    MaxLogFileSizeMB = AppConstants.MAX_LOG_FILE_SIZE_MB
                },
                UI = new UIConfig
                {
                    DefaultTaskPaneWidth = AppConstants.DEFAULT_TASKPANE_WIDTH,
                    DefaultTaskPaneHeight = AppConstants.DEFAULT_TASKPANE_HEIGHT,
                    AutoShowTaskPanes = AppConstants.AUTO_SHOW_TASKPANES,
                    RememberTaskPanePositions = AppConstants.REMEMBER_TASKPANE_POSITIONS,
                    EnableTooltips = true,
                    TooltipDelay = 1000
                },
                Matrix = new MatrixConfig
                {
                    DefaultText = AppConstants.MATRIX_DEFAULT_TEXT,
                    MaxRows = AppConstants.MATRIX_MAX_ROWS,
                    MaxColumns = AppConstants.MATRIX_MAX_COLS,
                    AutoResizeCells = true,
                    DefaultCellSize = 60
                },
                ColorPalette = new ColorPaletteConfig
                {
                    MaxColors = AppConstants.COLOR_PALETTE_MAX_COLORS,
                    DefaultColorSize = AppConstants.COLOR_PALETTE_DEFAULT_SIZE,
                    AutoSave = AppConstants.COLOR_PALETTE_AUTO_SAVE,
                    EnableDragAndDrop = true
                },
                Performance = new PerformanceConfig
                {
                    MaxObjectsPerOperation = AppConstants.MAX_OBJECTS_PER_OPERATION,
                    EnableOperationCancellation = AppConstants.ENABLE_OPERATION_CANCELLATION,
                    OperationTimeoutMs = AppConstants.OPERATION_TIMEOUT_MS,
                    EnableAsyncOperations = true
                }
            };
        }

        /// <summary>
        /// Saves current configuration to file
        /// </summary>
        public void SaveConfiguration()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                string jsonContent = JsonSerializer.Serialize(_currentConfig, options);
                File.WriteAllText(_configFilePath, jsonContent);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save configuration: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates configuration and saves to file
        /// </summary>
        public void UpdateConfiguration(Action<AppConfiguration> updateAction)
        {
            updateAction?.Invoke(_currentConfig);
            SaveConfiguration();
        }

        /// <summary>
        /// Resets configuration to defaults
        /// </summary>
        public void ResetToDefaults()
        {
            _currentConfig = CreateDefaultConfiguration();
            SaveConfiguration();
        }
    }

    /// <summary>
    /// Application configuration model
    /// </summary>
    public class AppConfiguration
    {
        public LoggingConfig Logging { get; set; } = new LoggingConfig();
        public UIConfig UI { get; set; } = new UIConfig();
        public MatrixConfig Matrix { get; set; } = new MatrixConfig();
        public ColorPaletteConfig ColorPalette { get; set; } = new ColorPaletteConfig();
        public PerformanceConfig Performance { get; set; } = new PerformanceConfig();
    }

    public class LoggingConfig
    {
        public bool EnableFileLogging { get; set; } = true;
        public bool EnableDebugLogging { get; set; } = true;
        public int LogRetentionDays { get; set; } = 30;
        public int MaxLogFileSizeMB { get; set; } = 10;
    }

    public class UIConfig
    {
        public int DefaultTaskPaneWidth { get; set; } = 300;
        public int DefaultTaskPaneHeight { get; set; } = 600;
        public bool AutoShowTaskPanes { get; set; } = true;
        public bool RememberTaskPanePositions { get; set; } = true;
        public bool EnableTooltips { get; set; } = true;
        public int TooltipDelay { get; set; } = 1000;
    }

    public class MatrixConfig
    {
        public string DefaultText { get; set; } = "XXXX";
        public int MaxRows { get; set; } = 100;
        public int MaxColumns { get; set; } = 100;
        public bool AutoResizeCells { get; set; } = true;
        public int DefaultCellSize { get; set; } = 60;
    }

    public class ColorPaletteConfig
    {
        public int MaxColors { get; set; } = 50;
        public int DefaultColorSize { get; set; } = 24;
        public bool AutoSave { get; set; } = true;
        public bool EnableDragAndDrop { get; set; } = true;
    }

    public class PerformanceConfig
    {
        public int MaxObjectsPerOperation { get; set; } = 1000;
        public bool EnableOperationCancellation { get; set; } = true;
        public int OperationTimeoutMs { get; set; } = 30000;
        public bool EnableAsyncOperations { get; set; } = true;
    }
} 