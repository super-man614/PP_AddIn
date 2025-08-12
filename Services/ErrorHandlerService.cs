using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Text;
using PowerPointAddIn.Constants;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Centralized error handling service implementation
    /// </summary>
    public class ErrorHandlerService : IErrorHandlerService
    {
        private readonly string _logFilePath;
        private readonly object _logLock = new object();

        public ErrorHandlerService()
        {
            try
            {
                // Create logs directory in user's AppData folder
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logsDir = Path.Combine(appDataPath, "PowerPointAddIn", "Logs");
                Directory.CreateDirectory(logsDir);
                
                // Create log file with date stamp
                string logFileName = $"PowerPointAddIn_{DateTime.Now:yyyy-MM-dd}.log";
                _logFilePath = Path.Combine(logsDir, logFileName);
            }
            catch
            {
                // Fallback to temp directory if AppData fails
                _logFilePath = Path.Combine(Path.GetTempPath(), "PowerPointAddIn.log");
            }
        }

        /// <summary>
        /// Handles exceptions with user-friendly error messages
        /// </summary>
        public void HandleError(Exception exception, string userMessage, string title = "Error")
        {
            // Log the error
            LogError(exception, userMessage);

            // Show user-friendly message
            MessageBox.Show(userMessage, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Logs error without showing UI
        /// </summary>
        public void LogError(Exception exception, string context)
        {
            var errorMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR in {context}: {exception.Message}";
            if (exception.InnerException != null)
            {
                errorMessage += $" | Inner: {exception.InnerException.Message}";
            }
            errorMessage += $" | Stack: {exception.StackTrace}";

            // Write to debug output
            Debug.WriteLine(errorMessage);

            // Write to file log
            WriteToLogFile(errorMessage);
        }

        /// <summary>
        /// Logs informational messages
        /// </summary>
        public void LogInfo(string message)
        {
            var infoMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {message}";
            Debug.WriteLine(infoMessage);
            WriteToLogFile(infoMessage);
        }

        /// <summary>
        /// Logs warning messages
        /// </summary>
        public void LogWarning(string message)
        {
            var warningMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] WARNING: {message}";
            Debug.WriteLine(warningMessage);
            WriteToLogFile(warningMessage);
        }

        /// <summary>
        /// Shows warning message to user
        /// </summary>
        public void ShowWarning(string message, string title = "Warning")
        {
            LogWarning(message);
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Shows information message to user
        /// </summary>
        public void ShowInfo(string message, string title = "Information")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Creates standardized error message for common scenarios
        /// </summary>
        public static string CreateErrorMessage(string operation, Exception exception = null)
        {
            var message = $"Error during {operation}.";
            
            if (exception != null)
            {
                // Add specific error details for common exceptions - C# 7.3 compatible
                if (exception is UnauthorizedAccessException)
                {
                    message += " Access denied. Please check permissions.";
                }
                else if (exception is System.IO.FileNotFoundException)
                {
                    message += " Required file not found.";
                }
                else if (exception is System.Runtime.InteropServices.COMException)
                {
                    message += " PowerPoint operation failed. Please ensure PowerPoint is running properly.";
                }
                else
                {
                    message += " Please try again or contact support if the problem persists.";
                }
            }

            return message;
        }

        /// <summary>
        /// Executes an action with error handling
        /// </summary>
        public bool ExecuteWithErrorHandling(Action action, string operationName, string userMessage = null)
        {
            try
            {
                action();
                return true;
            }
            catch (Exception ex)
            {
                var message = userMessage ?? CreateErrorMessage(operationName, ex);
                HandleError(ex, message);
                return false;
            }
        }

        /// <summary>
        /// Executes a function with error handling
        /// </summary>
        public T ExecuteWithErrorHandling<T>(Func<T> function, string operationName, string userMessage = null, T defaultValue = default(T))
        {
            try
            {
                return function();
            }
            catch (Exception ex)
            {
                var message = userMessage ?? CreateErrorMessage(operationName, ex);
                HandleError(ex, message);
                return defaultValue;
            }
        }

        /// <summary>
        /// Writes message to log file with thread safety
        /// </summary>
        private void WriteToLogFile(string message)
        {
            try
            {
                lock (_logLock)
                {
                    File.AppendAllText(_logFilePath, message + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                // If file logging fails, fall back to debug output only
                Debug.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the current log file path
        /// </summary>
        public string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// Cleans up old log files (keeps last 30 days)
        /// </summary>
        public void CleanupOldLogs()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logsDir = Path.Combine(appDataPath, "PowerPointAddIn", "Logs");
                
                if (!Directory.Exists(logsDir)) return;

                var cutoffDate = DateTime.Now.AddDays(-30);
                var logFiles = Directory.GetFiles(logsDir, "PowerPointAddIn_*.log");

                foreach (var logFile in logFiles)
                {
                    try
                    {
                        var fileInfo = new FileInfo(logFile);
                        if (fileInfo.CreationTime < cutoffDate)
                        {
                            File.Delete(logFile);
                        }
                    }
                    catch
                    {
                        // Ignore individual file deletion errors
                    }
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
} 