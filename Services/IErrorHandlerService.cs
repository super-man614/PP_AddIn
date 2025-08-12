using System;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Interface for centralized error handling throughout the application
    /// </summary>
    public interface IErrorHandlerService
    {
        /// <summary>
        /// Handles exceptions with user-friendly error messages
        /// </summary>
        /// <param name="exception">The exception to handle</param>
        /// <param name="userMessage">User-friendly message to display</param>
        /// <param name="title">Error dialog title</param>
        void HandleError(Exception exception, string userMessage, string title = "Error");

        /// <summary>
        /// Logs error without showing UI
        /// </summary>
        /// <param name="exception">The exception to log</param>
        /// <param name="context">Context where error occurred</param>
        void LogError(Exception exception, string context);

        /// <summary>
        /// Logs informational messages
        /// </summary>
        /// <param name="message">Information message to log</param>
        void LogInfo(string message);

        /// <summary>
        /// Logs warning messages
        /// </summary>
        /// <param name="message">Warning message to log</param>
        void LogWarning(string message);

        /// <summary>
        /// Shows warning message to user
        /// </summary>
        /// <param name="message">Warning message</param>
        /// <param name="title">Warning dialog title</param>
        void ShowWarning(string message, string title = "Warning");

        /// <summary>
        /// Shows information message to user
        /// </summary>
        /// <param name="message">Information message</param>
        /// <param name="title">Information dialog title</param>
        void ShowInfo(string message, string title = "Information");

        /// <summary>
        /// Gets the current log file path
        /// </summary>
        /// <returns>Path to the current log file</returns>
        string GetLogFilePath();

        /// <summary>
        /// Cleans up old log files (keeps last 30 days)
        /// </summary>
        void CleanupOldLogs();

        /// <summary>
        /// Executes an action with error handling
        /// </summary>
        /// <param name="action">Action to execute</param>
        /// <param name="operationName">Name of the operation for logging</param>
        /// <param name="userMessage">Optional user message</param>
        /// <returns>True if successful, false if error occurred</returns>
        bool ExecuteWithErrorHandling(Action action, string operationName, string userMessage = null);

        /// <summary>
        /// Executes a function with error handling
        /// </summary>
        /// <typeparam name="T">Return type</typeparam>
        /// <param name="function">Function to execute</param>
        /// <param name="operationName">Name of the operation for logging</param>
        /// <param name="userMessage">Optional user message</param>
        /// <param name="defaultValue">Default value to return on error</param>
        /// <returns>Function result or default value on error</returns>
        T ExecuteWithErrorHandling<T>(Func<T> function, string operationName, string userMessage = null, T defaultValue = default(T));
    }
} 