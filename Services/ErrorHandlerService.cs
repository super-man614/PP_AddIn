using System;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPointAddIn.Constants;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Centralized error handling service implementation
    /// </summary>
    public class ErrorHandlerService : IErrorHandlerService
    {
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

            // TODO: Add file logging or other logging mechanisms
        }

        /// <summary>
        /// Shows warning message to user
        /// </summary>
        public void ShowWarning(string message, string title = "Warning")
        {
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
    }
} 