using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PowerPointAddIn.Constants;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// PowerPoint operations service implementation
    /// </summary>
    public class PowerPointService : IPowerPointService
    {
        private readonly IErrorHandlerService _errorHandler;

        public PowerPointService(IErrorHandlerService errorHandler)
        {
            _errorHandler = errorHandler ?? throw new ArgumentNullException(nameof(errorHandler));
        }

        /// <summary>
        /// Gets the current active presentation
        /// </summary>
        public PowerPoint.Presentation ActivePresentation
        {
            get
            {
                try
                {
                    return my_addin.Globals.ThisAddIn.Application?.ActivePresentation;
                }
                catch (Exception ex)
                {
                    _errorHandler.LogError(ex, "Getting active presentation");
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the current active slide
        /// </summary>
        public PowerPoint.Slide ActiveSlide
        {
            get
            {
                try
                {
                    var app = my_addin.Globals.ThisAddIn?.Application;
                    if (app == null || app.ActivePresentation == null || app.ActiveWindow == null)
                        return null;

                    var selection = app.ActiveWindow.Selection;
                    if (selection != null && selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides && selection.SlideRange != null && selection.SlideRange.Count > 0)
                        return selection.SlideRange[1];

                    var viewSlide = app.ActiveWindow.View?.Slide;
                    if (viewSlide != null)
                        return viewSlide;

                    return app.ActivePresentation.Slides?.Count > 0 ? app.ActivePresentation.Slides[1] : null;
                }
                catch (Exception ex)
                {
                    _errorHandler.LogError(ex, "Getting active slide");
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the currently selected shapes
        /// </summary>
        public PowerPoint.ShapeRange SelectedShapes
        {
            get
            {
                try
                {
                    var selection = my_addin.Globals.ThisAddIn.Application?.ActiveWindow?.Selection;
                    if (selection?.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        return selection.ShapeRange;
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    _errorHandler.LogError(ex, "Getting selected shapes");
                    return null;
                }
            }
        }

        /// <summary>
        /// Creates a new presentation
        /// </summary>
        public PowerPoint.Presentation CreateNewPresentation()
        {
            return ExecuteWithErrorHandling(() =>
            {
                var app = my_addin.Globals.ThisAddIn.Application;
                return app.Presentations.Add();
            }, "creating new presentation");
        }

        /// <summary>
        /// Opens an existing presentation
        /// </summary>
        public PowerPoint.Presentation OpenPresentation(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be empty", nameof(filePath));

            return ExecuteWithErrorHandling(() =>
            {
                var app = my_addin.Globals.ThisAddIn.Application;
                return app.Presentations.Open(filePath);
            }, "opening presentation");
        }

        /// <summary>
        /// Saves the active presentation
        /// </summary>
        public void SavePresentation(string filePath = null)
        {
            ExecuteWithErrorHandling(() =>
            {
                var presentation = ActivePresentation;
                if (presentation == null)
                {
                    _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                    return;
                }

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    presentation.Save();
                }
                else
                {
                    presentation.SaveAs(filePath);
                }
            }, "saving presentation");
        }

        /// <summary>
        /// Gets all shapes from the active slide
        /// </summary>
        public IEnumerable<PowerPoint.Shape> GetActiveSlideShapes()
        {
            try
            {
                var slide = ActiveSlide;
                if (slide?.Shapes != null)
                {
                    return slide.Shapes.Cast<PowerPoint.Shape>();
                }
                return Enumerable.Empty<PowerPoint.Shape>();
            }
            catch (Exception ex)
            {
                _errorHandler.LogError(ex, "Getting active slide shapes");
                return Enumerable.Empty<PowerPoint.Shape>();
            }
        }

        /// <summary>
        /// Checks if shapes are currently selected
        /// </summary>
        public bool HasSelectedShapes()
        {
            try
            {
                var selection = my_addin.Globals.ThisAddIn.Application?.ActiveWindow?.Selection;
                return selection?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && 
                       selection.ShapeRange?.Count > 0;
            }
            catch (Exception ex)
            {
                _errorHandler.LogError(ex, "Checking selected shapes");
                return false;
            }
        }

        /// <summary>
        /// Checks if a presentation is currently active
        /// </summary>
        public bool HasActivePresentation()
        {
            try
            {
                return ActivePresentation != null;
            }
            catch (Exception ex)
            {
                _errorHandler.LogError(ex, "Checking active presentation");
                return false;
            }
        }

        /// <summary>
        /// Gets the PowerPoint application instance
        /// </summary>
        public PowerPoint.Application GetApplication()
        {
            try
            {
                return my_addin.Globals.ThisAddIn.Application;
            }
            catch (Exception ex)
            {
                _errorHandler.LogError(ex, "Getting PowerPoint application");
                return null;
            }
        }

        /// <summary>
        /// Validates minimum shape selection for operations
        /// </summary>
        public bool ValidateShapeSelection(int minimumShapes = 1, string operationName = "operation")
        {
            if (!HasActivePresentation())
            {
                _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                return false;
            }

            if (!HasSelectedShapes())
            {
                _errorHandler.ShowWarning(AppConstants.Messages.NoSelectedShapes);
                return false;
            }

            var shapeCount = SelectedShapes?.Count ?? 0;
            if (shapeCount < minimumShapes)
            {
                var message = string.Format(AppConstants.Messages.InsufficientShapes, minimumShapes);
                _errorHandler.ShowWarning(message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Exports presentation to PDF
        /// </summary>
        public bool ExportToPdf(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be empty", nameof(filePath));

            return ExecuteWithErrorHandling(() =>
            {
                var presentation = ActivePresentation;
                if (presentation == null)
                {
                    _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                    return false;
                }

                presentation.ExportAsFixedFormat(filePath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                return true;
            }, "exporting to PDF", false);
        }

        /// <summary>
        /// Prints the active presentation
        /// </summary>
        public bool PrintPresentation()
        {
            return ExecuteWithErrorHandling(() =>
            {
                var presentation = ActivePresentation;
                if (presentation == null)
                {
                    _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                    return false;
                }

                presentation.PrintOut();
                return true;
            }, "printing presentation", false);
        }

        /// <summary>
        /// Helper method for executing actions with error handling
        /// </summary>
        private void ExecuteWithErrorHandling(Action action, string operationName)
        {
            _errorHandler.ExecuteWithErrorHandling(action, operationName);
        }

        /// <summary>
        /// Helper method for executing functions with error handling
        /// </summary>
        private T ExecuteWithErrorHandling<T>(Func<T> function, string operationName, T defaultValue = default(T))
        {
            return _errorHandler.ExecuteWithErrorHandling(function, operationName, null, defaultValue);
        }
    }
} 