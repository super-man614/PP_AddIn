using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Interface for PowerPoint operations abstraction
    /// </summary>
    public interface IPowerPointService
    {
        /// <summary>
        /// Gets the current active presentation
        /// </summary>
        PowerPoint.Presentation ActivePresentation { get; }

        /// <summary>
        /// Gets the current active slide
        /// </summary>
        PowerPoint.Slide ActiveSlide { get; }

        /// <summary>
        /// Gets the currently selected shapes
        /// </summary>
        PowerPoint.ShapeRange SelectedShapes { get; }

        /// <summary>
        /// Creates a new presentation
        /// </summary>
        /// <returns>The new presentation</returns>
        PowerPoint.Presentation CreateNewPresentation();

        /// <summary>
        /// Opens an existing presentation
        /// </summary>
        /// <param name="filePath">Path to the presentation file</param>
        /// <returns>The opened presentation</returns>
        PowerPoint.Presentation OpenPresentation(string filePath);

        /// <summary>
        /// Saves the active presentation
        /// </summary>
        /// <param name="filePath">Optional file path for Save As</param>
        void SavePresentation(string filePath = null);

        /// <summary>
        /// Gets all shapes from the active slide
        /// </summary>
        /// <returns>Collection of shapes</returns>
        IEnumerable<PowerPoint.Shape> GetActiveSlideShapes();

        /// <summary>
        /// Checks if shapes are currently selected
        /// </summary>
        /// <returns>True if shapes are selected</returns>
        bool HasSelectedShapes();

        /// <summary>
        /// Checks if a presentation is currently active
        /// </summary>
        /// <returns>True if presentation is active</returns>
        bool HasActivePresentation();

        /// <summary>
        /// Validates minimum shape selection for operations
        /// </summary>
        /// <param name="minimumShapes">Minimum number of shapes required</param>
        /// <param name="operationName">Name of the operation for error messages</param>
        /// <returns>True if validation passes</returns>
        bool ValidateShapeSelection(int minimumShapes = 1, string operationName = "operation");

        /// <summary>
        /// Exports presentation to PDF
        /// </summary>
        /// <param name="filePath">Path for the PDF file</param>
        /// <returns>True if successful</returns>
        bool ExportToPdf(string filePath);

        /// <summary>
        /// Prints the active presentation
        /// </summary>
        /// <returns>True if successful</returns>
        bool PrintPresentation();
    }
} 