using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn.Models;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Interface for shape-related operations
    /// </summary>
    public interface IShapeService
    {
        /// <summary>
        /// Aligns selected shapes in a process chain
        /// </summary>
        /// <returns>True if successful</returns>
        bool AlignProcessChain();

        /// <summary>
        /// Aligns angles of selected shapes to master
        /// </summary>
        /// <returns>True if successful</returns>
        bool AlignAngles();

        /// <summary>
        /// Aligns selected objects to process arrow
        /// </summary>
        /// <returns>True if successful</returns>
        bool AlignToProcessArrow();

        /// <summary>
        /// Adjusts pentagon header boxes
        /// </summary>
        /// <returns>True if successful</returns>
        bool AdjustPentagonHeader();

        /// <summary>
        /// Aligns block arrows with standard style
        /// </summary>
        /// <returns>True if successful</returns>
        bool AlignBlockArrows();

        /// <summary>
        /// Aligns rounded rectangle radius using master
        /// </summary>
        /// <returns>True if successful</returns>
        bool AlignRoundedRectangleRadius();

        /// <summary>
        /// Gets alignment options for selected shapes
        /// </summary>
        /// <returns>Available alignment options</returns>
        IEnumerable<AlignmentOption> GetAlignmentOptions();

        /// <summary>
        /// Validates if shapes can be aligned with specified method
        /// </summary>
        /// <param name="alignmentType">Type of alignment</param>
        /// <returns>Validation result</returns>
        ShapeValidationResult ValidateShapesForAlignment(AlignmentType alignmentType);
    }
} 