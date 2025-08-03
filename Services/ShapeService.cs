using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PowerPointAddIn.Models;
using PowerPointAddIn.Constants;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Shape operations service implementation
    /// </summary>
    public class ShapeService : IShapeService
    {
        private readonly IPowerPointService _powerPointService;
        private readonly IErrorHandlerService _errorHandler;

        public ShapeService(IPowerPointService powerPointService, IErrorHandlerService errorHandler)
        {
            _powerPointService = powerPointService ?? throw new ArgumentNullException(nameof(powerPointService));
            _errorHandler = errorHandler ?? throw new ArgumentNullException(nameof(errorHandler));
        }

        /// <summary>
        /// Aligns selected shapes in a process chain
        /// </summary>
        public bool AlignProcessChain()
        {
            if (!_powerPointService.ValidateShapeSelection(2, "process chain alignment"))
                return false;

            return ExecuteWithErrorHandling(() =>
            {
                var shapes = _powerPointService.SelectedShapes;
                var sortedShapes = shapes.Cast<PowerPoint.Shape>().OrderBy(s => s.Left).ToArray();

                // Calculate equal spacing
                float totalWidth = sortedShapes.Sum(s => s.Width);
                float availableWidth = _powerPointService.ActivePresentation.PageSetup.SlideWidth;
                float spacing = (availableWidth - totalWidth) / (sortedShapes.Length + 1);
                float currentX = spacing;

                foreach (var shape in sortedShapes)
                {
                    shape.Left = currentX;
                    currentX += shape.Width + spacing;
                }

                _errorHandler.ShowInfo("Process chain aligned successfully!", AppConstants.Messages.SuccessTitle);
                return true;
            }, "aligning process chain", false);
        }

        /// <summary>
        /// Aligns angles of selected shapes to master
        /// </summary>
        public bool AlignAngles()
        {
            if (!_powerPointService.ValidateShapeSelection(2, "angle alignment"))
                return false;

            return ExecuteWithErrorHandling(() =>
            {
                var shapes = _powerPointService.SelectedShapes;
                var masterShape = shapes[shapes.Count]; // Last selected is master
                float targetAngle = masterShape.Rotation;

                foreach (PowerPoint.Shape shape in shapes)
                {
                    shape.Rotation = targetAngle;
                }

                _errorHandler.ShowInfo("Shapes aligned to master angle!", AppConstants.Messages.SuccessTitle);
                return true;
            }, "aligning angles", false);
        }

        /// <summary>
        /// Aligns selected objects to process arrow
        /// </summary>
        public bool AlignToProcessArrow()
        {
            if (!_powerPointService.ValidateShapeSelection(2, "process arrow alignment"))
                return false;

            return ExecuteWithErrorHandling(() =>
            {
                var shapes = _powerPointService.SelectedShapes;
                PowerPoint.Shape arrowShape = null;

                // Find arrow shape
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (IsArrowShape(shape))
                    {
                        arrowShape = shape;
                        break;
                    }
                }

                if (arrowShape == null)
                {
                    _errorHandler.ShowWarning("Please include an arrow shape in your selection.");
                    return false;
                }

                // Align other shapes to arrow's center
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (shape != arrowShape)
                    {
                        shape.Top = arrowShape.Top + (arrowShape.Height - shape.Height) / 2;
                    }
                }

                _errorHandler.ShowInfo("Shapes aligned to process arrow!", AppConstants.Messages.SuccessTitle);
                return true;
            }, "aligning to process arrow", false);
        }

        /// <summary>
        /// Adjusts pentagon header boxes
        /// </summary>
        public bool AdjustPentagonHeader()
        {
            if (!_powerPointService.ValidateShapeSelection(2, "pentagon header adjustment"))
                return false;

            return ExecuteWithErrorHandling(() =>
            {
                var shapes = _powerPointService.SelectedShapes;
                PowerPoint.Shape pentagon = null;
                PowerPoint.Shape header = null;

                // Find pentagon and header shapes
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (IsPentagonShape(shape))
                        pentagon = shape;
                    else if (IsHeaderShape(shape))
                        header = shape;
                }

                if (pentagon == null || header == null)
                {
                    _errorHandler.ShowWarning("Please select a pentagon and header box.");
                    return false;
                }

                // Adjust header to pentagon
                header.Left = pentagon.Left;
                header.Width = pentagon.Width * 0.8f; // 80% of pentagon width
                header.Top = pentagon.Top - header.Height - 5; // 5pt gap

                _errorHandler.ShowInfo("Pentagon header adjusted!", AppConstants.Messages.SuccessTitle);
                return true;
            }, "adjusting pentagon header", false);
        }

        /// <summary>
        /// Aligns block arrows with standard style
        /// </summary>
        public bool AlignBlockArrows()
        {
            if (!_powerPointService.ValidateShapeSelection(2, "block arrow alignment"))
                return false;

            return ExecuteWithErrorHandling(() =>
            {
                var shapes = _powerPointService.SelectedShapes;
                var blockArrows = shapes.Cast<PowerPoint.Shape>()
                    .Where(IsBlockArrowShape)
                    .ToList();

                if (blockArrows.Count < 2)
                {
                    _errorHandler.ShowWarning("Please select multiple block arrow shapes.");
                    return false;
                }

                var masterArrow = blockArrows.Last(); // Last selected is master
                
                foreach (var arrow in blockArrows)
                {
                    if (arrow != masterArrow)
                    {
                        // Apply master's style
                        arrow.Width = masterArrow.Width;
                        arrow.Height = masterArrow.Height;
                        arrow.Rotation = masterArrow.Rotation;
                    }
                }

                _errorHandler.ShowInfo("Block arrows aligned to master style!", AppConstants.Messages.SuccessTitle);
                return true;
            }, "aligning block arrows", false);
        }

        /// <summary>
        /// Aligns rounded rectangle radius using master
        /// </summary>
        public bool AlignRoundedRectangleRadius()
        {
            if (!_powerPointService.ValidateShapeSelection(2, "rounded rectangle radius alignment"))
                return false;

            return ExecuteWithErrorHandling(() =>
            {
                var shapes = _powerPointService.SelectedShapes;
                var roundedRects = shapes.Cast<PowerPoint.Shape>()
                    .Where(IsRoundedRectangleShape)
                    .ToList();

                if (roundedRects.Count < 2)
                {
                    _errorHandler.ShowWarning("Please select multiple rounded rectangle shapes.");
                    return false;
                }

                var masterRect = roundedRects.Last(); // Last selected is master

                foreach (var rect in roundedRects)
                {
                    if (rect != masterRect)
                    {
                        try
                        {
                            rect.Adjustments[1] = masterRect.Adjustments[1];
                        }
                        catch
                        {
                            // Fallback: copy other properties if adjustments fail
                            rect.Fill.ForeColor = masterRect.Fill.ForeColor;
                            rect.Line.Weight = masterRect.Line.Weight;
                        }
                    }
                }

                _errorHandler.ShowInfo("Rounded rectangle radius aligned!", AppConstants.Messages.SuccessTitle);
                return true;
            }, "aligning rounded rectangle radius", false);
        }

        /// <summary>
        /// Gets alignment options for selected shapes
        /// </summary>
        public IEnumerable<AlignmentOption> GetAlignmentOptions()
        {
            var options = new List<AlignmentOption>
            {
                new AlignmentOption(AlignmentType.ProcessChain, "Process Chain", 
                    "Align selected block arrows to form a process chain", false, 2),
                new AlignmentOption(AlignmentType.Angles, "Align Angles", 
                    "Align angles of all selected shapes to the master", true, 2),
                new AlignmentOption(AlignmentType.ToProcessArrow, "To Process Arrow", 
                    "Align selected objects to the process arrow", false, 2),
                new AlignmentOption(AlignmentType.PentagonHeader, "Pentagon Header", 
                    "Adjust header boxes of pentagon shapes", false, 2),
                new AlignmentOption(AlignmentType.BlockArrows, "Block Arrows", 
                    "Apply standard style to all selected block arrows", true, 2),
                new AlignmentOption(AlignmentType.RoundedRectangleRadius, "Rounded Rectangle Radius", 
                    "Define radius for selected rounded rectangles", true, 2)
            };

            return options;
        }

        /// <summary>
        /// Validates if shapes can be aligned with specified method
        /// </summary>
        public ShapeValidationResult ValidateShapesForAlignment(AlignmentType alignmentType)
        {
            if (!_powerPointService.HasSelectedShapes())
            {
                return ShapeValidationResult.Failure(AppConstants.Messages.NoSelectedShapes, alignmentType);
            }

            var shapes = _powerPointService.SelectedShapes;
            var shapeCount = shapes?.Count ?? 0;
            var alignmentOption = GetAlignmentOptions().FirstOrDefault(o => o.Type == alignmentType);

            if (alignmentOption == null)
            {
                return ShapeValidationResult.Failure("Unknown alignment type", alignmentType);
            }

            if (shapeCount < alignmentOption.MinimumShapes)
            {
                var message = string.Format(AppConstants.Messages.InsufficientShapes, alignmentOption.MinimumShapes);
                return ShapeValidationResult.Failure(message, alignmentType);
            }

            var result = ShapeValidationResult.Success(shapeCount, alignmentType);

            // Add specific warnings based on alignment type
            switch (alignmentType)
            {
                case AlignmentType.ToProcessArrow:
                    if (!shapes.Cast<PowerPoint.Shape>().Any(IsArrowShape))
                        result.Warnings.Add("No arrow shape detected in selection");
                    break;
                case AlignmentType.PentagonHeader:
                    if (!shapes.Cast<PowerPoint.Shape>().Any(IsPentagonShape))
                        result.Warnings.Add("No pentagon shape detected in selection");
                    break;
                case AlignmentType.BlockArrows:
                    if (!shapes.Cast<PowerPoint.Shape>().Any(IsBlockArrowShape))
                        result.Warnings.Add("No block arrow shapes detected in selection");
                    break;
                case AlignmentType.RoundedRectangleRadius:
                    if (!shapes.Cast<PowerPoint.Shape>().Any(IsRoundedRectangleShape))
                        result.Warnings.Add("No rounded rectangle shapes detected in selection");
                    break;
            }

            return result;
        }

        #region Private Helper Methods

        private bool IsArrowShape(PowerPoint.Shape shape)
        {
            return shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRightArrow ||
                   shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeLeftArrow ||
                   shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeUpArrow ||
                   shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeDownArrow ||
                   shape.Name.ToLower().Contains("arrow");
        }

        private bool IsPentagonShape(PowerPoint.Shape shape)
        {
            return shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRegularPentagon ||
                   shape.Name.ToLower().Contains("pentagon");
        }

        private bool IsHeaderShape(PowerPoint.Shape shape)
        {
            return shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle ||
                   shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRoundedRectangle ||
                   (shape.HasTextFrame == Office.MsoTriState.msoTrue && 
                    shape.Name.ToLower().Contains("header"));
        }

        private bool IsBlockArrowShape(PowerPoint.Shape shape)
        {
            return shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeBlockArc ||
                   shape.Name.ToLower().Contains("block") ||
                   IsArrowShape(shape);
        }

        private bool IsRoundedRectangleShape(PowerPoint.Shape shape)
        {
            return shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRoundedRectangle ||
                   shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRoundedRectangularCallout;
        }

        /// <summary>
        /// Helper method for executing functions with error handling
        /// </summary>
        private T ExecuteWithErrorHandling<T>(Func<T> function, string operationName, T defaultValue = default(T))
        {
            return _errorHandler.ExecuteWithErrorHandling(function, operationName, null, defaultValue);
        }

        #endregion
    }
} 