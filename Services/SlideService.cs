using System;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PowerPointAddIn.Models;
using PowerPointAddIn.Constants;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Slide operations service implementation
    /// </summary>
    public class SlideService : ISlideService
    {
        private readonly IPowerPointService _powerPointService;
        private readonly IErrorHandlerService _errorHandler;

        public SlideService(IPowerPointService powerPointService, IErrorHandlerService errorHandler)
        {
            _powerPointService = powerPointService ?? throw new ArgumentNullException(nameof(powerPointService));
            _errorHandler = errorHandler ?? throw new ArgumentNullException(nameof(errorHandler));
        }

        /// <summary>
        /// Changes slide size with smart scaling
        /// </summary>
        public bool ChangeSlideSize(SlideSizeInfo sizeInfo)
        {
            if (sizeInfo == null)
                throw new ArgumentNullException(nameof(sizeInfo));

            return ExecuteWithErrorHandling(() =>
            {
                var presentation = _powerPointService.ActivePresentation;
                if (presentation == null)
                {
                    _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                    return false;
                }

                // Store current dimensions for scaling calculations
                float oldWidth = presentation.PageSetup.SlideWidth;
                float oldHeight = presentation.PageSetup.SlideHeight;

                // Apply new size
                presentation.PageSetup.SlideWidth = (float)sizeInfo.Width * AppConstants.Measurements.PointsPerInch;
                presentation.PageSetup.SlideHeight = (float)sizeInfo.Height * AppConstants.Measurements.PointsPerInch;

                // Scale content if requested
                if (sizeInfo.ScaleContent)
                {
                    float scaleX = presentation.PageSetup.SlideWidth / oldWidth;
                    float scaleY = presentation.PageSetup.SlideHeight / oldHeight;
                    ScaleContentIntelligently(presentation, scaleX, scaleY);
                }

                _errorHandler.ShowInfo($"Slide size changed to {sizeInfo}!\nContent has been intelligently scaled.", 
                    "Smart Size Applied");

                return true;
            }, "changing slide size", false);
        }

        /// <summary>
        /// Gets current slide dimensions
        /// </summary>
        public SlideSizeInfo GetCurrentSlideSize()
        {
            return ExecuteWithErrorHandling(() =>
            {
                var presentation = _powerPointService.ActivePresentation;
                if (presentation == null)
                    return null;

                float widthInches = presentation.PageSetup.SlideWidth / AppConstants.Measurements.PointsPerInch;
                float heightInches = presentation.PageSetup.SlideHeight / AppConstants.Measurements.PointsPerInch;

                // Find closest preset
                var closestPreset = FindClosestPreset((decimal)widthInches, (decimal)heightInches);
                return closestPreset ?? new SlideSizeInfo((decimal)widthInches, (decimal)heightInches, "Custom", SlideSizeType.Custom);
            }, "getting current slide size", (SlideSizeInfo)null);
        }

        /// <summary>
        /// Analyzes presentation content and suggests optimal size
        /// </summary>
        public SlideSizeSuggestion AnalyzeAndSuggestSize()
        {
            return ExecuteWithErrorHandling(() =>
            {
                var presentation = _powerPointService.ActivePresentation;
                if (presentation == null)
                {
                    _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                    return null;
                }

                var analysisResult = AnalyzeContent(presentation);
                var recommendation = GenerateSizeRecommendation(analysisResult);

                return new SlideSizeSuggestion(recommendation.size, recommendation.reasoning, recommendation.confidence)
                {
                    AnalysisResult = analysisResult
                };
            }, "analyzing content for size suggestion", (SlideSizeSuggestion)null);
        }

        /// <summary>
        /// Scales content intelligently when slide size changes
        /// </summary>
        public void ScaleContentIntelligently(PowerPoint.Presentation presentation, float scaleX, float scaleY)
        {
            ExecuteWithErrorHandling(() =>
            {
                // Use the smaller scale factor to maintain aspect ratios
                float scaleFactor = Math.Min(scaleX, scaleY);

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // Scale position and size
                        shape.Left *= scaleFactor;
                        shape.Top *= scaleFactor;
                        shape.Width *= scaleFactor;
                        shape.Height *= scaleFactor;

                        // Scale font size proportionally
                        ScaleShapeText(shape, scaleFactor);
                    }
                }
            }, "scaling content intelligently");
        }

        /// <summary>
        /// Auto-fits content to slide bounds
        /// </summary>
        public int AutoFitContentToSlide()
        {
            return ExecuteWithErrorHandling(() =>
            {
                var presentation = _powerPointService.ActivePresentation;
                if (presentation == null)
                {
                    _errorHandler.ShowWarning(AppConstants.Messages.NoActivePresentation);
                    return 0;
                }

                int adjustedShapes = 0;
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (IsShapeOutOfBounds(shape, presentation))
                        {
                            FitShapeToSlide(shape, presentation);
                            adjustedShapes++;
                        }
                    }
                }

                if (adjustedShapes > 0)
                {
                    _errorHandler.ShowInfo($"Auto-fitted {adjustedShapes} shapes to slide bounds!", 
                        "Content Auto-Fit");
                }

                return adjustedShapes;
            }, "auto-fitting content to slide", 0);
        }

        #region Private Helper Methods

        private SlideSizeInfo FindClosestPreset(decimal width, decimal height)
        {
            SlideSizeInfo closestPreset = null;
            double minDifference = double.MaxValue;

            foreach (var preset in AppConstants.SlideSizePresets.Values)
            {
                double diff = Math.Abs((double)(preset.Width - width)) + Math.Abs((double)(preset.Height - height));
                if (diff < minDifference)
                {
                    minDifference = diff;
                    closestPreset = preset;
                }
            }

            return minDifference < AppConstants.Measurements.ClosestPresetTolerance ? closestPreset : null;
        }

        private ContentAnalysisResult AnalyzeContent(PowerPoint.Presentation presentation)
        {
            var result = new ContentAnalysisResult { TotalSlides = presentation.Slides.Count };

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                bool hasImages = false, hasCharts = false, hasTables = false, 
                     hasText = false, hasVideos = false, hasSmartArt = false;

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    switch (shape.Type)
                    {
                        case Office.MsoShapeType.msoPicture:
                        case Office.MsoShapeType.msoLinkedPicture:
                            hasImages = true;
                            break;
                        case Office.MsoShapeType.msoMedia:
                            hasVideos = true;
                            break;
                        case Office.MsoShapeType.msoSmartArt:
                            hasSmartArt = true;
                            break;
                        default:
                            if (shape.HasChart == Office.MsoTriState.msoTrue)
                                hasCharts = true;
                            else if (shape.HasTable == Office.MsoTriState.msoTrue)
                                hasTables = true;
                            else if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                                hasText = true;
                            break;
                    }
                }

                if (hasImages) result.SlidesWithImages++;
                if (hasCharts) result.SlidesWithCharts++;
                if (hasTables) result.SlidesWithTables++;
                if (hasText) result.SlidesWithText++;
                if (hasVideos) result.SlidesWithVideos++;
                if (hasSmartArt) result.SlidesWithSmartArt++;
            }

            return result;
        }

        private (SlideSizeInfo size, string reasoning, double confidence) GenerateSizeRecommendation(ContentAnalysisResult analysis)
        {
            // Priority-based recommendations
            if (analysis.VideoRatio > 0.3)
                return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                       "Optimal for video content and modern displays", 0.9);

            if (analysis.ImageRatio > 0.7)
                return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                       "Best for image-heavy presentations and visual storytelling", 0.85);

            if (analysis.ChartRatio > 0.5)
                return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                       "Perfect for data visualization and charts", 0.8);

            if (analysis.SmartArtRatio > 0.4)
                return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                       "Ideal for SmartArt and modern diagrams", 0.8);

            if (analysis.TableRatio > 0.6)
                return (AppConstants.SlideSizePresets[SlideSizeType.A4Landscape], 
                       "Excellent for table-heavy content and detailed data", 0.75);

            if (analysis.TextRatio > 0.8)
                return (AppConstants.SlideSizePresets[SlideSizeType.Standard], 
                       "Traditional format for text-heavy academic presentations", 0.7);

            if (analysis.TotalSlides > 20)
                return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                       "Recommended for large presentations", 0.6);

            if (analysis.TotalSlides < 5)
                return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                       "Perfect for short, impactful presentations", 0.6);

            return (AppConstants.SlideSizePresets[SlideSizeType.Widescreen16x9], 
                   "Modern standard for most presentations", 0.5);
        }

        private void ScaleShapeText(PowerPoint.Shape shape, float scaleFactor)
        {
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;
                    if (textRange.Font.Size > 0)
                    {
                        textRange.Font.Size *= scaleFactor;
                    }
                }
            }
            catch (Exception ex)
            {
                _errorHandler.LogError(ex, $"Scaling text for shape {shape.Name}");
            }
        }

        private bool IsShapeOutOfBounds(PowerPoint.Shape shape, PowerPoint.Presentation presentation)
        {
            return shape.Left < 0 || shape.Top < 0 ||
                   shape.Left + shape.Width > presentation.PageSetup.SlideWidth ||
                   shape.Top + shape.Height > presentation.PageSetup.SlideHeight;
        }

        private void FitShapeToSlide(PowerPoint.Shape shape, PowerPoint.Presentation presentation)
        {
            if (shape.Left < 0) shape.Left = 0;
            if (shape.Top < 0) shape.Top = 0;

            if (shape.Left + shape.Width > presentation.PageSetup.SlideWidth)
                shape.Left = presentation.PageSetup.SlideWidth - shape.Width;
            if (shape.Top + shape.Height > presentation.PageSetup.SlideHeight)
                shape.Top = presentation.PageSetup.SlideHeight - shape.Height;
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

        #endregion
    }
} 