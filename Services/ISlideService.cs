using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn.Models;

namespace PowerPointAddIn.Services
{
    /// <summary>
    /// Interface for slide-related operations
    /// </summary>
    public interface ISlideService
    {
        /// <summary>
        /// Changes slide size with smart scaling
        /// </summary>
        /// <param name="sizeInfo">Slide size information</param>
        /// <returns>True if successful</returns>
        bool ChangeSlideSize(SlideSizeInfo sizeInfo);

        /// <summary>
        /// Gets current slide dimensions
        /// </summary>
        /// <returns>Current slide size info</returns>
        SlideSizeInfo GetCurrentSlideSize();

        /// <summary>
        /// Analyzes presentation content and suggests optimal size
        /// </summary>
        /// <returns>Suggested size with reasoning</returns>
        SlideSizeSuggestion AnalyzeAndSuggestSize();

        /// <summary>
        /// Scales content intelligently when slide size changes
        /// </summary>
        /// <param name="presentation">Presentation to scale</param>
        /// <param name="scaleX">X scale factor</param>
        /// <param name="scaleY">Y scale factor</param>
        void ScaleContentIntelligently(PowerPoint.Presentation presentation, float scaleX, float scaleY);

        /// <summary>
        /// Auto-fits content to slide bounds
        /// </summary>
        /// <returns>Number of shapes adjusted</returns>
        int AutoFitContentToSlide();
    }
} 