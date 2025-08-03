using System.Collections.Generic;

namespace PowerPointAddIn.Models
{
    /// <summary>
    /// Represents an alignment option for shapes
    /// </summary>
    public class AlignmentOption
    {
        public AlignmentType Type { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public bool RequiresMaster { get; set; }
        public int MinimumShapes { get; set; }

        public AlignmentOption(AlignmentType type, string name, string description, bool requiresMaster = false, int minimumShapes = 2)
        {
            Type = type;
            Name = name;
            Description = description;
            RequiresMaster = requiresMaster;
            MinimumShapes = minimumShapes;
        }
    }

    /// <summary>
    /// Result of shape validation for alignment operations
    /// </summary>
    public class ShapeValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; }
        public int ShapeCount { get; set; }
        public AlignmentType RequestedAlignment { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();

        public static ShapeValidationResult Success(int shapeCount, AlignmentType alignmentType)
        {
            return new ShapeValidationResult
            {
                IsValid = true,
                ShapeCount = shapeCount,
                RequestedAlignment = alignmentType
            };
        }

        public static ShapeValidationResult Failure(string errorMessage, AlignmentType alignmentType)
        {
            return new ShapeValidationResult
            {
                IsValid = false,
                ErrorMessage = errorMessage,
                RequestedAlignment = alignmentType
            };
        }
    }

    /// <summary>
    /// Types of shape alignment operations
    /// </summary>
    public enum AlignmentType
    {
        ProcessChain,
        Angles,
        ToProcessArrow,
        PentagonHeader,
        BlockArrows,
        RoundedRectangleRadius,
        Left,
        Center,
        Right,
        Top,
        Bottom,
        Middle,
        DistributeHorizontal,
        DistributeVertical,
        MatchWidth,
        MatchHeight,
        MatchBoth
    }

    /// <summary>
    /// Position-related operations
    /// </summary>
    public enum PositionOperation
    {
        AlignLeft,
        AlignCenter,
        AlignRight,
        AlignTop,
        AlignBottom,
        AlignMiddle,
        DockLeft,
        DockRight,
        DockTop,
        DockBottom,
        Distribute,
        DistributeHorizontal,
        DistributeVertical,
        MatchBoth,
        MatchHeight,
        MatchWidth,
        MakeVertical,
        MakeHorizontal,
        SwapLocations,
        GoldenCanon,
        AlignMatrix,
        SliceShape,
        DuplicateRight,
        CenterTopLeft,
        SavePosition,
        ApplyPosition
    }
} 