namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>One deterministic layout problem found by slide check-layout.</summary>
public sealed class SlideLayoutIssue
{
    /// <summary>
    /// Issue code: text-overflow, off-slide, text-overlap, partial-overlap, small-text,
    /// empty-placeholder, or near-misaligned.
    /// </summary>
    public required string Code { get; init; }

    /// <summary>error (visible defect), warning (likely defect), or info (polish).</summary>
    public required string Severity { get; init; }

    /// <summary>1-based index of the slide the issue is on.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>1-based indexes of the shapes involved.</summary>
    public required IReadOnlyList<int> ShapeIndexes { get; init; }

    /// <summary>What is wrong, with the measured values.</summary>
    public required string Message { get; init; }

    /// <summary>A concrete fix, in points where it applies.</summary>
    public required string Suggestion { get; init; }
}
