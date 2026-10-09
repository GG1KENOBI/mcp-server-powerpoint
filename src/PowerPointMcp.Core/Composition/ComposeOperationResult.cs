namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>A slide created by a composition, with its semantic map.</summary>
public sealed class ComposedSlide
{
    /// <summary>PowerPoint SlideID (stable).</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based position.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>Persistent slide id (PPTMCP_ID).</summary>
    public required string AppId { get; init; }

    /// <summary>Semantic map: every created object by key, with ids, geometry, and fit outcome.</summary>
    public required IReadOnlyList<RenderedElement> Elements { get; init; }
}

/// <summary>A planned slide in a dry run.</summary>
public sealed class PlannedSlideSummary
{
    /// <summary>Persistent slide id it would get.</summary>
    public required string AppId { get; init; }

    /// <summary>Title.</summary>
    public string? Title { get; init; }

    /// <summary>Items (points, rows, references) on this slide.</summary>
    public int? Items { get; init; }

    /// <summary>Planned elements: key, role, type, box (left, top, width, height), estimated text height, smallest font.</summary>
    public required IReadOnlyList<PlannedElementSummary> Elements { get; init; }
}

/// <summary>One planned element in a dry run.</summary>
public sealed record PlannedElementSummary(string Key, string Role, string Type, IReadOnlyList<float> Box, float? EstimatedHeight, float? MinFontSize, string? Text);

/// <summary>Result of a compose operation. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class ComposeOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Validation errors with JSON paths.</summary>
    public IReadOnlyList<string>? Errors { get; init; }

    /// <summary>Non-fatal notes: ignored fields, planning estimates, styling that could not be applied.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }

    /// <summary>Design profile used (name@version).</summary>
    public string? Profile { get; init; }

    /// <summary>Slides created.</summary>
    public IReadOnlyList<ComposedSlide>? Slides { get; init; }

    /// <summary>Planned slides, for plan (dry run).</summary>
    public IReadOnlyList<PlannedSlideSummary>? Planned { get; init; }

    /// <summary>Measured overflow that fitting could not resolve; text is never cut.</summary>
    public IReadOnlyList<string>? Overflow { get; init; }

    /// <summary>How many measure-and-split passes rendering needed.</summary>
    public int? RenderPasses { get; init; }

    /// <summary>Composition kinds, for kinds.</summary>
    public IReadOnlyList<CompositionKind>? Kinds { get; init; }

    /// <summary>Schema identifier.</summary>
    public string? Schema { get; init; }

    /// <summary>Stored composition JSON, for get-spec.</summary>
    public string? Spec { get; init; }

    /// <summary>Text height measured after update-text (including margins).</summary>
    public float? MeasuredHeight { get; init; }

    /// <summary>Box height after update-text.</summary>
    public float? BoxHeight { get; init; }
}
