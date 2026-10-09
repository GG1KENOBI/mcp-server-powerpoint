namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Thresholds and rule selection for deck validation. Lengths are points.</summary>
public sealed class ValidationOptions
{
    /// <summary>Body text below this size is small-text.</summary>
    public float MinFontSize { get; init; } = 10f;

    /// <summary>Footers, slide numbers, dates, and source lines below this size are small-text.</summary>
    public float MinFootnoteFontSize { get; init; } = 8f;

    /// <summary>When set, content closer than this to a slide edge is outside-safe-area.</summary>
    public float? SafeMargin { get; init; }

    /// <summary>Edges closer than this (but not equal) are near-misaligned.</summary>
    public float AlignmentBand { get; init; } = 4f;

    /// <summary>Minimum WCAG contrast ratio for normal text (large text uses 3.0).</summary>
    public double MinContrast { get; init; } = 4.5;

    /// <summary>Include heuristic findings (default true).</summary>
    public bool IncludeHeuristic { get; init; } = true;

    /// <summary>Only run these rule codes (null runs all).</summary>
    public IReadOnlySet<string>? Rules { get; init; }

    /// <summary>Installed font family names; null skips the missing-font rule.</summary>
    public IReadOnlySet<string>? InstalledFonts { get; init; }
}
