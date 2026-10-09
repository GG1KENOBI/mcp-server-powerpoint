namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>
/// One validation finding. Deterministic findings are facts measured by PowerPoint or read from
/// the file (geometry, measured text bounds, font sizes, file existence). Heuristic findings are
/// judgments about intent (accidental overlap, inconsistent margins, contrast) and may be wrong.
/// Geometry is not visual understanding: render the slide to confirm what a heuristic reports.
/// </summary>
public sealed class DeckFinding
{
    /// <summary>Stable identifier of this finding (code, slide id, and shape ids), usable to select repairs.</summary>
    public required string Id { get; init; }

    /// <summary>Rule code, e.g. text-overflow or off-slide.</summary>
    public required string Code { get; init; }

    /// <summary>error, warning, or info.</summary>
    public required string Severity { get; init; }

    /// <summary>deterministic or heuristic.</summary>
    public required string Certainty { get; init; }

    /// <summary>How it was detected, e.g. "measured: TextRange bounds vs shape height".</summary>
    public required string Method { get; init; }

    /// <summary>1-based slide position (0 for deck-wide findings).</summary>
    public required int SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID (0 for deck-wide findings).</summary>
    public required int SlideId { get; init; }

    /// <summary>Shape.Id values of the objects involved.</summary>
    public IReadOnlyList<int>? ShapeIds { get; init; }

    /// <summary>Top-level 1-based shape indexes of the objects involved, where they are top level.</summary>
    public IReadOnlyList<int>? ShapeIndexes { get; init; }

    /// <summary>Persistent PPTMCP_ID values of the objects involved, where set.</summary>
    public IReadOnlyList<string>? AppIds { get; init; }

    /// <summary>Slide positions involved, for deck-wide findings.</summary>
    public IReadOnlyList<int>? SlideIndexes { get; init; }

    /// <summary>Human-readable description.</summary>
    public required string Message { get; init; }

    /// <summary>Measured values that support the finding (points unless stated).</summary>
    public IReadOnlyDictionary<string, string>? Evidence { get; init; }

    /// <summary>Suggested correction in terms of existing tools and parameters.</summary>
    public required string Suggestion { get; init; }

    /// <summary>Whether review repair can fix it automatically.</summary>
    public required bool AutoRepairable { get; init; }
}
