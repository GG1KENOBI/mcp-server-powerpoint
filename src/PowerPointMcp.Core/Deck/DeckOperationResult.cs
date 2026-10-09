using Sbroenne.PowerPointMcp.Core.Master;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Result of a deck operation. Success == true implies ErrorMessage is null (Rule 1). Lists are
/// paginated: <see cref="TotalCount"/> and <see cref="HasMore"/> describe what was left out.
/// </summary>
public sealed class DeckOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Slide width in points.</summary>
    public float? SlideWidth { get; init; }

    /// <summary>Slide height in points.</summary>
    public float? SlideHeight { get; init; }

    /// <summary>Aspect ratio label (16:9, 4:3, 16:10, or width:height rounded).</summary>
    public string? AspectRatio { get; init; }

    /// <summary>Total number of slides in the presentation.</summary>
    public int? SlideCount { get; init; }

    /// <summary>Sections in order.</summary>
    public IReadOnlyList<DeckSectionInfo>? Sections { get; init; }

    /// <summary>Name of the first design (theme).</summary>
    public string? ThemeName { get; init; }

    /// <summary>Theme color palette of the first master (#RRGGBB by role).</summary>
    public IReadOnlyDictionary<string, string>? ThemeColors { get; init; }

    /// <summary>Theme heading (major) and body (minor) Latin fonts of the first master.</summary>
    public IReadOnlyDictionary<string, string?>? ThemeFonts { get; init; }

    /// <summary>Slide masters with their layouts and whether any slide uses each layout.</summary>
    public IReadOnlyList<MasterOperationResult.MasterInventoryEntry>? Masters { get; init; }

    /// <summary>Built-in document properties (Title, Subject, Author) that are set.</summary>
    public IReadOnlyDictionary<string, string>? Properties { get; init; }

    /// <summary>Slides on this page.</summary>
    public IReadOnlyList<DeckSlideInfo>? Slides { get; init; }

    /// <summary>Objects on this page.</summary>
    public IReadOnlyList<DeckObjectInfo>? Objects { get; init; }

    /// <summary>Total number of items before pagination.</summary>
    public int? TotalCount { get; init; }

    /// <summary>0-based offset of the first returned item.</summary>
    public int? Offset { get; init; }

    /// <summary>Whether more items exist after this page.</summary>
    public bool? HasMore { get; init; }

    /// <summary>Deck content fingerprint; pass it as expected_revision to detect stale edits.</summary>
    public string? Revision { get; init; }

    /// <summary>PPTMCP_ID values used more than once ("slide:ID" or "object:ID").</summary>
    public IReadOnlyList<string>? DuplicateAppIds { get; init; }

    /// <summary>Planned or applied id assignments, for assign-ids.</summary>
    public IReadOnlyList<AppIdAssignment>? Assignments { get; init; }

    /// <summary>Structural changes, for diff-style results.</summary>
    public IReadOnlyList<DeckChange>? Changes { get; init; }

    /// <summary>Non-fatal notes about what was read or skipped.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
