using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Result of a review operation. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class ReviewOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Deck revision the findings refer to (before repair).</summary>
    public string? Revision { get; init; }

    /// <summary>Deck revision after repair.</summary>
    public string? RevisionAfter { get; init; }

    /// <summary>Findings on this page (after repair: the findings that remain).</summary>
    public IReadOnlyList<DeckFinding>? Findings { get; init; }

    /// <summary>Total findings before pagination.</summary>
    public int? TotalFindings { get; init; }

    /// <summary>0-based offset of the first returned finding.</summary>
    public int? Offset { get; init; }

    /// <summary>Whether more findings exist after this page.</summary>
    public bool? HasMore { get; init; }

    /// <summary>Counts by severity and certainty, e.g. error=2, heuristic=5.</summary>
    public IReadOnlyDictionary<string, int>? Counts { get; init; }

    /// <summary>Rule codes that ran.</summary>
    public IReadOnlyList<string>? RulesRun { get; init; }

    /// <summary>What was not checked and why.</summary>
    public IReadOnlyList<string>? NotChecked { get; init; }

    /// <summary>The rule catalog, for list-rules.</summary>
    public IReadOnlyList<ValidationRule>? Rules { get; init; }

    /// <summary>Planned repair actions.</summary>
    public IReadOnlyList<RepairAction>? Actions { get; init; }

    /// <summary>Selected findings repair does not fix, with reasons.</summary>
    public IReadOnlyList<UnresolvedFinding>? Unresolved { get; init; }

    /// <summary>What happened to each action, for repair.</summary>
    public IReadOnlyList<RepairOutcome>? Outcomes { get; init; }

    /// <summary>Structural changes made by repair.</summary>
    public IReadOnlyList<DeckChange>? Changes { get; init; }

    /// <summary>Non-fatal notes.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
