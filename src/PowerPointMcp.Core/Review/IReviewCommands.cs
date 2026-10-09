using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>
/// Deck validation and repair. Findings separate deterministic facts (measured text bounds,
/// geometry, font sizes, files) from heuristic judgments (accidental overlap, margins, contrast).
/// Repair only moves, resizes, rescales fonts proportionally, or renames ids; it never deletes,
/// shortens, flattens, or rasterizes content, and it does not save the file.
/// </summary>
[ServiceCategory("review", "Review")]
[McpTool("review", Title = "Deck Validation and Repair", Destructive = true, Category = "content",
    Description = "Validate slides for layout, typography, image, asset, font, footer, and accessibility problems with evidence and suggested fixes; plan repairs as a dry run; apply repairs that move, resize, or proportionally scale text without deleting or rewriting content. Loop: validate, plan-repair, repair, preview.")]
[McpReadOnlyActions("validate", "plan-repair", "list-rules")]
public interface IReviewCommands
{
    /// <summary>
    /// Validates the deck (or one slide) and returns findings, most severe first, each with code,
    /// severity, certainty (deterministic or heuristic), detection method, ids, evidence, and a
    /// suggested correction. Paginated; counts cover all findings.
    /// </summary>
    /// <param name="slideIndex">Limit to this 1-based slide (deck-wide rules then see only that slide).</param>
    /// <param name="rules">Rule codes to run (default all). See list-rules.</param>
    /// <param name="minFontSize">Smallest acceptable body text in points (default 10).</param>
    /// <param name="safeMargin">Points to keep free at slide edges; enables outside-safe-area.</param>
    /// <param name="includeHeuristic">Include heuristic rules (default true).</param>
    /// <param name="minSeverity">Lowest severity returned: info (default), warning, or error.</param>
    /// <param name="offset">0-based index of the first finding returned (default 0).</param>
    /// <param name="limit">Findings per page (1-200, default 50).</param>
    ReviewOperationResult Validate(
        IPresentationBatch batch,
        int? slideIndex = null,
        IReadOnlyList<string>? rules = null,
        float minFontSize = 10f,
        float? safeMargin = null,
        bool includeHeuristic = true,
        string minSeverity = "info",
        int offset = 0,
        int limit = 50);

    /// <summary>
    /// Validates, then plans repairs without changing anything: each action names the object,
    /// geometry before and after, the basis (geometry, measured, or estimated), and the finding it
    /// fixes; findings that will not be repaired are returned with the reason.
    /// </summary>
    /// <param name="codes">Rule codes to repair (default: off-slide, text-overflow, text-overlap, image-distortion, inconsistent-title-position, near-misaligned, duplicate-app-id).</param>
    /// <param name="findingIds">Repair only these finding ids from validate.</param>
    /// <param name="fitPolicy">Text-overflow strategies in order: expand, shrink (default both).</param>
    ReviewOperationResult PlanRepair(
        IPresentationBatch batch,
        int? slideIndex = null,
        IReadOnlyList<string>? codes = null,
        IReadOnlyList<string>? findingIds = null,
        IReadOnlyList<string>? fitPolicy = null,
        float minFontSize = 10f,
        float? safeMargin = null);

    /// <summary>
    /// Applies the repair plan, then validates again. Font scaling is measured: every run is
    /// scaled by the same factor, stepping down until PowerPoint reports the text fits, never
    /// below min_font_size. Actions on objects that moved since planning are skipped. Returns
    /// outcomes, remaining findings, and structural changes. Does not save.
    /// </summary>
    /// <param name="expectedRevision">Refuse to repair unless the deck revision (from deck fingerprint or validate) still matches.</param>
    ReviewOperationResult Repair(
        IPresentationBatch batch,
        int? slideIndex = null,
        IReadOnlyList<string>? codes = null,
        IReadOnlyList<string>? findingIds = null,
        IReadOnlyList<string>? fitPolicy = null,
        float minFontSize = 10f,
        float? safeMargin = null,
        string? expectedRevision = null);

    /// <summary>Lists every validation rule with certainty, detection method, and whether repair fixes it.</summary>
    ReviewOperationResult ListRules(IPresentationBatch batch);
}
