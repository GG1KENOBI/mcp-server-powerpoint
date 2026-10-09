using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>
/// Semantic slide composition (schema pptmcp.composition/1). A JSON spec describes intent (kind,
/// title, points, cards, KPIs, table, chart, timeline, process, hierarchy, quote, references); the
/// server validates it, lays it out on a deterministic grid from a design profile, creates native
/// editable PowerPoint objects, measures text with PowerPoint, fits by shrinking within limits or
/// splitting onto continuation slides, and returns a semantic map. Text is never shortened.
/// </summary>
[ServiceCategory("compose", "Compose")]
[McpTool("compose", Title = "Slide Composition", Destructive = true, Category = "content",
    Description = "Build complete, editable slides from a JSON composition (kinds: title, section, executive-summary, comparison, cards, kpis, image-text, table, chart, timeline, process, hierarchy, quote, appendix). Use kinds for fields and examples, plan for a dry run, create to build, get-spec to read a slide's composition back, update-text to change one part's text keeping its formatting.")]
[McpReadOnlyActions("kinds", "plan", "get-spec")]
public interface IComposeCommands
{
    /// <summary>Lists composition kinds with required and optional fields and a complete example for each.</summary>
    ComposeOperationResult Kinds(IPresentationBatch batch);

    /// <summary>
    /// Validates a composition and returns the planned slides and element boxes without changing
    /// the presentation. Text heights are estimates; create measures with PowerPoint.
    /// </summary>
    /// <param name="spec">Composition JSON (pptmcp.composition/1). Alternatively pass spec_path.</param>
    /// <param name="specPath">Full path of a local .json file holding the composition.</param>
    /// <param name="profile">Design profile: theme (default: the deck's own theme), a built-in (default, corporate-blue, high-contrast, minimal-mono), or a saved profile name.</param>
    ComposeOperationResult Plan(IPresentationBatch batch, string? spec = null, string? specPath = null, string? profile = null);

    /// <summary>
    /// Creates the slide(s). Inserts at insert_at (1-based; default after the last slide). With
    /// replace=true, slides previously created with the same id are replaced in place. Returns
    /// slide ids, persistent ids, the semantic map (key, role, shape id, measured box, fit outcome),
    /// and any overflow that could not be fitted.
    /// </summary>
    /// <param name="insertAt">1-based position for the first slide (overrides the spec's insert_at).</param>
    /// <param name="replace">Replace slides that carry the spec's id (default false: an existing id is an error).</param>
    ComposeOperationResult Create(IPresentationBatch batch, string? spec = null, string? specPath = null, string? profile = null, int? insertAt = null, bool replace = false);

    /// <summary>Returns the composition JSON stored on a composed slide (for editing and re-creating with replace=true).</summary>
    /// <param name="slideAppId">Persistent id of the composed slide (the spec id).</param>
    ComposeOperationResult GetSpec(IPresentationBatch batch, string slideAppId);

    /// <summary>
    /// Replaces the text of one composed part (e.g. app id "summary/point-2" or "kpis/kpi-1.value")
    /// keeping the first run's formatting, then measures whether it still fits.
    /// </summary>
    /// <param name="appId">PPTMCP_ID of the part (slide id, slash, element key).</param>
    /// <param name="text">New text; newlines become line breaks inside the paragraph.</param>
    ComposeOperationResult UpdateText(IPresentationBatch batch, string appId, string text);
}
