using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Deck-level inspection, semantic addressing, and identity commands. Geometry is in points.
/// Objects are addressed by PowerPoint SlideID + Shape.Id (stable across reordering), by the
/// persistent PPTMCP_ID tag, or by a selector such as <c>kind:table slide:3</c>.
/// </summary>
[ServiceCategory("deck", "Deck")]
[McpTool("deck", Title = "Deck Inspection and Addressing", Destructive = true, Category = "content",
    Description = "Inspect a whole presentation compactly, list objects with stable ids, find objects by semantic selectors (title, role, kind, text, tag, group, geometry), fingerprint revisions, and assign persistent ids. Selector example: kind:table slide:3, text:\"Q3*\" font<12.")]
[McpReadOnlyActions("summary", "inspect-objects", "find", "fingerprint")]
public interface IDeckCommands
{
    /// <summary>
    /// Returns slide size and aspect ratio, sections, theme name, colors and fonts, masters with
    /// their layouts, document properties, and one page of slides (id, title, layout, hidden,
    /// section, notes preview, shape count, fingerprint). Paginated by start slide and count; the
    /// deck revision is included when the page covers every slide (otherwise use fingerprint).
    /// </summary>
    /// <param name="startSlide">1-based first slide of the page (default 1).</param>
    /// <param name="maxSlides">Slides per page (1-100, default 25).</param>
    /// <param name="includeTheme">Include theme colors, fonts, masters, and layouts (default true).</param>
    DeckOperationResult Summary(IPresentationBatch batch, int startSlide = 1, int maxSlides = 25, bool includeTheme = true);

    /// <summary>
    /// Lists objects with stable ids, kind, role, geometry, z-order, visibility, text, font sizes,
    /// autofit, measured text bounds, group membership, and table/chart/link/connector metadata.
    /// Target one slide (slide_index or slide_id) and/or filter with a selector; paginated.
    /// </summary>
    /// <param name="slideIndex">1-based slide position to inspect.</param>
    /// <param name="slideId">PowerPoint SlideID to inspect (stable across reordering).</param>
    /// <param name="selector">Optional filter, e.g. "kind:text-box font&lt;12" or "group:Cards".</param>
    /// <param name="detail">compact (default; text shortened to 200 characters) or full (all text, font names, margins, alt text, tags).</param>
    /// <param name="offset">0-based index of the first object to return (default 0).</param>
    /// <param name="limit">Objects per page (1-200, default 50).</param>
    DeckOperationResult InspectObjects(
        IPresentationBatch batch,
        int? slideIndex = null,
        int? slideId = null,
        string? selector = null,
        string detail = "compact",
        int offset = 0,
        int limit = 50);

    /// <summary>
    /// Finds objects matching a selector across the deck and returns their slide_index,
    /// slide_id, shape_index, and shape_id for use with other tools. With require_single, more
    /// than one match is an error listing the candidates.
    /// </summary>
    /// <param name="selector">Selector, e.g. title:"Revenue" kind:chart, appid:kpi-1, name:"Title 1" slide:2.</param>
    /// <param name="requireSingle">Fail with the candidate list unless exactly one object matches (default false).</param>
    /// <param name="limit">Maximum matches returned (1-200, default 20).</param>
    DeckOperationResult Find(IPresentationBatch batch, string selector, bool requireSingle = false, int limit = 20);

    /// <summary>
    /// Returns the deck revision fingerprint and every slide's fingerprint. Pass the revision as
    /// expected_revision to batch or repair calls to refuse edits on a deck that changed since it
    /// was inspected.
    /// </summary>
    DeckOperationResult Fingerprint(IPresentationBatch batch);

    /// <summary>
    /// Assigns persistent PPTMCP_ID tags: every slide without one, duplicated ids left by copying
    /// (later copies get a ~N suffix), and objects matching the selector. With app_id, the selector
    /// must match exactly one object, which receives that id. Use dry_run to preview.
    /// </summary>
    /// <param name="selector">Objects to give ids, e.g. "kind:chart" (optional).</param>
    /// <param name="appId">Exact id for the single object the selector matches.</param>
    /// <param name="prefix">Prefix for generated ids (default "id").</param>
    /// <param name="dryRun">Only return the plan (default false).</param>
    DeckOperationResult AssignIds(
        IPresentationBatch batch,
        string? selector = null,
        string? appId = null,
        string prefix = "id",
        bool dryRun = false);
}
