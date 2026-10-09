namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Read-only snapshot of one object on a slide. Geometry is in points (1/72 inch), PowerPoint's
/// native unit and the canonical unit of every deck, compose, and design operation.
/// </summary>
/// <remarks>
/// <see cref="ShapeId"/> is PowerPoint's <c>Shape.Id</c>: unique within its slide and stable when
/// shapes are reordered or other shapes are deleted. <see cref="AppId"/> is the optional persistent
/// <c>PPTMCP_ID</c> tag; unlike <c>Shape.Id</c> it survives copying, so copies can share it until
/// <c>deck assign-ids</c> repairs the duplicates.
/// </remarks>
public sealed class DeckObjectInfo
{
    /// <summary>1-based position of the slide holding the object.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID of the slide holding the object (stable across reordering).</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based index in the slide's Shapes collection; null for objects inside a group.</summary>
    public int? ShapeIndex { get; init; }

    /// <summary>PowerPoint Shape.Id, unique within the slide.</summary>
    public required int ShapeId { get; init; }

    /// <summary>Selection Pane name.</summary>
    public required string Name { get; init; }

    /// <summary>
    /// placeholder, text-box, auto-shape, connector, line, picture, table, chart, smart-art,
    /// group, media, text-effect, freeform, or other.
    /// </summary>
    public required string Kind { get; init; }

    /// <summary>
    /// Semantic role: the PPTMCP_ROLE tag when present, otherwise derived from the placeholder
    /// type (title, subtitle, body, footer, slide-number, date, picture, chart, table, object).
    /// </summary>
    public string? Role { get; init; }

    /// <summary>Native PpPlaceholderType member name for placeholders.</summary>
    public string? PlaceholderType { get; init; }

    /// <summary>Persistent application identifier (PPTMCP_ID tag).</summary>
    public string? AppId { get; init; }

    /// <summary>Component identity "name@version" (PPTMCP_COMPONENT tag) for reusable components.</summary>
    public string? Component { get; init; }

    /// <summary>Shape.Id of the enclosing group, for objects inside a group.</summary>
    public int? GroupShapeId { get; init; }

    /// <summary>Shape.Id values of the group's members, for groups.</summary>
    public IReadOnlyList<int>? MemberShapeIds { get; init; }

    /// <summary>Left edge in points.</summary>
    public required float Left { get; init; }

    /// <summary>Top edge in points.</summary>
    public required float Top { get; init; }

    /// <summary>Width in points.</summary>
    public required float Width { get; init; }

    /// <summary>Height in points.</summary>
    public required float Height { get; init; }

    /// <summary>Clockwise rotation in degrees.</summary>
    public required float Rotation { get; init; }

    /// <summary>1-based stacking position among top-level shapes (higher draws on top).</summary>
    public required int ZOrder { get; init; }

    /// <summary>Whether the object is visible.</summary>
    public required bool Visible { get; init; }

    /// <summary>Whether the object holds non-empty text.</summary>
    public required bool HasText { get; init; }

    /// <summary>Text (full in detailed mode, at most 200 characters in compact mode).</summary>
    public string? Text { get; init; }

    /// <summary>Whether <see cref="Text"/> was shortened for the compact response.</summary>
    public bool? TextTruncated { get; init; }

    /// <summary>Number of paragraphs in the text.</summary>
    public int? ParagraphCount { get; init; }

    /// <summary>Smallest font size in points used by the text.</summary>
    public float? MinFontSize { get; init; }

    /// <summary>Largest font size in points used by the text.</summary>
    public float? MaxFontSize { get; init; }

    /// <summary>Distinct font names used by the text (detailed mode).</summary>
    public IReadOnlyList<string>? FontNames { get; init; }

    /// <summary>Text autofit: none, shape-to-fit-text, shrink-text-on-overflow, or mixed.</summary>
    public string? AutoSize { get; init; }

    /// <summary>Whether text wraps inside the shape.</summary>
    public bool? WordWrap { get; init; }

    /// <summary>Rendered text bounds measured by PowerPoint (left, top, width, height in points).</summary>
    public IReadOnlyList<float>? TextBounds { get; init; }

    /// <summary>Text-frame margins (left, top, right, bottom in points; detailed mode).</summary>
    public IReadOnlyList<float>? TextMargins { get; init; }

    /// <summary>Whether a placeholder is filled with a picture.</summary>
    public bool? HasPicture { get; init; }

    /// <summary>Table rows, for tables.</summary>
    public int? TableRows { get; init; }

    /// <summary>Table columns, for tables.</summary>
    public int? TableColumns { get; init; }

    /// <summary>XlChartType member value, for charts.</summary>
    public int? ChartType { get; init; }

    /// <summary>Linked source file for linked pictures or media.</summary>
    public string? LinkSource { get; init; }

    /// <summary>Shape.Id the connector's start is glued to, for attached connectors.</summary>
    public int? ConnectorBeginShapeId { get; init; }

    /// <summary>Shape.Id the connector's end is glued to, for attached connectors.</summary>
    public int? ConnectorEndShapeId { get; init; }

    /// <summary>Alternative text.</summary>
    public string? AltText { get; init; }

    /// <summary>All string tags (detailed mode).</summary>
    public IReadOnlyDictionary<string, string>? Tags { get; init; }
}
