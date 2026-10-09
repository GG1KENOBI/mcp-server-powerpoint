using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>A paragraph of planned text. Text is exactly the spec's text; nothing is shortened.</summary>
public sealed class PlannedParagraph
{
    /// <summary>Paragraph text.</summary>
    public required string Text { get; init; }

    /// <summary>Font size in points.</summary>
    public required float Size { get; init; }

    /// <summary>Font family.</summary>
    public required string Font { get; init; }

    /// <summary>#RRGGBB text color.</summary>
    public required string Color { get; init; }

    /// <summary>Bold.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic.</summary>
    public bool Italic { get; init; }

    /// <summary>Indent level 0-4.</summary>
    public int Level { get; init; }

    /// <summary>none, bullet, or number (native PowerPoint numbering; no characters are added to the text).</summary>
    public string Bullet { get; init; } = "none";

    /// <summary>left, center, or right.</summary>
    public string Align { get; init; } = "left";

    /// <summary>Space after the paragraph in points.</summary>
    public float SpaceAfter { get; init; }

    /// <summary>Hyperlink address applied to the whole paragraph.</summary>
    public string? Hyperlink { get; init; }
}

/// <summary>Table content and style for one slide.</summary>
public sealed record TablePlan(
    IReadOnlyList<string> Header,
    IReadOnlyList<IReadOnlyList<string>> Rows,
    IReadOnlyList<float> ColumnWidths,
    IReadOnlyList<string> Align,
    bool TotalRow,
    IReadOnlyList<int> HighlightRows,
    float FontSize,
    string Font,
    ResolvedTable Style,
    int FirstRowNumber);

/// <summary>Chart content and look.</summary>
public sealed record ChartPlan(
    SpecChart Spec,
    IReadOnlyList<string> Palette,
    string Font,
    float FontSize,
    string TextColor,
    bool Gridlines,
    bool DataLabels,
    string Legend);

/// <summary>A local picture placement.</summary>
public sealed record PicturePlan(string Path, PicturePlacement Placement, string? Alt, int PixelWidth, int PixelHeight, string Fit, string? Attribution);

/// <summary>One native object to create.</summary>
public sealed class PlannedElement
{
    /// <summary>Semantic key, unique on the slide (e.g. card-2.body). Becomes the shape name and part of its PPTMCP_ID.</summary>
    public required string Key { get; init; }

    /// <summary>Semantic role (title, subtitle, body, card, kpi-value, source, takeaway, decoration, ...).</summary>
    public required string Role { get; init; }

    /// <summary>text, shape, line, table, chart, picture, or connector.</summary>
    public required string Type { get; init; }

    /// <summary>Bounds in points.</summary>
    public required Box Box { get; set; }

    /// <summary>Component identity (name@version) for component parts.</summary>
    public string? Component { get; init; }

    /// <summary>Group key: elements with the same group are grouped after creation.</summary>
    public string? Group { get; init; }

    /// <summary>Shape geometry for shape elements: rectangle, rounded-rectangle, oval, chevron, pentagon, triangle.</summary>
    public string Geometry { get; init; } = "rectangle";

    /// <summary>Fill color (#RRGGBB) or null for no fill.</summary>
    public string? Fill { get; init; }

    /// <summary>Outline color or null for no outline.</summary>
    public string? LineColor { get; init; }

    /// <summary>Outline width in points.</summary>
    public float LineWidth { get; init; }

    /// <summary>Corner radius for rounded rectangles, in points.</summary>
    public float Radius { get; init; }

    /// <summary>Rotation in degrees.</summary>
    public float Rotation { get; init; }

    /// <summary>Text paragraphs (text elements, and shapes that hold text).</summary>
    public IReadOnlyList<PlannedParagraph>? Paragraphs { get; set; }

    /// <summary>Vertical anchor: top, middle, or bottom.</summary>
    public string VAlign { get; init; } = "top";

    /// <summary>Inner padding (left, top, right, bottom) in points.</summary>
    public IReadOnlyList<float> Padding { get; init; } = [0, 0, 0, 0];

    /// <summary>Line spacing multiple.</summary>
    public float LineSpacing { get; init; } = 1f;

    /// <summary>Render into the slide's title placeholder instead of a new text box.</summary>
    public bool TitlePlaceholder { get; init; }

    /// <summary>Fit region this element belongs to.</summary>
    public string? Region { get; init; }

    /// <summary>Key of an element whose top this element follows when its region is restacked.</summary>
    public string? AnchorKey { get; init; }

    /// <summary>Vertical offset from the anchor's top.</summary>
    public float AnchorOffset { get; init; }

    /// <summary>Estimated text height in points (planning only).</summary>
    public float? EstimatedTextHeight { get; init; }

    /// <summary>Table content.</summary>
    public TablePlan? Table { get; init; }

    /// <summary>Chart content.</summary>
    public ChartPlan? Chart { get; init; }

    /// <summary>Picture placement.</summary>
    public PicturePlan? Picture { get; init; }

    /// <summary>Connector start element key.</summary>
    public string? FromKey { get; init; }

    /// <summary>Connector end element key.</summary>
    public string? ToKey { get; init; }

    /// <summary>Connection site on the start shape (rectangles: 1 top, 2 left, 3 bottom, 4 right).</summary>
    public int FromSite { get; init; } = 3;

    /// <summary>Connection site on the end shape.</summary>
    public int ToSite { get; init; } = 1;

    /// <summary>elbow or straight.</summary>
    public string ConnectorKind { get; init; } = "elbow";

    /// <summary>Arrowhead at the end of a connector or line.</summary>
    public bool EndArrow { get; init; }

    /// <summary>Extra tags (e.g. PPTMCP_NODE).</summary>
    public IReadOnlyDictionary<string, string>? Tags { get; init; }

    /// <summary>Alternative text.</summary>
    public string? AltText { get; init; }
}

/// <summary>A group of elements that are fitted together.</summary>
public sealed class FitRegion
{
    /// <summary>Region key.</summary>
    public required string Key { get; init; }

    /// <summary>Available area.</summary>
    public required Box Box { get; init; }

    /// <summary>stack (restack vertically by measured height), each (every element fits its own box), paragraphs (one text box, split by paragraph), or table.</summary>
    public required string Mode { get; init; }

    /// <summary>Gap between stacked elements.</summary>
    public float Gap { get; init; }

    /// <summary>Element keys in order.</summary>
    public required IReadOnlyList<string> Keys { get; init; }

    /// <summary>Whether overflow may move trailing items to a continuation slide.</summary>
    public bool Splittable { get; init; }

    /// <summary>Font floor for shrinking.</summary>
    public required float MinFontSize { get; init; }
}

/// <summary>One slide of a plan.</summary>
public sealed class PlannedSlide
{
    /// <summary>0-based position within the plan.</summary>
    public required int Sequence { get; init; }

    /// <summary>Persistent id for the slide.</summary>
    public required string AppId { get; init; }

    /// <summary>Title text (null for no title placeholder).</summary>
    public string? Title { get; init; }

    /// <summary>Whether this slide continues the previous one.</summary>
    public bool Continuation { get; init; }

    /// <summary>First item (point, row, or reference) on this slide, 0-based.</summary>
    public int ItemOffset { get; init; }

    /// <summary>Number of items on this slide.</summary>
    public int ItemCount { get; init; }

    /// <summary>Elements in creation (z) order.</summary>
    public required IReadOnlyList<PlannedElement> Elements { get; init; }

    /// <summary>Fit regions.</summary>
    public required IReadOnlyList<FitRegion> Regions { get; init; }

    /// <summary>Speaker notes.</summary>
    public string? Notes { get; init; }
}

/// <summary>A complete plan: slides, profile, and planning warnings.</summary>
public sealed class CompositionPlan
{
    /// <summary>The spec.</summary>
    public required CompositionSpec Spec { get; init; }

    /// <summary>Slides to create.</summary>
    public required IReadOnlyList<PlannedSlide> Slides { get; init; }

    /// <summary>Profile used.</summary>
    public required string ProfileName { get; init; }

    /// <summary>Slide width.</summary>
    public required float Width { get; init; }

    /// <summary>Slide height.</summary>
    public required float Height { get; init; }

    /// <summary>What split items are called (points, rows, references), when splittable.</summary>
    public string? ItemName { get; init; }

    /// <summary>Planning notes: estimated overflow, splits, accessibility.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}

/// <summary>Inputs the planner needs from outside the spec.</summary>
public sealed class PlanContext
{
    /// <summary>Title placeholder bounds from the deck's layout, used instead of the profile's title area when set.</summary>
    public Box? TemplateTitle { get; init; }

    /// <summary>Pixel size of images by path.</summary>
    public IReadOnlyDictionary<string, (int Width, int Height)> ImageSizes { get; init; } = new Dictionary<string, (int, int)>();

    /// <summary>Forced item counts per slide (from a measured split); null lets the planner decide.</summary>
    public IReadOnlyList<int>? Partition { get; init; }

    /// <summary>Base persistent id for the slide(s).</summary>
    public required string AppId { get; init; }
}
