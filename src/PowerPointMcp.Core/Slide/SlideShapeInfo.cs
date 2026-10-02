namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>Read-only snapshot of one top-level shape on a slide.</summary>
public sealed class SlideShapeInfo
{
    /// <summary>1-based index in the slide's Shapes collection (use it as shape_index).</summary>
    public required int ShapeIndex { get; init; }

    /// <summary>Shape name shown in PowerPoint's Selection Pane.</summary>
    public required string Name { get; init; }

    /// <summary>
    /// Shape kind: placeholder, text-box, auto-shape, connector, line, picture, table, chart,
    /// smart-art, group, media, text-effect, freeform, or other.
    /// </summary>
    public required string Kind { get; init; }

    /// <summary>Native PpPlaceholderType member name when <see cref="Kind"/> is placeholder.</summary>
    public string? PlaceholderType { get; init; }

    /// <summary>Horizontal position in points.</summary>
    public required float Left { get; init; }

    /// <summary>Vertical position in points.</summary>
    public required float Top { get; init; }

    /// <summary>Width in points.</summary>
    public required float Width { get; init; }

    /// <summary>Height in points.</summary>
    public required float Height { get; init; }

    /// <summary>Clockwise rotation in degrees.</summary>
    public required float Rotation { get; init; }

    /// <summary>1-based stacking position; higher values are drawn on top.</summary>
    public required int ZOrder { get; init; }

    /// <summary>Whether the shape is visible.</summary>
    public required bool Visible { get; init; }

    /// <summary>Whether the shape has non-empty text.</summary>
    public required bool HasText { get; init; }

    /// <summary>The shape's text, when it has any.</summary>
    public string? Text { get; init; }

    /// <summary>Smallest font size in points used by the text.</summary>
    public float? MinFontSize { get; init; }

    /// <summary>Largest font size in points used by the text.</summary>
    public float? MaxFontSize { get; init; }

    /// <summary>Text autofit: none, shape-to-fit-text, shrink-text-on-overflow, or mixed.</summary>
    public string? AutoSize { get; init; }

    /// <summary>Left edge of the rendered text in points.</summary>
    public float? TextLeft { get; init; }

    /// <summary>Top edge of the rendered text in points.</summary>
    public float? TextTop { get; init; }

    /// <summary>Width of the rendered text in points.</summary>
    public float? TextWidth { get; init; }

    /// <summary>Height of the rendered text in points.</summary>
    public float? TextHeight { get; init; }

    /// <summary>Whether a placeholder is filled with a picture.</summary>
    public bool? HasPicture { get; init; }

    /// <summary>Whether the shape holds a chart.</summary>
    public bool? HasChart { get; init; }

    /// <summary>Row count when the shape holds a table.</summary>
    public int? TableRows { get; init; }

    /// <summary>Column count when the shape holds a table.</summary>
    public int? TableColumns { get; init; }
}
