using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>
/// A design profile (schema <c>pptmcp.design-profile/1</c>): named colors, fonts, a type scale,
/// spacing and margin tokens, table/chart/footer/source styles, and versioned component styles.
/// All lengths and font sizes are points. Color values are #RRGGBB or the name of another color
/// token (for example <c>"header_fill": "primary"</c>). Omitted values inherit from the base
/// profile (<see cref="Base"/>, default "default").
/// </summary>
public sealed class DesignProfile
{
    /// <summary>The schema identifier this class reads and writes.</summary>
    public const string SchemaId = "pptmcp.design-profile/1";

    /// <summary>Schema identifier; must be pptmcp.design-profile/1.</summary>
    public string? Schema { get; set; } = SchemaId;

    /// <summary>Profile name (letters, digits, dash, underscore).</summary>
    public string? Name { get; set; }

    /// <summary>Profile version, e.g. 1.0.0.</summary>
    public string? Version { get; set; }

    /// <summary>What the profile is for.</summary>
    public string? Description { get; set; }

    /// <summary>Profile this one inherits omitted values from (default "default"; "none" inherits nothing).</summary>
    public string? Base { get; set; }

    /// <summary>Color tokens: primary, secondary, accent, positive, negative, neutral, text, muted, background, surface, border, plus any custom names.</summary>
    public Dictionary<string, string>? Colors { get; set; }

    /// <summary>Font families: heading, body, mono.</summary>
    public Dictionary<string, string>? Fonts { get; set; }

    /// <summary>Font sizes by text role: title, subtitle, heading, body, caption, footnote, kpi_value, kpi_label, quote.</summary>
    public Dictionary<string, float>? TypeScale { get; set; }

    /// <summary>Line spacing multiple for body text (1.0-2.0).</summary>
    public float? LineSpacing { get; set; }

    /// <summary>Space after paragraphs, in points.</summary>
    public float? ParagraphSpacing { get; set; }

    /// <summary>Spacing tokens: xs, sm, md, lg, xl.</summary>
    public Dictionary<string, float>? Spacing { get; set; }

    /// <summary>Slide margins: left, right, top, bottom.</summary>
    public Dictionary<string, float>? Margins { get; set; }

    /// <summary>Reserved areas: title (height), footer (height), source (height).</summary>
    public Dictionary<string, float>? Areas { get; set; }

    /// <summary>Grid: columns and gutter.</summary>
    public GridStyle? Grid { get; set; }

    /// <summary>Smallest font size compositions and repairs may use.</summary>
    public float? MinFontSize { get; set; }

    /// <summary>Table style.</summary>
    public TableStyle? Table { get; set; }

    /// <summary>Chart style.</summary>
    public ChartStyle? Chart { get; set; }

    /// <summary>Slide background color token.</summary>
    public string? Background { get; set; }

    /// <summary>Footer text and slide numbers.</summary>
    public FooterStyle? Footer { get; set; }

    /// <summary>Source/citation line style.</summary>
    public SourceStyle? Source { get; set; }

    /// <summary>Component styles by component name (card, kpi, badge, callout, legend, timeline, section-header, quote, process-step, node).</summary>
    public Dictionary<string, ComponentStyle>? Components { get; set; }

    /// <summary>Unknown properties, reported as unsupported during validation.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>Layout grid.</summary>
public sealed class GridStyle
{
    /// <summary>Number of columns (1-24).</summary>
    public int? Columns { get; set; }

    /// <summary>Gap between columns in points.</summary>
    public float? Gutter { get; set; }

    /// <summary>Unknown properties.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>Table styling.</summary>
public sealed class TableStyle
{
    /// <summary>Header row fill color.</summary>
    public string? HeaderFill { get; set; }

    /// <summary>Header row text color.</summary>
    public string? HeaderText { get; set; }

    /// <summary>Fill for alternate body rows (null disables banding).</summary>
    public string? BandFill { get; set; }

    /// <summary>Fill for other body rows.</summary>
    public string? BodyFill { get; set; }

    /// <summary>Body text color.</summary>
    public string? BodyText { get; set; }

    /// <summary>Border color.</summary>
    public string? Border { get; set; }

    /// <summary>Border width in points (0 hides borders).</summary>
    public float? BorderWidth { get; set; }

    /// <summary>Total/summary row fill.</summary>
    public string? TotalFill { get; set; }

    /// <summary>Text size in table cells.</summary>
    public float? FontSize { get; set; }

    /// <summary>Cell padding in points.</summary>
    public float? CellPadding { get; set; }

    /// <summary>Unknown properties.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>Chart styling.</summary>
public sealed class ChartStyle
{
    /// <summary>Series colors in order (tokens or #RRGGBB).</summary>
    public List<string>? Palette { get; set; }

    /// <summary>Chart text size.</summary>
    public float? FontSize { get; set; }

    /// <summary>Show major gridlines on the value axis.</summary>
    public bool? Gridlines { get; set; }

    /// <summary>Show data labels.</summary>
    public bool? DataLabels { get; set; }

    /// <summary>Legend position: bottom, right, top, left, or none.</summary>
    public string? Legend { get; set; }

    /// <summary>Unknown properties.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>Footer configuration.</summary>
public sealed class FooterStyle
{
    /// <summary>Footer text (null leaves footers as they are).</summary>
    public string? Text { get; set; }

    /// <summary>Show slide numbers.</summary>
    public bool? SlideNumber { get; set; }

    /// <summary>Footer text size.</summary>
    public float? FontSize { get; set; }

    /// <summary>Footer text color.</summary>
    public string? Color { get; set; }

    /// <summary>Unknown properties.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>Source / citation line style.</summary>
public sealed class SourceStyle
{
    /// <summary>Text size.</summary>
    public float? FontSize { get; set; }

    /// <summary>Text color.</summary>
    public string? Color { get; set; }

    /// <summary>Prefix added when the source text does not start with it (e.g. "Source: ").</summary>
    public string? Prefix { get; set; }

    /// <summary>Unknown properties.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>Style of one component.</summary>
public sealed class ComponentStyle
{
    /// <summary>Component version recorded in the PPTMCP_COMPONENT tag (name@version).</summary>
    public string? Version { get; set; }

    /// <summary>Fill color.</summary>
    public string? Fill { get; set; }

    /// <summary>Border color (null for none).</summary>
    public string? Border { get; set; }

    /// <summary>Border width in points.</summary>
    public float? BorderWidth { get; set; }

    /// <summary>Corner radius in points (0 for square corners).</summary>
    public float? Radius { get; set; }

    /// <summary>Inner padding in points.</summary>
    public float? Padding { get; set; }

    /// <summary>Accent color (bars, markers, badges).</summary>
    public string? Accent { get; set; }

    /// <summary>Text color.</summary>
    public string? Text { get; set; }

    /// <summary>Heading text size (number of points or a type_scale role name).</summary>
    public string? HeadingSize { get; set; }

    /// <summary>Body text size (number of points or a type_scale role name).</summary>
    public string? BodySize { get; set; }

    /// <summary>Unknown properties.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Unknown { get; set; }
}

/// <summary>JSON options for profiles and compositions: snake_case names, comments and trailing commas allowed.</summary>
public static class DesignJson
{
    /// <summary>Shared options.</summary>
    public static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        DictionaryKeyPolicy = null,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
    };
}
