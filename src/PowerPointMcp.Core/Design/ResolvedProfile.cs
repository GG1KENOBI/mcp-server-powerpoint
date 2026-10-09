using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>A fully resolved component style: every value present, colors as #RRGGBB.</summary>
public sealed record ResolvedComponent(
    string Name,
    string Version,
    string Fill,
    string? Border,
    float BorderWidth,
    float Radius,
    float Padding,
    string Accent,
    string Text,
    float HeadingSize,
    float BodySize)
{
    /// <summary>Value for the PPTMCP_COMPONENT tag.</summary>
    public string Identity => $"{Name}@{Version}";
}

/// <summary>
/// A design profile with inheritance applied and every token resolved. Pure and immutable;
/// compositions, diagrams, and design apply read from this.
/// </summary>
public sealed partial class ResolvedProfile
{
    private readonly IReadOnlyDictionary<string, string> _colors;
    private readonly IReadOnlyDictionary<string, float> _typeScale;
    private readonly IReadOnlyDictionary<string, float> _spacing;
    private readonly IReadOnlyDictionary<string, ResolvedComponent> _components;

    internal ResolvedProfile(
        string name,
        string version,
        IReadOnlyDictionary<string, string> colors,
        IReadOnlyDictionary<string, string> fonts,
        IReadOnlyDictionary<string, float> typeScale,
        IReadOnlyDictionary<string, float> spacing,
        IReadOnlyDictionary<string, float> margins,
        IReadOnlyDictionary<string, float> areas,
        int gridColumns,
        float gutter,
        float lineSpacing,
        float paragraphSpacing,
        float minFontSize,
        ResolvedTable table,
        ResolvedChart chart,
        string background,
        ResolvedFooter footer,
        ResolvedSource source,
        IReadOnlyDictionary<string, ResolvedComponent> components)
    {
        Name = name;
        Version = version;
        _colors = colors;
        Fonts = fonts;
        _typeScale = typeScale;
        _spacing = spacing;
        Margins = margins;
        Areas = areas;
        GridColumns = gridColumns;
        Gutter = gutter;
        LineSpacing = lineSpacing;
        ParagraphSpacing = paragraphSpacing;
        MinFontSize = minFontSize;
        Table = table;
        Chart = chart;
        Background = background;
        Footer = footer;
        Source = source;
        _components = components;
    }

    /// <summary>Profile name.</summary>
    public string Name { get; }

    /// <summary>Profile version.</summary>
    public string Version { get; }

    /// <summary>All color tokens as #RRGGBB.</summary>
    public IReadOnlyDictionary<string, string> Colors => _colors;

    /// <summary>Font families: heading, body, mono.</summary>
    public IReadOnlyDictionary<string, string> Fonts { get; }

    /// <summary>Font sizes by role.</summary>
    public IReadOnlyDictionary<string, float> TypeScale => _typeScale;

    /// <summary>Spacing tokens.</summary>
    public IReadOnlyDictionary<string, float> SpacingTokens => _spacing;

    /// <summary>Margins: left, right, top, bottom.</summary>
    public IReadOnlyDictionary<string, float> Margins { get; }

    /// <summary>Reserved area heights: title, footer, source.</summary>
    public IReadOnlyDictionary<string, float> Areas { get; }

    /// <summary>Grid columns.</summary>
    public int GridColumns { get; }

    /// <summary>Grid gutter.</summary>
    public float Gutter { get; }

    /// <summary>Body line spacing multiple.</summary>
    public float LineSpacing { get; }

    /// <summary>Space after paragraphs.</summary>
    public float ParagraphSpacing { get; }

    /// <summary>Smallest font size allowed.</summary>
    public float MinFontSize { get; }

    /// <summary>Table style.</summary>
    public ResolvedTable Table { get; }

    /// <summary>Chart style.</summary>
    public ResolvedChart Chart { get; }

    /// <summary>Background color.</summary>
    public string Background { get; }

    /// <summary>Footer configuration.</summary>
    public ResolvedFooter Footer { get; }

    /// <summary>Source line style.</summary>
    public ResolvedSource Source { get; }

    /// <summary>Component styles.</summary>
    public IReadOnlyDictionary<string, ResolvedComponent> Components => _components;

    /// <summary>Heading font family.</summary>
    public string HeadingFont => Fonts["heading"];

    /// <summary>Body font family.</summary>
    public string BodyFont => Fonts["body"];

    /// <summary>Resolves a color token or #RRGGBB value; unknown tokens fall back to the text color.</summary>
    public string Color(string tokenOrHex) =>
        IsHex(tokenOrHex) ? tokenOrHex.ToUpperInvariant() : _colors.GetValueOrDefault(tokenOrHex) ?? _colors["text"];

    /// <summary>Font size for a type-scale role, or the parsed number.</summary>
    public float Size(string roleOrPoints) =>
        float.TryParse(roleOrPoints, NumberStyles.Float, CultureInfo.InvariantCulture, out var points)
            ? points
            : _typeScale.GetValueOrDefault(roleOrPoints, _typeScale["body"]);

    /// <summary>A spacing token value.</summary>
    public float Spacing(string token) => _spacing.GetValueOrDefault(token, _spacing["md"]);

    /// <summary>A component style (falls back to the card style for unknown names).</summary>
    public ResolvedComponent Component(string name) =>
        _components.TryGetValue(name, out var component) ? component : _components["card"] with { Name = name };

    /// <summary>Whether the value is #RRGGBB.</summary>
    public static bool IsHex(string? value) => value is not null && HexColor().IsMatch(value);

    [GeneratedRegex("^#[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant)]
    private static partial Regex HexColor();

    /// <summary>Serializes the resolved values (for export and inspection).</summary>
    public string ToJson() => JsonSerializer.Serialize(new
    {
        schema = DesignProfile.SchemaId,
        name = Name,
        version = Version,
        colors = Colors,
        fonts = Fonts,
        type_scale = TypeScale,
        line_spacing = LineSpacing,
        paragraph_spacing = ParagraphSpacing,
        spacing = SpacingTokens,
        margins = Margins,
        areas = Areas,
        grid = new { columns = GridColumns, gutter = Gutter },
        min_font_size = MinFontSize,
        table = Table,
        chart = Chart,
        background = Background,
        footer = Footer,
        source = Source,
        components = Components,
    }, DesignJson.Options);
}

/// <summary>Resolved table style.</summary>
public sealed record ResolvedTable(
    string HeaderFill, string HeaderText, string? BandFill, string BodyFill, string BodyText, string Border,
    float BorderWidth, string TotalFill, float FontSize, float CellPadding);

/// <summary>Resolved chart style.</summary>
public sealed record ResolvedChart(IReadOnlyList<string> Palette, float FontSize, bool Gridlines, bool DataLabels, string Legend);

/// <summary>Resolved footer configuration.</summary>
public sealed record ResolvedFooter(string? Text, bool SlideNumber, float FontSize, string Color);

/// <summary>Resolved source style.</summary>
public sealed record ResolvedSource(float FontSize, string Color, string Prefix);
