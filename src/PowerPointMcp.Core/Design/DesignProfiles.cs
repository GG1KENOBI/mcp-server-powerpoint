using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>Outcome of validating or resolving a profile.</summary>
public sealed class ProfileValidation
{
    /// <summary>Problems that make the profile unusable (path: message).</summary>
    public required IReadOnlyList<string> Errors { get; init; }

    /// <summary>Unsupported or suspicious values that are ignored or adjusted.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }

    /// <summary>The resolved profile when there are no errors.</summary>
    public ResolvedProfile? Resolved { get; init; }

    /// <summary>Whether the profile is usable.</summary>
    public bool IsValid => Errors.Count == 0 && Resolved is not null;
}

/// <summary>Built-in design profiles, parsing, validation, and inheritance. Pure.</summary>
public static partial class DesignProfiles
{
    /// <summary>Component names compositions use.</summary>
    public static readonly IReadOnlyList<string> KnownComponents =
        ["card", "kpi", "badge", "callout", "legend", "timeline", "section-header", "quote", "process-step", "node"];

    private static readonly string[] RequiredColors =
        ["primary", "secondary", "accent", "positive", "negative", "neutral", "text", "muted", "background", "surface", "border"];

    private static readonly string[] RequiredTypeRoles =
        ["title", "subtitle", "heading", "body", "caption", "footnote", "kpi_value", "kpi_label", "quote"];

    private static readonly string[] LegendPositions = ["bottom", "right", "top", "left", "none"];

    /// <summary>Names of the built-in profiles.</summary>
    public static IReadOnlyList<string> BuiltInNames => [.. BuiltIn.Keys];

    private static readonly Dictionary<string, Func<DesignProfile>> BuiltIn = new(StringComparer.OrdinalIgnoreCase)
    {
        ["default"] = Default,
        ["corporate-blue"] = () => new DesignProfile
        {
            Name = "corporate-blue",
            Version = "1.0.0",
            Description = "Navy and blue business style with Segoe UI.",
            Colors = new() { ["primary"] = "#1F3864", ["secondary"] = "#2E75B6", ["accent"] = "#ED7D31", ["surface"] = "#EEF3FA" },
        },
        ["high-contrast"] = () => new DesignProfile
        {
            Name = "high-contrast",
            Version = "1.0.0",
            Description = "Black on white with strong accents and larger text, for projection and accessibility.",
            Colors = new()
            {
                ["primary"] = "#000000",
                ["secondary"] = "#003A8C",
                ["accent"] = "#B00020",
                ["text"] = "#000000",
                ["muted"] = "#333333",
                ["surface"] = "#F2F2F2",
                ["border"] = "#000000",
            },
            TypeScale = new() { ["body"] = 20, ["caption"] = 14, ["footnote"] = 12, ["heading"] = 24 },
            MinFontSize = 14,
        },
        ["minimal-mono"] = () => new DesignProfile
        {
            Name = "minimal-mono",
            Version = "1.0.0",
            Description = "Greyscale with one accent, generous margins.",
            Colors = new() { ["primary"] = "#222222", ["secondary"] = "#555555", ["accent"] = "#0A84FF", ["surface"] = "#F5F5F5", ["border"] = "#DDDDDD" },
            Margins = new() { ["left"] = 64, ["right"] = 64, ["top"] = 44, ["bottom"] = 40 },
        },
    };

    /// <summary>The root profile every other profile inherits from.</summary>
    public static DesignProfile Default() => new()
    {
        Name = "default",
        Version = "1.0.0",
        Description = "Neutral business style: Segoe UI, blue primary, 16 pt body text, 48 pt side margins.",
        Base = "none",
        Colors = new()
        {
            ["primary"] = "#1F4E79",
            ["secondary"] = "#2E75B6",
            ["accent"] = "#C55A11",
            ["positive"] = "#2E7D32",
            ["negative"] = "#C62828",
            ["neutral"] = "#7F7F7F",
            ["text"] = "#1A1A1A",
            ["muted"] = "#595959",
            ["background"] = "#FFFFFF",
            ["surface"] = "#F2F4F7",
            ["border"] = "#D0D5DD",
        },
        Fonts = new() { ["heading"] = "Segoe UI Semibold", ["body"] = "Segoe UI", ["mono"] = "Consolas" },
        TypeScale = new()
        {
            ["title"] = 30,
            ["subtitle"] = 20,
            ["heading"] = 18,
            ["body"] = 16,
            ["caption"] = 12,
            ["footnote"] = 10,
            ["kpi_value"] = 36,
            ["kpi_label"] = 12,
            ["quote"] = 26,
        },
        LineSpacing = 1.1f,
        ParagraphSpacing = 6,
        Spacing = new() { ["xs"] = 4, ["sm"] = 8, ["md"] = 16, ["lg"] = 24, ["xl"] = 40 },
        Margins = new() { ["left"] = 48, ["right"] = 48, ["top"] = 32, ["bottom"] = 28 },
        Areas = new() { ["title"] = 64, ["footer"] = 22, ["source"] = 18 },
        Grid = new GridStyle { Columns = 12, Gutter = 16 },
        MinFontSize = 12,
        Table = new TableStyle
        {
            HeaderFill = "primary",
            HeaderText = "#FFFFFF",
            BandFill = "surface",
            BodyFill = "background",
            BodyText = "text",
            Border = "border",
            BorderWidth = 0.75f,
            TotalFill = "#E4E9F0",
            FontSize = 12,
            CellPadding = 4,
        },
        Chart = new ChartStyle { Palette = ["primary", "secondary", "accent", "neutral", "positive", "negative"], FontSize = 12, Gridlines = false, DataLabels = true, Legend = "bottom" },
        Background = "background",
        Footer = new FooterStyle { Text = null, SlideNumber = true, FontSize = 10, Color = "muted" },
        Source = new SourceStyle { FontSize = 10, Color = "muted", Prefix = "Source: " },
        Components = new()
        {
            ["card"] = new() { Version = "1", Fill = "surface", Border = "border", BorderWidth = 0.75f, Radius = 6, Padding = 14, Accent = "primary", Text = "text", HeadingSize = "heading", BodySize = "body" },
            ["kpi"] = new() { Version = "1", Fill = "surface", Border = null, BorderWidth = 0, Radius = 6, Padding = 14, Accent = "primary", Text = "text", HeadingSize = "kpi_value", BodySize = "kpi_label" },
            ["badge"] = new() { Version = "1", Fill = "accent", Border = null, BorderWidth = 0, Radius = 10, Padding = 4, Accent = "accent", Text = "#FFFFFF", HeadingSize = "caption", BodySize = "caption" },
            ["callout"] = new() { Version = "1", Fill = "#FFF4E5", Border = "accent", BorderWidth = 1, Radius = 4, Padding = 12, Accent = "accent", Text = "text", HeadingSize = "heading", BodySize = "body" },
            ["legend"] = new() { Version = "1", Fill = "background", Border = null, BorderWidth = 0, Radius = 0, Padding = 4, Accent = "primary", Text = "muted", HeadingSize = "caption", BodySize = "caption" },
            ["timeline"] = new() { Version = "1", Fill = "background", Border = "border", BorderWidth = 2, Radius = 0, Padding = 8, Accent = "primary", Text = "text", HeadingSize = "body", BodySize = "caption" },
            ["section-header"] = new() { Version = "1", Fill = "primary", Border = null, BorderWidth = 0, Radius = 0, Padding = 24, Accent = "accent", Text = "#FFFFFF", HeadingSize = "title", BodySize = "subtitle" },
            ["quote"] = new() { Version = "1", Fill = "background", Border = null, BorderWidth = 0, Radius = 0, Padding = 16, Accent = "primary", Text = "text", HeadingSize = "quote", BodySize = "body" },
            ["process-step"] = new() { Version = "1", Fill = "primary", Border = null, BorderWidth = 0, Radius = 4, Padding = 10, Accent = "secondary", Text = "#FFFFFF", HeadingSize = "body", BodySize = "caption" },
            ["node"] = new() { Version = "1", Fill = "surface", Border = "primary", BorderWidth = 1.25f, Radius = 4, Padding = 8, Accent = "primary", Text = "text", HeadingSize = "body", BodySize = "caption" },
        },
    };

    /// <summary>Whether the name is a built-in profile.</summary>
    public static bool IsBuiltIn(string name) => BuiltIn.ContainsKey(name);

    /// <summary>Returns a copy of a built-in profile definition, or null.</summary>
    public static DesignProfile? GetBuiltIn(string name) => BuiltIn.TryGetValue(name, out var factory) ? factory() : null;

    /// <summary>Parses profile JSON. Throws <see cref="JsonException"/> with the position of a syntax error.</summary>
    public static DesignProfile Parse(string json) =>
        JsonSerializer.Deserialize<DesignProfile>(json, DesignJson.Options) ?? throw new JsonException("The profile JSON is empty.");

    /// <summary>Serializes a profile definition.</summary>
    public static string ToJson(DesignProfile profile) => JsonSerializer.Serialize(profile, DesignJson.Options);

    /// <summary>
    /// Validates a profile and resolves it against its base chain. <paramref name="lookup"/> finds
    /// user profiles by name (built-ins are found automatically).
    /// </summary>
    public static ProfileValidation Resolve(DesignProfile profile, Func<string, DesignProfile?>? lookup = null)
    {
        ArgumentNullException.ThrowIfNull(profile);
        var errors = new List<string>();
        var warnings = new List<string>();

        if (profile.Schema is not null && profile.Schema != DesignProfile.SchemaId)
            errors.Add($"schema: expected {DesignProfile.SchemaId}, got {profile.Schema}.");
        if (string.IsNullOrWhiteSpace(profile.Name) || !ProfileName().IsMatch(profile.Name))
            errors.Add("name: required; letters, digits, dash, and underscore only (max 64).");

        // Build the inheritance chain, root first.
        var chain = new List<DesignProfile> { profile };
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { profile.Name ?? "" };
        var current = profile;
        while (true)
        {
            var baseName = current.Base ?? (string.Equals(current.Name, "default", StringComparison.OrdinalIgnoreCase) ? "none" : "default");
            if (string.Equals(baseName, "none", StringComparison.OrdinalIgnoreCase))
                break;
            if (!seen.Add(baseName))
            {
                errors.Add($"base: inheritance cycle through '{baseName}'.");
                break;
            }
            var next = lookup?.Invoke(baseName) ?? GetBuiltIn(baseName);
            if (next is null)
            {
                errors.Add($"base: profile '{baseName}' was not found.");
                break;
            }
            chain.Insert(0, next);
            current = next;
        }

        foreach (var item in chain)
            ReportUnknown(item, warnings);

        var colors = MergeDictionaries(chain.Select(item => item.Colors));
        var fonts = MergeDictionaries(chain.Select(item => item.Fonts));
        var typeScale = MergeDictionaries(chain.Select(item => item.TypeScale));
        var spacing = MergeDictionaries(chain.Select(item => item.Spacing));
        var margins = MergeDictionaries(chain.Select(item => item.Margins));
        var areas = MergeDictionaries(chain.Select(item => item.Areas));
        var components = new Dictionary<string, ComponentStyle>(StringComparer.Ordinal);
        foreach (var item in chain.Where(item => item.Components is not null))
        {
            foreach (var (name, style) in item.Components!)
                components[name] = components.TryGetValue(name, out var existing) ? Merge(existing, style) : Clone(style);
        }

        T? Last<T>(Func<DesignProfile, T?> selector) where T : class => chain.Select(selector).LastOrDefault(value => value is not null);
        T? LastValue<T>(Func<DesignProfile, T?> selector) where T : struct => chain.Select(selector).LastOrDefault(value => value is not null);
        var table = MergeObjects(chain.Select(item => item.Table));
        var chart = MergeObjects(chain.Select(item => item.Chart));
        var footer = MergeObjects(chain.Select(item => item.Footer));
        var source = MergeObjects(chain.Select(item => item.Source));
        var grid = MergeObjects(chain.Select(item => item.Grid));

        // Resolve color tokens (references to other tokens allowed, no cycles).
        var resolvedColors = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var name in colors.Keys)
        {
            if (ResolveColor(name, colors, 0) is { } hex)
                resolvedColors[name] = hex;
            else
                errors.Add($"colors.{name}: '{colors[name]}' is neither #RRGGBB nor a defined color token (or the references form a cycle).");
        }
        foreach (var required in RequiredColors.Where(required => !resolvedColors.ContainsKey(required)))
            errors.Add($"colors.{required}: required.");

        string ColorOf(string path, string? value, string fallback)
        {
            var token = value ?? fallback;
            if (ResolvedProfile.IsHex(token))
                return token.ToUpperInvariant();
            if (resolvedColors.TryGetValue(token, out var hex))
                return hex;
            errors.Add($"{path}: '{token}' is neither #RRGGBB nor a color token.");
            return "#000000";
        }

        foreach (var role in new[] { "heading", "body" })
        {
            if (!fonts.TryGetValue(role, out var family) || string.IsNullOrWhiteSpace(family))
                errors.Add($"fonts.{role}: required.");
        }
        if (!fonts.ContainsKey("mono"))
            fonts["mono"] = "Consolas";

        foreach (var role in RequiredTypeRoles)
        {
            if (!typeScale.TryGetValue(role, out var size))
                errors.Add($"type_scale.{role}: required.");
            else if (size is < 6 or > 200)
                errors.Add($"type_scale.{role}: {Fmt(size)} pt is outside 6-200.");
        }
        foreach (var token in new[] { "xs", "sm", "md", "lg", "xl" })
        {
            if (!spacing.TryGetValue(token, out var value))
                errors.Add($"spacing.{token}: required.");
            else if (value is < 0 or > 200)
                errors.Add($"spacing.{token}: {Fmt(value)} is outside 0-200.");
        }
        foreach (var side in new[] { "left", "right", "top", "bottom" })
        {
            if (!margins.TryGetValue(side, out var value))
                errors.Add($"margins.{side}: required.");
            else if (value is < 0 or > 200)
                errors.Add($"margins.{side}: {Fmt(value)} is outside 0-200.");
        }
        foreach (var area in new[] { "title", "footer", "source" })
        {
            if (!areas.TryGetValue(area, out var value))
                errors.Add($"areas.{area}: required.");
            else if (value is < 0 or > 300)
                errors.Add($"areas.{area}: {Fmt(value)} is outside 0-300.");
        }

        var minFont = LastValue(item => item.MinFontSize) ?? 12f;
        if (minFont is < 6 or > 40)
            errors.Add($"min_font_size: {Fmt(minFont)} is outside 6-40.");
        if (typeScale.TryGetValue("body", out var body) && body < minFont)
            errors.Add($"type_scale.body ({Fmt(body)}) is below min_font_size ({Fmt(minFont)}).");
        foreach (var (role, size) in typeScale.Where(entry => entry.Value < minFont && entry.Key is not ("footnote" or "caption" or "kpi_label")))
            warnings.Add($"type_scale.{role} ({Fmt(size)}) is below min_font_size ({Fmt(minFont)}).");

        var lineSpacing = LastValue(item => item.LineSpacing) ?? 1.1f;
        if (lineSpacing is < 0.8f or > 3f)
            errors.Add($"line_spacing: {Fmt(lineSpacing)} is outside 0.8-3.");
        var paragraphSpacing = LastValue(item => item.ParagraphSpacing) ?? 6f;
        if (paragraphSpacing is < 0 or > 72)
            errors.Add($"paragraph_spacing: {Fmt(paragraphSpacing)} is outside 0-72.");

        int columns = grid?.Columns ?? 12;
        float gutter = grid?.Gutter ?? 16;
        if (columns is < 1 or > 24)
            errors.Add($"grid.columns: {columns} is outside 1-24.");
        if (gutter is < 0 or > 100)
            errors.Add($"grid.gutter: {Fmt(gutter)} is outside 0-100.");

        var legend = chart?.Legend ?? "bottom";
        if (Array.IndexOf(LegendPositions, legend) < 0)
            errors.Add($"chart.legend: '{legend}' must be one of {string.Join(", ", LegendPositions)}.");

        var resolvedTable = new ResolvedTable(
            ColorOf("table.header_fill", table?.HeaderFill, "primary"),
            ColorOf("table.header_text", table?.HeaderText, "background"),
            table?.BandFill is null ? null : ColorOf("table.band_fill", table.BandFill, "surface"),
            ColorOf("table.body_fill", table?.BodyFill, "background"),
            ColorOf("table.body_text", table?.BodyText, "text"),
            ColorOf("table.border", table?.Border, "border"),
            table?.BorderWidth ?? 0.75f,
            ColorOf("table.total_fill", table?.TotalFill, "surface"),
            table?.FontSize ?? 12f,
            table?.CellPadding ?? 4f);
        if (resolvedTable.FontSize < minFont - 0.01f)
            warnings.Add($"table.font_size ({Fmt(resolvedTable.FontSize)}) is below min_font_size ({Fmt(minFont)}).");

        var palette = (chart?.Palette ?? ["primary", "secondary", "accent", "neutral"])
            .Select((token, index) => ColorOf($"chart.palette[{index}]", token, "primary")).ToList();
        var resolvedChart = new ResolvedChart(palette, chart?.FontSize ?? 12f, chart?.Gridlines ?? false, chart?.DataLabels ?? true, legend);
        var resolvedFooter = new ResolvedFooter(footer?.Text, footer?.SlideNumber ?? true, footer?.FontSize ?? 10f, ColorOf("footer.color", footer?.Color, "muted"));
        var resolvedSource = new ResolvedSource(source?.FontSize ?? 10f, ColorOf("source.color", source?.Color, "muted"), source?.Prefix ?? "Source: ");

        var resolvedComponents = new Dictionary<string, ResolvedComponent>(StringComparer.Ordinal);
        foreach (var (name, style) in components)
        {
            if (!ComponentName().IsMatch(name))
            {
                errors.Add($"components.{name}: names use lowercase letters, digits, and dashes.");
                continue;
            }
            if (!KnownComponents.Contains(name))
                warnings.Add($"components.{name}: not used by built-in compositions; available for your own tagging.");
            float SizeOf(string path, string? value, string fallback)
            {
                var token = value ?? fallback;
                if (float.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var points))
                    return points;
                if (typeScale.TryGetValue(token, out var size))
                    return size;
                errors.Add($"{path}: '{token}' is neither a number nor a type_scale role.");
                return 12f;
            }
            resolvedComponents[name] = new ResolvedComponent(
                name,
                style.Version ?? "1",
                ColorOf($"components.{name}.fill", style.Fill, "surface"),
                style.Border is null ? null : ColorOf($"components.{name}.border", style.Border, "border"),
                style.BorderWidth ?? 0.75f,
                style.Radius ?? 0f,
                style.Padding ?? 12f,
                ColorOf($"components.{name}.accent", style.Accent, "primary"),
                ColorOf($"components.{name}.text", style.Text, "text"),
                SizeOf($"components.{name}.heading_size", style.HeadingSize, "heading"),
                SizeOf($"components.{name}.body_size", style.BodySize, "body"));
        }
        if (!resolvedComponents.ContainsKey("card"))
            errors.Add("components.card: required (it is the fallback component).");

        var background = ColorOf("background", Last(item => item.Background), "background");
        if (errors.Count > 0)
            return new ProfileValidation { Errors = errors, Warnings = warnings };

        return new ProfileValidation
        {
            Errors = errors,
            Warnings = warnings,
            Resolved = new ResolvedProfile(
                profile.Name!,
                profile.Version ?? "1.0.0",
                resolvedColors,
                fonts,
                typeScale,
                spacing,
                margins,
                areas,
                columns,
                gutter,
                lineSpacing,
                paragraphSpacing,
                minFont,
                resolvedTable,
                resolvedChart,
                background,
                resolvedFooter,
                resolvedSource,
                resolvedComponents),
        };
    }

    /// <summary>Resolves a built-in profile (never fails).</summary>
    public static ResolvedProfile ResolveBuiltIn(string name) =>
        Resolve(GetBuiltIn(name) ?? throw new ArgumentException($"Unknown built-in profile '{name}'.", nameof(name))).Resolved!;

    private static string? ResolveColor(string name, IReadOnlyDictionary<string, string> colors, int depth)
    {
        if (depth > 8 || !colors.TryGetValue(name, out var value))
            return null;
        if (ResolvedProfile.IsHex(value))
            return value.ToUpperInvariant();
        return ResolveColor(value, colors, depth + 1);
    }

    private static Dictionary<string, TValue> MergeDictionaries<TValue>(IEnumerable<Dictionary<string, TValue>?> layers)
    {
        var result = new Dictionary<string, TValue>(StringComparer.Ordinal);
        foreach (var layer in layers.Where(layer => layer is not null))
        {
            foreach (var (key, value) in layer!)
                result[key] = value;
        }
        return result;
    }

    /// <summary>Merges objects property by property: later non-null values win.</summary>
    private static T? MergeObjects<T>(IEnumerable<T?> layers) where T : class, new()
    {
        T? result = null;
        foreach (var layer in layers.Where(layer => layer is not null))
        {
            result ??= new T();
            foreach (var property in typeof(T).GetProperties().Where(property => property.CanWrite && property.Name != "Unknown"))
            {
                var value = property.GetValue(layer);
                if (value is not null)
                    property.SetValue(result, value);
            }
        }
        return result;
    }

    private static ComponentStyle Merge(ComponentStyle existing, ComponentStyle overlay) => MergeObjects([existing, overlay])!;

    private static ComponentStyle Clone(ComponentStyle style) => MergeObjects([style])!;

    private static void ReportUnknown(DesignProfile profile, List<string> warnings)
    {
        void Report(string path, Dictionary<string, JsonElement>? unknown)
        {
            if (unknown is null)
                return;
            foreach (var key in unknown.Keys)
                warnings.Add($"{path}{key}: unsupported property ignored (profile '{profile.Name}').");
        }
        Report("", profile.Unknown);
        Report("grid.", profile.Grid?.Unknown);
        Report("table.", profile.Table?.Unknown);
        Report("chart.", profile.Chart?.Unknown);
        Report("footer.", profile.Footer?.Unknown);
        Report("source.", profile.Source?.Unknown);
        foreach (var (name, component) in profile.Components ?? [])
            Report($"components.{name}.", component.Unknown);
    }

    private static string Fmt(float value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    [GeneratedRegex("^[A-Za-z0-9_-]{1,64}$", RegexOptions.CultureInvariant)]
    private static partial Regex ProfileName();

    [GeneratedRegex("^[a-z0-9-]{1,40}$", RegexOptions.CultureInvariant)]
    private static partial Regex ComponentName();
}
