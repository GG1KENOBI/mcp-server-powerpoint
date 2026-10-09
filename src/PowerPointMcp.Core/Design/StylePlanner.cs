using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>A formatting run as read from the deck (sizes in points, colors #RRGGBB).</summary>
public sealed record StyledRun(
    int SlideIndex,
    int SlideId,
    int ShapeId,
    string ShapeName,
    string? Role,
    int? Row,
    int? Column,
    int Start,
    int Length,
    string FontName,
    float Size,
    string? Color,
    bool ColorFromTheme,
    string Text);

/// <summary>A shape fill or outline color as read from the deck.</summary>
public sealed record StyledFill(int SlideIndex, int SlideId, int ShapeId, string ShapeName, string Part, string Color, bool FromTheme, string? Component);

/// <summary>A slide's profile tag (name@version) and explicit solid background, when it has one.</summary>
public sealed record StyledSlide(int SlideIndex, int SlideId, string? ProfileTag, string? Background);

/// <summary>One planned or applied formatting change.</summary>
public sealed class DesignChange
{
    /// <summary>1-based slide position (null for theme changes).</summary>
    public int? SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID.</summary>
    public int? SlideId { get; init; }

    /// <summary>Shape.Id.</summary>
    public int? ShapeId { get; init; }

    /// <summary>Shape name.</summary>
    public string? ShapeName { get; init; }

    /// <summary>Table row, for cell text.</summary>
    public int? Row { get; init; }

    /// <summary>Table column, for cell text.</summary>
    public int? Column { get; init; }

    /// <summary>1-based first character, for text changes.</summary>
    public int? Start { get; init; }

    /// <summary>Character count, for text changes.</summary>
    public int? Length { get; init; }

    /// <summary>font-name, font-size, text-color, fill, line, background, table-style, chart-style, theme-color, theme-font, component-version, or profile-tag.</summary>
    public required string Property { get; init; }

    /// <summary>Current value.</summary>
    public string? From { get; init; }

    /// <summary>New value.</summary>
    public required string To { get; init; }

    /// <summary>Why (token name or rule).</summary>
    public string? Reason { get; init; }

    /// <summary>Short text sample of what changes.</summary>
    public string? Sample { get; init; }
}

/// <summary>Which typography rules to apply.</summary>
public sealed record TypographyOptions(bool Fonts = true, bool Sizes = true, bool MinSize = true, float Tolerance = 0.25f);

/// <summary>
/// Plans profile-driven formatting changes from what was read from the deck: typography
/// normalization (font families by role, sizes snapped to the type scale, minimum size) and
/// color migration between profiles (each color that equals a source token becomes the target
/// profile's value for the same token). Theme-bound colors are left to a theme update. Pure.
/// </summary>
public static class StylePlanner
{
    private static readonly HashSet<string> SymbolFonts = new(StringComparer.OrdinalIgnoreCase)
    {
        "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings", "Marlett", "Segoe UI Symbol", "Segoe MDL2 Assets",
        "Segoe Fluent Icons", "Segoe UI Emoji", "MT Extra", "Font Awesome 5 Free", "Material Icons",
    };

    private static readonly HashSet<string> MonoFonts = new(StringComparer.OrdinalIgnoreCase)
    {
        "Consolas", "Courier New", "Cascadia Code", "Cascadia Mono", "Lucida Console", "Courier", "Menlo", "Monaco", "Source Code Pro",
    };

    /// <summary>Roles whose text uses the heading font and the title/subtitle sizes.</summary>
    public static bool IsHeadingRole(string? role) =>
        role is not null && (role is "title" or "subtitle" or "section-number" || role.EndsWith("heading", StringComparison.Ordinal));

    /// <summary>Font sizes the profile defines (type scale plus table, chart, footer, and source sizes), ascending.</summary>
    public static IReadOnlyList<float> AllowedSizes(ResolvedProfile profile)
    {
        ArgumentNullException.ThrowIfNull(profile);
        return profile.TypeScale.Values
            .Append(profile.Table.FontSize).Append(profile.Chart.FontSize).Append(profile.Footer.FontSize).Append(profile.Source.FontSize)
            .Select(size => MathF.Round(size * 2f) / 2f).Distinct().Order().ToList();
    }

    /// <summary>Plans typography normalization; returns changes and notes (off-scale sizes, raised sizes).</summary>
    public static (List<DesignChange> Changes, List<string> Notes) PlanTypography(IEnumerable<StyledRun> runs, ResolvedProfile profile, TypographyOptions options)
    {
        ArgumentNullException.ThrowIfNull(runs);
        ArgumentNullException.ThrowIfNull(profile);
        ArgumentNullException.ThrowIfNull(options);
        var changes = new List<DesignChange>();
        var notes = new List<string>();
        var allowed = AllowedSizes(profile);
        var offScale = new SortedSet<float>();
        int raised = 0;
        foreach (var run in runs)
        {
            if (run.Length == 0 || string.IsNullOrWhiteSpace(run.Text))
                continue;
            if (options.Fonts && TargetFont(run, profile) is { } font && !string.Equals(font, run.FontName, StringComparison.OrdinalIgnoreCase))
                changes.Add(Change(run, "font-name", run.FontName, font, IsHeadingRole(run.Role) ? "fonts.heading" : MonoFonts.Contains(run.FontName) ? "fonts.mono" : "fonts.body"));

            if (run.Size <= 0)
                continue;
            float size = run.Size;
            string? reason = null;
            if (options.Sizes)
            {
                if (run.Role is "title" or "subtitle" && run.Row is null)
                {
                    var roleSize = profile.Size(run.Role);
                    if (roleSize != size)
                        (size, reason) = (roleSize, $"type_scale.{run.Role}");
                }
                else if (Nearest(allowed, size) is { } nearest && nearest != size)
                {
                    if (MathF.Abs(nearest - size) / size <= options.Tolerance)
                        (size, reason) = (nearest, "type scale");
                    else
                        offScale.Add(size);
                }
            }
            if (options.MinSize && size < profile.MinFontSize)
            {
                (size, reason) = (profile.MinFontSize, "min_font_size");
                raised++;
            }
            if (size != run.Size)
                changes.Add(Change(run, "font-size", Points(run.Size), Points(size), reason));
        }
        if (offScale.Count > 0)
            notes.Add($"Sizes far from the type scale were left as they are: {string.Join(", ", offScale.Select(Points))} pt (scale: {string.Join(", ", allowed.Select(Points))}).");
        if (raised > 0)
            notes.Add($"{raised} run(s) are raised to min_font_size {Points(profile.MinFontSize)} pt; run review validate afterwards to catch overflow.");
        return (changes, notes);
    }

    /// <summary>
    /// Builds the color mapping from a source profile to a target profile: color tokens first,
    /// then table, chart palette, footer, source, and component colors. The first token to claim a
    /// source color wins; colors that stay the same are claimed but not changed.
    /// </summary>
    public static IReadOnlyDictionary<string, (string To, string Token)> ColorMap(ResolvedProfile source, ResolvedProfile target)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(target);
        var map = new Dictionary<string, (string To, string Token)>(StringComparer.OrdinalIgnoreCase);
        void Claim(string from, string to, string token)
        {
            from = from.ToUpperInvariant();
            map.TryAdd(from, (to.ToUpperInvariant(), token));
        }
        foreach (var (token, hex) in source.Colors.OrderBy(pair => TokenRank(pair.Key)).ThenBy(pair => pair.Key, StringComparer.Ordinal))
        {
            if (target.Colors.TryGetValue(token, out var to))
                Claim(hex, to, $"colors.{token}");
        }
        Claim(source.Table.HeaderFill, target.Table.HeaderFill, "table.header_fill");
        Claim(source.Table.HeaderText, target.Table.HeaderText, "table.header_text");
        if (source.Table.BandFill is not null && target.Table.BandFill is not null)
            Claim(source.Table.BandFill, target.Table.BandFill, "table.band_fill");
        Claim(source.Table.BodyFill, target.Table.BodyFill, "table.body_fill");
        Claim(source.Table.BodyText, target.Table.BodyText, "table.body_text");
        Claim(source.Table.Border, target.Table.Border, "table.border");
        Claim(source.Table.TotalFill, target.Table.TotalFill, "table.total_fill");
        for (int index = 0; index < Math.Min(source.Chart.Palette.Count, target.Chart.Palette.Count); index++)
            Claim(source.Chart.Palette[index], target.Chart.Palette[index], $"chart.palette[{index.ToString(CultureInfo.InvariantCulture)}]");
        Claim(source.Footer.Color, target.Footer.Color, "footer.color");
        Claim(source.Source.Color, target.Source.Color, "source.color");
        foreach (var (name, component) in source.Components.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            if (!target.Components.TryGetValue(name, out var other))
                continue;
            Claim(component.Fill, other.Fill, $"components.{name}.fill");
            if (component.Border is not null && other.Border is not null)
                Claim(component.Border, other.Border, $"components.{name}.border");
            Claim(component.Accent, other.Accent, $"components.{name}.accent");
            Claim(component.Text, other.Text, $"components.{name}.text");
        }
        return map;
    }

    /// <summary>
    /// Plans color migration. <paramref name="mapFor"/> returns the color map for a slide (by its
    /// source profile). Theme-bound colors are skipped when <paramref name="themeUpdated"/> is true
    /// because they follow the updated theme. Unmapped explicit colors are counted in the notes.
    /// </summary>
    public static (List<DesignChange> Changes, List<string> Notes) PlanColors(
        IEnumerable<StyledRun> runs,
        IEnumerable<StyledFill> fills,
        IEnumerable<StyledSlide> slides,
        Func<int, IReadOnlyDictionary<string, (string To, string Token)>> mapFor,
        bool themeUpdated)
    {
        ArgumentNullException.ThrowIfNull(runs);
        ArgumentNullException.ThrowIfNull(fills);
        ArgumentNullException.ThrowIfNull(slides);
        ArgumentNullException.ThrowIfNull(mapFor);
        var changes = new List<DesignChange>();
        var unmapped = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int themeBound = 0;

        bool TryMap(int slideId, string? color, bool fromTheme, out (string To, string Token) target)
        {
            target = default;
            if (color is null)
                return false;
            if (fromTheme && themeUpdated)
            {
                themeBound++;
                return false;
            }
            if (mapFor(slideId).TryGetValue(color, out target))
                return !string.Equals(target.To, color, StringComparison.OrdinalIgnoreCase);
            unmapped[color.ToUpperInvariant()] = unmapped.GetValueOrDefault(color.ToUpperInvariant()) + 1;
            return false;
        }

        foreach (var run in runs)
        {
            if (run.Length > 0 && TryMap(run.SlideId, run.Color, run.ColorFromTheme, out var target))
                changes.Add(Change(run, "text-color", run.Color!.ToUpperInvariant(), target.To, target.Token));
        }
        foreach (var fill in fills)
        {
            if (TryMap(fill.SlideId, fill.Color, fill.FromTheme, out var target))
            {
                changes.Add(new DesignChange
                {
                    SlideIndex = fill.SlideIndex,
                    SlideId = fill.SlideId,
                    ShapeId = fill.ShapeId,
                    ShapeName = fill.ShapeName,
                    Property = fill.Part,
                    From = fill.Color.ToUpperInvariant(),
                    To = target.To,
                    Reason = target.Token,
                });
            }
        }
        foreach (var slide in slides)
        {
            if (slide.Background is not null && TryMap(slide.SlideId, slide.Background, false, out var target))
                changes.Add(new DesignChange { SlideIndex = slide.SlideIndex, SlideId = slide.SlideId, Property = "background", From = slide.Background.ToUpperInvariant(), To = target.To, Reason = target.Token });
        }

        var notes = new List<string>();
        if (themeBound > 0)
            notes.Add($"{themeBound} theme-colored item(s) follow the updated theme instead of being recolored.");
        if (unmapped.Count > 0)
            notes.Add("Colors that match no source profile token were kept: " + string.Join(", ", unmapped.OrderByDescending(pair => pair.Value).Take(12).Select(pair => $"{pair.Key} ×{pair.Value.ToString(CultureInfo.InvariantCulture)}")) + (unmapped.Count > 12 ? ", …" : "") + ". Pass from_profile if the deck was built with another profile.");
        return (changes, notes);
    }

    private static readonly string[] SizeRoleOrder = ["title", "subtitle", "heading", "body", "quote", "kpi_value", "kpi_label", "caption", "footnote"];

    /// <summary>
    /// Plans font and size migration between profiles: a run in the source profile's heading,
    /// body, or mono font gets the target's font for the same role, and a run whose size equals a
    /// source type-scale size gets the target's size for that role (titles and subtitles prefer
    /// their own role). With <paramref name="themeUpdated"/>, runs in the current theme fonts are
    /// left to follow the theme.
    /// </summary>
    public static (List<DesignChange> Changes, List<string> Notes) PlanFontAndSizeMigration(
        IEnumerable<StyledRun> runs,
        Func<int, ResolvedProfile> sourceFor,
        ResolvedProfile target,
        bool fonts,
        bool sizes,
        IReadOnlyCollection<string> themeFonts,
        bool themeUpdated)
    {
        ArgumentNullException.ThrowIfNull(runs);
        ArgumentNullException.ThrowIfNull(sourceFor);
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(themeFonts);
        var changes = new List<DesignChange>();
        var unmappedFonts = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        int followTheme = 0;
        foreach (var run in runs)
        {
            if (run.Length == 0 || string.IsNullOrWhiteSpace(run.Text))
                continue;
            var source = sourceFor(run.SlideId);
            if (fonts && !SymbolFonts.Contains(run.FontName))
            {
                if (themeUpdated && themeFonts.Contains(run.FontName, StringComparer.OrdinalIgnoreCase))
                {
                    followTheme++;
                }
                else
                {
                    string? token = string.Equals(run.FontName, source.HeadingFont, StringComparison.OrdinalIgnoreCase) && (IsHeadingRole(run.Role) || !string.Equals(source.HeadingFont, source.BodyFont, StringComparison.OrdinalIgnoreCase))
                        ? "heading"
                        : string.Equals(run.FontName, source.BodyFont, StringComparison.OrdinalIgnoreCase) ? "body"
                        : string.Equals(run.FontName, source.Fonts.GetValueOrDefault("mono"), StringComparison.OrdinalIgnoreCase) ? "mono"
                        : null;
                    if (token is null)
                        unmappedFonts.Add(run.FontName);
                    else if (target.Fonts.GetValueOrDefault(token) is { } font && !string.Equals(font, run.FontName, StringComparison.OrdinalIgnoreCase))
                        changes.Add(Change(run, "font-name", run.FontName, font, $"fonts.{token}"));
                }
            }
            if (sizes && run.Size > 0)
            {
                string? role = run.Role is "title" or "subtitle" && source.TypeScale.GetValueOrDefault(run.Role) == run.Size
                    ? run.Role
                    : SizeRoleOrder.FirstOrDefault(candidate => source.TypeScale.GetValueOrDefault(candidate) == run.Size);
                if (role is not null && target.TypeScale.TryGetValue(role, out var size) && size != run.Size)
                    changes.Add(Change(run, "font-size", Points(run.Size), Points(size), $"type_scale.{role}"));
            }
        }
        var notes = new List<string>();
        if (followTheme > 0)
            notes.Add($"{followTheme} run(s) use the theme fonts and follow the updated theme.");
        if (unmappedFonts.Count > 0)
            notes.Add($"Fonts that match no source profile font were kept: {string.Join(", ", unmappedFonts)}. Use design normalize-typography to unify them.");
        return (changes, notes);
    }

    private static string? TargetFont(StyledRun run, ResolvedProfile profile)
    {
        if (SymbolFonts.Contains(run.FontName) || run.FontName.StartsWith('+'))
            return null;
        if (MonoFonts.Contains(run.FontName) || string.Equals(run.FontName, profile.Fonts.GetValueOrDefault("mono"), StringComparison.OrdinalIgnoreCase))
            return profile.Fonts.GetValueOrDefault("mono") ?? run.FontName;
        return IsHeadingRole(run.Role) && run.Row is null ? profile.HeadingFont : profile.BodyFont;
    }

    private static float? Nearest(IReadOnlyList<float> allowed, float size)
    {
        float? best = null;
        foreach (var candidate in allowed)
        {
            if (best is null || MathF.Abs(candidate - size) < MathF.Abs(best.Value - size) || (MathF.Abs(candidate - size) == MathF.Abs(best.Value - size) && candidate > best))
                best = candidate;
        }
        return best;
    }

    private static int TokenRank(string token) => token switch
    {
        "primary" => 0,
        "secondary" => 1,
        "accent" => 2,
        "text" => 3,
        "background" => 4,
        "surface" => 5,
        "muted" => 6,
        "border" => 7,
        "positive" => 8,
        "negative" => 9,
        "neutral" => 10,
        _ => 20,
    };

    private static DesignChange Change(StyledRun run, string property, string from, string to, string? reason) => new()
    {
        SlideIndex = run.SlideIndex,
        SlideId = run.SlideId,
        ShapeId = run.ShapeId,
        ShapeName = run.ShapeName,
        Row = run.Row,
        Column = run.Column,
        Start = run.Start,
        Length = run.Length,
        Property = property,
        From = from,
        To = to,
        Reason = reason,
        Sample = run.Text.Length > 40 ? run.Text[..40] + "…" : run.Text,
    };

    /// <summary>Formats points without trailing zeros.</summary>
    public static string Points(float value) => value.ToString("0.#", CultureInfo.InvariantCulture);
}
