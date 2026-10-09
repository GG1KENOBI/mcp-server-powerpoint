using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>What a deck uses, gathered through COM for <see cref="ProfileExtractor"/>.</summary>
public sealed record DeckStyleSample(
    IReadOnlyDictionary<string, string>? ThemeColors,
    string? HeadingFont,
    string? BodyFont,
    float SlideWidth,
    float SlideHeight,
    IReadOnlyList<StyledRun> Runs,
    IReadOnlyList<StyledFill> Fills,
    IReadOnlyList<(float Left, float Top, float Width, float Height)> TitleBoxes,
    IReadOnlyList<(float Left, float Top, float Width, float Height)> ContentBoxes);

/// <summary>
/// Drafts a design profile from an existing deck: theme colors and fonts, the most used title,
/// body, and caption sizes (weighted by characters), the dominant explicit text color, the most
/// used explicit fill as surface, and margins from where titles and content sit. Values that
/// cannot be inferred are inherited from the default profile; the notes say what was inferred
/// from what. Pure.
/// </summary>
public static class ProfileExtractor
{
    /// <summary>Builds the draft profile and notes.</summary>
    public static (DesignProfile Profile, IReadOnlyList<string> Notes) Extract(DeckStyleSample sample, string name)
    {
        ArgumentNullException.ThrowIfNull(sample);
        var (profile, themeNotes) = ThemeProfile.FromTheme(sample.ThemeColors, sample.HeadingFont, sample.BodyFont);
        var notes = new List<string>(themeNotes);
        profile.Name = name;
        profile.Version = "1.0.0";
        profile.Description = "Extracted from an existing presentation.";
        var colors = profile.Colors ??= [];

        var textRuns = sample.Runs.Where(run => run.Length > 0 && run.Size > 0 && !string.IsNullOrWhiteSpace(run.Text)).ToList();
        var typeScale = new Dictionary<string, float>(StringComparer.Ordinal);
        var titleSize = Mode(textRuns.Where(run => run.Role == "title"));
        if (titleSize is { } title)
        {
            typeScale["title"] = title;
            notes.Add($"type_scale.title {StylePlanner.Points(title)} pt: most used title size.");
        }
        var subtitleSize = Mode(textRuns.Where(run => run.Role == "subtitle"));
        if (subtitleSize is { } subtitle)
            typeScale["subtitle"] = subtitle;
        var bodyRuns = textRuns.Where(run => !StylePlanner.IsHeadingRole(run.Role) && run.Role is not ("footer" or "slide-number" or "date") && run.Row is null).ToList();
        var bodySize = Mode(bodyRuns);
        if (bodySize is { } body)
        {
            typeScale["body"] = body;
            notes.Add($"type_scale.body {StylePlanner.Points(body)} pt: most used body text size (by characters).");
            var smaller = Mode(bodyRuns.Where(run => run.Size < body));
            if (smaller is { } caption)
            {
                typeScale["caption"] = caption;
                typeScale["footnote"] = MathF.Min(caption, MathF.Max(8, caption - 2));
            }
            var larger = Mode(bodyRuns.Where(run => run.Size > body && (titleSize is null || run.Size < titleSize)));
            if (larger is { } heading)
                typeScale["heading"] = heading;
            var minimum = bodyRuns.Min(run => run.Size);
            profile.MinFontSize = MathF.Min(12, MathF.Max(8, minimum));
        }
        var tableSize = Mode(textRuns.Where(run => run.Row is not null));
        if (tableSize is { } tableFont)
            profile.Table = new TableStyle { FontSize = tableFont };
        if (typeScale.Count > 0)
            profile.TypeScale = typeScale;

        var explicitText = textRuns.Where(run => run.Color is not null && !run.ColorFromTheme)
            .GroupBy(run => run.Color!.ToUpperInvariant())
            .Select(group => (Color: group.Key, Characters: group.Sum(run => run.Length)))
            .OrderByDescending(entry => entry.Characters).ToList();
        if (explicitText.Count > 0 && explicitText[0].Characters * 2 >= explicitText.Sum(entry => entry.Characters))
        {
            colors["text"] = explicitText[0].Color;
            notes.Add($"colors.text {explicitText[0].Color}: dominant explicit text color.");
        }
        var explicitFills = sample.Fills.Where(fill => fill.Part == "fill" && !fill.FromTheme)
            .GroupBy(fill => fill.Color.ToUpperInvariant()).OrderByDescending(group => group.Count()).ToList();
        if (explicitFills.Count > 0 && explicitFills[0].Count() >= 2 && !explicitFills[0].Key.Equals(colors.GetValueOrDefault("background"), StringComparison.OrdinalIgnoreCase))
        {
            colors["surface"] = explicitFills[0].Key;
            notes.Add($"colors.surface {explicitFills[0].Key}: most used explicit shape fill ({explicitFills[0].Count().ToString(CultureInfo.InvariantCulture)} shapes).");
        }

        if (sample.SlideWidth > 0 && (sample.TitleBoxes.Count > 0 || sample.ContentBoxes.Count > 0))
        {
            var boxes = sample.TitleBoxes.Concat(sample.ContentBoxes).Where(box => box.Width > 0 && box.Left >= 0 && box.Left + box.Width <= sample.SlideWidth).ToList();
            if (boxes.Count > 0)
            {
                float left = Percentile(boxes.Select(box => box.Left), 0.2f);
                float right = Percentile(boxes.Select(box => sample.SlideWidth - box.Left - box.Width), 0.2f);
                float top = sample.TitleBoxes.Count > 0 ? Percentile(sample.TitleBoxes.Select(box => box.Top), 0.5f) : Percentile(boxes.Select(box => box.Top), 0.2f);
                profile.Margins = new Dictionary<string, float>(StringComparer.Ordinal)
                {
                    ["left"] = Round(Math.Clamp(left, 0, sample.SlideWidth / 4)),
                    ["right"] = Round(Math.Clamp(right, 0, sample.SlideWidth / 4)),
                    ["top"] = Round(Math.Clamp(top, 0, sample.SlideHeight / 4)),
                };
                notes.Add("margins.left/right/top: where titles and content usually start and end.");
                if (sample.TitleBoxes.Count > 0)
                {
                    float titleHeight = Percentile(sample.TitleBoxes.Select(box => box.Height), 0.5f);
                    profile.Areas = new Dictionary<string, float>(StringComparer.Ordinal) { ["title"] = Round(Math.Clamp(titleHeight, 24, sample.SlideHeight / 3)) };
                }
            }
        }
        return (profile, notes);
    }

    /// <summary>Most used size weighted by characters (ties go to the larger size), or null.</summary>
    private static float? Mode(IEnumerable<StyledRun> runs)
    {
        var best = runs.GroupBy(run => MathF.Round(run.Size * 2f) / 2f)
            .Select(group => (Size: group.Key, Characters: group.Sum(run => run.Length)))
            .OrderByDescending(entry => entry.Characters).ThenByDescending(entry => entry.Size)
            .FirstOrDefault();
        return best.Characters > 0 ? best.Size : null;
    }

    private static float Percentile(IEnumerable<float> values, float fraction)
    {
        var sorted = values.Order().ToList();
        return sorted.Count == 0 ? 0 : sorted[(int)MathF.Floor(fraction * (sorted.Count - 1))];
    }

    private static float Round(float value) => MathF.Round(value);
}
