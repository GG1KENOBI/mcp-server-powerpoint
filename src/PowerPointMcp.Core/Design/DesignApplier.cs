extern alias OfficeInterop;

using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>Writes planned design changes to the deck. STA thread only.</summary>
internal static class DesignApplier
{
    /// <summary>Theme color roles a profile writes, as (Office role, profile token).</summary>
    internal static readonly (Office.MsoThemeColorSchemeIndex Role, string Token)[] ThemeRoles =
    [
        (Office.MsoThemeColorSchemeIndex.msoThemeDark1, "text"),
        (Office.MsoThemeColorSchemeIndex.msoThemeLight1, "background"),
        (Office.MsoThemeColorSchemeIndex.msoThemeDark2, "muted"),
        (Office.MsoThemeColorSchemeIndex.msoThemeLight2, "surface"),
        (Office.MsoThemeColorSchemeIndex.msoThemeAccent1, "primary"),
        (Office.MsoThemeColorSchemeIndex.msoThemeAccent2, "accent"),
        (Office.MsoThemeColorSchemeIndex.msoThemeAccent3, "neutral"),
        (Office.MsoThemeColorSchemeIndex.msoThemeAccent5, "secondary"),
        (Office.MsoThemeColorSchemeIndex.msoThemeAccent6, "positive"),
    ];

    /// <summary>Plans the theme changes of every master (colors and Latin fonts) without writing.</summary>
    internal static List<DesignChange> PlanTheme(PowerPoint.Presentation presentation, ResolvedProfile profile)
    {
        var changes = new List<DesignChange>();
        ForEachTheme(presentation, (masterIndex, scheme, fonts) =>
        {
            foreach (var (role, token) in ThemeRoles)
            {
                Office.ThemeColor? color = null;
                try
                {
                    color = scheme.Colors(role);
                    var current = DeckSnapshotReader.Hex(color.RGB);
                    var target = profile.Color(token);
                    if (!string.Equals(current, target, StringComparison.OrdinalIgnoreCase))
                        changes.Add(new DesignChange { Property = "theme-color", From = current, To = target, Reason = $"master {masterIndex.ToString(CultureInfo.InvariantCulture)} {role.ToString()["msoTheme".Length..]} = colors.{token}" });
                }
                finally
                {
                    if (color is not null) ComUtilities.Release(ref color);
                }
            }
            foreach (var (major, token) in new[] { (true, "heading"), (false, "body") })
            {
                var current = ReadThemeFont(fonts, major);
                var target = profile.Fonts[token];
                if (!string.Equals(current, target, StringComparison.OrdinalIgnoreCase))
                    changes.Add(new DesignChange { Property = "theme-font", From = current, To = target, Reason = $"master {masterIndex.ToString(CultureInfo.InvariantCulture)} {(major ? "heading" : "body")} = fonts.{token}" });
            }
        });
        return changes;
    }

    /// <summary>Writes the profile colors and fonts into every master's theme.</summary>
    internal static void WriteTheme(PowerPoint.Presentation presentation, ResolvedProfile profile)
    {
        ForEachTheme(presentation, (_, scheme, fonts) =>
        {
            foreach (var (role, token) in ThemeRoles)
            {
                Office.ThemeColor? color = null;
                try
                {
                    color = scheme.Colors(role);
                    color.RGB = ChartStyler.Bgr(profile.Color(token));
                }
                finally
                {
                    if (color is not null) ComUtilities.Release(ref color);
                }
            }
            WriteThemeFont(fonts, true, profile.HeadingFont);
            WriteThemeFont(fonts, false, profile.BodyFont);
        });
    }

    /// <summary>Reads the first master's theme colors (role name to #RRGGBB) and Latin heading/body fonts.</summary>
    internal static (Dictionary<string, string> Colors, string? Heading, string? Body) ReadTheme(PowerPoint.Presentation presentation)
    {
        var colors = new Dictionary<string, string>(StringComparer.Ordinal);
        string? heading = null, body = null;
        bool first = true;
        ForEachTheme(presentation, (_, scheme, fonts) =>
        {
            if (!first)
                return;
            first = false;
            foreach (var role in Enum.GetValues<Office.MsoThemeColorSchemeIndex>())
            {
                Office.ThemeColor? color = null;
                try
                {
                    color = scheme.Colors(role);
                    colors[role.ToString()["msoTheme".Length..]] = DeckSnapshotReader.Hex(color.RGB);
                }
                finally
                {
                    if (color is not null) ComUtilities.Release(ref color);
                }
            }
            heading = ReadThemeFont(fonts, true);
            body = ReadThemeFont(fonts, false);
        });
        return (colors, heading, body);
    }

    private static void ForEachTheme(PowerPoint.Presentation presentation, Action<int, Office.ThemeColorScheme, Office.ThemeFontScheme> action)
    {
        PowerPoint.Designs? designs = null;
        try
        {
            designs = presentation.Designs;
            for (int index = 1; index <= designs.Count; index++)
            {
                PowerPoint.Design? design = null;
                PowerPoint.Master? master = null;
                Office.OfficeTheme? theme = null;
                Office.ThemeColorScheme? colors = null;
                Office.ThemeFontScheme? fonts = null;
                try
                {
                    design = designs[index];
                    master = design.SlideMaster;
                    theme = master.Theme;
                    colors = theme.ThemeColorScheme;
                    fonts = theme.ThemeFontScheme;
                    action(index, colors, fonts);
                }
                finally
                {
                    if (fonts is not null) ComUtilities.Release(ref fonts);
                    if (colors is not null) ComUtilities.Release(ref colors);
                    if (theme is not null) ComUtilities.Release(ref theme);
                    if (master is not null) ComUtilities.Release(ref master);
                    if (design is not null) ComUtilities.Release(ref design);
                }
            }
        }
        finally
        {
            if (designs is not null) ComUtilities.Release(ref designs);
        }
    }

    private static string? ReadThemeFont(Office.ThemeFontScheme scheme, bool major)
    {
        Office.ThemeFonts? fonts = null;
        Office.ThemeFont? font = null;
        try
        {
            fonts = major ? scheme.MajorFont : scheme.MinorFont;
            font = fonts.Item(Office.MsoFontLanguageIndex.msoThemeLatin);
            return font.Name;
        }
        finally
        {
            if (font is not null) ComUtilities.Release(ref font);
            if (fonts is not null) ComUtilities.Release(ref fonts);
        }
    }

    private static void WriteThemeFont(Office.ThemeFontScheme scheme, bool major, string name)
    {
        Office.ThemeFonts? fonts = null;
        Office.ThemeFont? font = null;
        try
        {
            fonts = major ? scheme.MajorFont : scheme.MinorFont;
            font = fonts.Item(Office.MsoFontLanguageIndex.msoThemeLatin);
            font.Name = name;
        }
        finally
        {
            if (font is not null) ComUtilities.Release(ref font);
            if (fonts is not null) ComUtilities.Release(ref fonts);
        }
    }

    /// <summary>
    /// Applies run, fill, line, and background changes. Returns how many were written and notes
    /// for objects that disappeared or refused a change.
    /// </summary>
    internal static (int Applied, List<string> Notes) ApplyChanges(PowerPoint.Presentation presentation, IReadOnlyList<DesignChange> changes, CancellationToken cancellationToken)
    {
        int applied = 0;
        var notes = new List<string>();
        foreach (var bySlide in changes.Where(change => change.SlideId is not null).GroupBy(change => change.SlideId!.Value))
        {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPoint.Slide? slide = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(presentation, bySlide.Key);
                if (slide is null)
                {
                    notes.Add($"Slide {bySlide.Key} disappeared; its changes were skipped.");
                    continue;
                }
                foreach (var background in bySlide.Where(change => change.Property == "background"))
                {
                    SetBackground(slide, background.To);
                    applied++;
                }
                foreach (var byShape in bySlide.Where(change => change.ShapeId is not null).GroupBy(change => change.ShapeId!.Value))
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = DeckShapeLocator.FindShape(slide, byShape.Key);
                        if (shape is null)
                        {
                            notes.Add($"Shape {byShape.Key} on slide {bySlide.Key} disappeared; its changes were skipped.");
                            continue;
                        }
                        foreach (var change in byShape)
                        {
                            try
                            {
                                ApplyToShape(shape, change);
                                applied++;
                            }
                            catch (COMException ex)
                            {
                                notes.Add($"{change.Property} on '{change.ShapeName}' (slide {change.SlideIndex}) was not changed: {ex.Message}");
                            }
                        }
                    }
                    finally
                    {
                        if (shape is not null) ComUtilities.Release(ref shape);
                    }
                }
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        }
        return (applied, notes);
    }

    private static void ApplyToShape(PowerPoint.Shape shape, DesignChange change)
    {
        switch (change.Property)
        {
            case "fill":
                WithColor(shape, fill: true, color => color.RGB = ChartStyler.Bgr(change.To));
                return;
            case "line":
                WithColor(shape, fill: false, color => color.RGB = ChartStyler.Bgr(change.To));
                return;
            case "font-name":
            case "font-size":
            case "text-color":
                WithRun(shape, change, font =>
                {
                    if (change.Property == "font-name")
                    {
                        font.Name = change.To;
                    }
                    else if (change.Property == "font-size")
                    {
                        font.Size = float.Parse(change.To, CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        PowerPoint.ColorFormat? color = null;
                        try
                        {
                            color = font.Color;
                            color.RGB = ChartStyler.Bgr(change.To);
                        }
                        finally
                        {
                            if (color is not null) ComUtilities.Release(ref color);
                        }
                    }
                });
                return;
            default:
                throw new InvalidOperationException($"Unknown design change '{change.Property}'.");
        }
    }

    private static void WithColor(PowerPoint.Shape shape, bool fill, Action<PowerPoint.ColorFormat> write)
    {
        PowerPoint.FillFormat? fillFormat = null;
        PowerPoint.LineFormat? line = null;
        PowerPoint.ColorFormat? color = null;
        try
        {
            if (fill)
            {
                fillFormat = shape.Fill;
                color = fillFormat.ForeColor;
            }
            else
            {
                line = shape.Line;
                color = line.ForeColor;
            }
            write(color);
        }
        finally
        {
            if (color is not null) ComUtilities.Release(ref color);
            if (line is not null) ComUtilities.Release(ref line);
            if (fillFormat is not null) ComUtilities.Release(ref fillFormat);
        }
    }

    private static void WithRun(PowerPoint.Shape shape, DesignChange change, Action<PowerPoint.Font> write)
    {
        PowerPoint.Table? table = null;
        PowerPoint.Cell? cell = null;
        PowerPoint.Shape? cellShape = null;
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        PowerPoint.TextRange? run = null;
        PowerPoint.Font? font = null;
        try
        {
            if (change.Row is { } row && change.Column is { } column)
            {
                table = shape.Table;
                cell = table.Cell(row, column);
                cellShape = cell.Shape;
                frame = cellShape.TextFrame;
            }
            else
            {
                frame = shape.TextFrame;
            }
            range = frame.TextRange;
            run = range.Characters(change.Start ?? 1, change.Length ?? range.Length);
            font = run.Font;
            write(font);
        }
        finally
        {
            if (font is not null) ComUtilities.Release(ref font);
            if (run is not null) ComUtilities.Release(ref run);
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
            if (cellShape is not null) ComUtilities.Release(ref cellShape);
            if (cell is not null) ComUtilities.Release(ref cell);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static void SetBackground(PowerPoint.Slide slide, string hex)
    {
        PowerPoint.ShapeRange? background = null;
        PowerPoint.FillFormat? fill = null;
        PowerPoint.ColorFormat? color = null;
        try
        {
            slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
            background = slide.Background;
            fill = background.Fill;
            fill.Solid();
            color = fill.ForeColor;
            color.RGB = ChartStyler.Bgr(hex);
        }
        finally
        {
            if (color is not null) ComUtilities.Release(ref color);
            if (fill is not null) ComUtilities.Release(ref fill);
            if (background is not null) ComUtilities.Release(ref background);
        }
    }

    /// <summary>Restyles tables and charts with the profile; returns notes for what failed.</summary>
    internal static List<string> RestyleObjects(PowerPoint.Presentation presentation, IEnumerable<ScannedObject> objects, ResolvedProfile profile, CancellationToken cancellationToken)
    {
        var notes = new List<string>();
        foreach (var item in objects.Where(item => item.Kind is "table" or "chart"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPoint.Slide? slide = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.Table? table = null;
            PowerPoint.Chart? chart = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(presentation, item.SlideId);
                shape = slide is null ? null : DeckShapeLocator.FindShape(slide, item.ShapeId);
                if (shape is null)
                {
                    notes.Add($"{item.Kind} '{item.ShapeName}' disappeared before it could be restyled.");
                    continue;
                }
                if (item.Kind == "table")
                {
                    table = shape.Table;
                    TableStyler.StyleExisting(table, profile.Table, profile.BodyFont, profile.Table.FontSize, item.HeaderRow, item.TotalRow, item.Banded, null);
                }
                else
                {
                    chart = shape.Chart;
                    var chartNotes = ChartStyler.Apply(chart, new ChartLook(profile.Chart.Palette, profile.BodyFont, profile.Chart.FontSize, profile.Color("text"),
                        profile.Chart.Gridlines, profile.Chart.DataLabels, profile.Chart.Legend, null, null, null), item.IsPie);
                    notes.AddRange(chartNotes.Select(note => $"Chart '{item.ShapeName}' (slide {item.SlideIndex.ToString(CultureInfo.InvariantCulture)}): {note}"));
                }
            }
            catch (COMException ex)
            {
                notes.Add($"{item.Kind} '{item.ShapeName}' (slide {item.SlideIndex.ToString(CultureInfo.InvariantCulture)}) was not restyled: {ex.Message}");
            }
            finally
            {
                if (chart is not null) ComUtilities.Release(ref chart);
                if (table is not null) ComUtilities.Release(ref table);
                if (shape is not null) ComUtilities.Release(ref shape);
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        }
        return notes;
    }

    /// <summary>Sets a tag on shapes (by SlideID and Shape.Id) or, with no shape id, on slides.</summary>
    internal static void SetTags(PowerPoint.Presentation presentation, IEnumerable<(int SlideId, int? ShapeId, string Name, string Value)> tags)
    {
        foreach (var bySlide in tags.GroupBy(tag => tag.SlideId))
        {
            PowerPoint.Slide? slide = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(presentation, bySlide.Key);
                if (slide is null)
                    continue;
                foreach (var (_, shapeId, name, value) in bySlide)
                {
                    PowerPoint.Shape? shape = null;
                    PowerPoint.Tags? collection = null;
                    try
                    {
                        if (shapeId is { } id)
                        {
                            shape = DeckShapeLocator.FindShape(slide, id);
                            if (shape is null)
                                continue;
                            collection = shape.Tags;
                        }
                        else
                        {
                            collection = slide.Tags;
                        }
                        collection.Add(name, value);
                    }
                    finally
                    {
                        if (collection is not null) ComUtilities.Release(ref collection);
                        if (shape is not null) ComUtilities.Release(ref shape);
                    }
                }
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        }
    }
}
