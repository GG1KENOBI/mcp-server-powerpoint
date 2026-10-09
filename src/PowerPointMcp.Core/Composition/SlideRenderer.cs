extern alias OfficeInterop;

using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>Where a planned element ended up.</summary>
public sealed class RenderedElement
{
    /// <summary>Semantic key.</summary>
    public required string Key { get; init; }

    /// <summary>Semantic role.</summary>
    public required string Role { get; init; }

    /// <summary>Shape.Id.</summary>
    public required int ShapeId { get; init; }

    /// <summary>PPTMCP_ID of the shape.</summary>
    public required string AppId { get; init; }

    /// <summary>Final geometry in points (left, top, width, height), read back from PowerPoint.</summary>
    public required IReadOnlyList<float> Box { get; init; }

    /// <summary>Estimated text height from planning, if any.</summary>
    public float? EstimatedHeight { get; init; }

    /// <summary>Text height measured by PowerPoint (including frame margins), if the element has text.</summary>
    public float? MeasuredHeight { get; set; }

    /// <summary>fits, restacked, shrunk, overflow, or not-measured.</summary>
    public string Fit { get; set; } = "not-measured";

    /// <summary>Smallest font size after fitting.</summary>
    public float? MinFontSize { get; set; }
}

/// <summary>The outcome of rendering one planned slide.</summary>
public sealed class RenderedSlide
{
    /// <summary>PowerPoint SlideID.</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based position.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>Persistent slide id.</summary>
    public required string AppId { get; init; }

    /// <summary>Elements by creation order.</summary>
    public required List<RenderedElement> Elements { get; init; }

    /// <summary>Measured overflow that fitting could not resolve.</summary>
    public required List<string> Overflow { get; init; }

    /// <summary>Notes about what could not be applied.</summary>
    public required List<string> Warnings { get; init; }

    /// <summary>When set, only this many items fit (measured); the composer re-plans with a split.</summary>
    public int? SplitAfter { get; set; }
}

/// <summary>
/// Creates native PowerPoint objects for a planned slide, measures text with PowerPoint, and fits
/// regions (restack, proportional shrink, or a split request). STA thread only.
/// </summary>
internal static class SlideRenderer
{
    private const float Tolerance = 1f;

    internal static RenderedSlide Render(
        PowerPoint.Presentation presentation,
        PlannedSlide planned,
        CompositionPlan plan,
        ResolvedProfile profile,
        int insertIndex,
        LayoutChoice layouts,
        bool applyFooter,
        CancellationToken cancellationToken)
    {
        PowerPoint.Slide? slide = null;
        var shapes = new Dictionary<string, PowerPoint.Shape>(StringComparer.Ordinal);
        var warnings = new List<string>();
        try
        {
            slide = planned.Title is null
                ? LayoutPicker.AddSlide(presentation, insertIndex, layouts.BlankIndex, PowerPoint.PpSlideLayout.ppLayoutBlank)
                : LayoutPicker.AddSlide(presentation, insertIndex, layouts.TitleOnlyIndex, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
            TagSlide(slide, planned, plan, profile);

            foreach (var (key, shape) in CreateElements(slide, planned.Elements, key => AppIdFor(planned, key), warnings, cancellationToken))
                shapes[key] = shape;

            if (planned.Notes is { Length: > 0 } notes)
                WriteNotes(slide, notes, warnings);
            if (applyFooter)
                ApplyFooter(slide, profile, warnings);

            var rendered = new RenderedSlide
            {
                SlideId = slide.SlideID,
                SlideIndex = slide.SlideIndex,
                AppId = planned.AppId,
                Elements = [],
                Overflow = [],
                Warnings = warnings,
            };
            var results = planned.Elements.Where(element => shapes.ContainsKey(element.Key)).ToDictionary(element => element.Key, element => new RenderedElement
            {
                Key = element.Key,
                Role = element.Role,
                ShapeId = shapes[element.Key].Id,
                AppId = AppIdFor(planned, element.Key),
                Box = [],
                EstimatedHeight = element.EstimatedTextHeight,
            });

            foreach (var region in planned.Regions)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Fit(region, planned, shapes, results, plan.Spec.Fit, rendered);
                if (rendered.SplitAfter is not null)
                    return rendered;
            }

            GroupParts(slide, planned, shapes, warnings);
            foreach (var element in planned.Elements.Where(element => shapes.ContainsKey(element.Key)))
            {
                var shape = shapes[element.Key];
                var result = results[element.Key];
                rendered.Elements.Add(new RenderedElement
                {
                    Key = result.Key,
                    Role = result.Role,
                    ShapeId = result.ShapeId,
                    AppId = result.AppId,
                    Box = [Round(shape.Left), Round(shape.Top), Round(shape.Width), Round(shape.Height)],
                    EstimatedHeight = result.EstimatedHeight,
                    MeasuredHeight = result.MeasuredHeight ?? (element.Paragraphs is not null ? Round(TextRuns.NeededHeight(shape) ?? 0) : null),
                    Fit = result.Fit,
                    MinFontSize = result.MinFontSize,
                });
            }
            return rendered;
        }
        finally
        {
            foreach (var shape in shapes.Values)
            {
                var item = shape;
                ComUtilities.Release(ref item!);
            }
            if (slide is not null) ComUtilities.Release(ref slide);
        }
    }

    internal static string AppIdFor(PlannedSlide slide, string key) => $"{slide.AppId}/{key}";

    /// <summary>
    /// Creates and styles the elements in order (connectors after their end shapes) and returns
    /// the shapes by key. The caller owns and releases the returned shapes.
    /// </summary>
    internal static Dictionary<string, PowerPoint.Shape> CreateElements(
        PowerPoint.Slide slide,
        IReadOnlyList<PlannedElement> elements,
        Func<string, string> appIdFor,
        List<string> warnings,
        CancellationToken cancellationToken)
    {
        var shapes = new Dictionary<string, PowerPoint.Shape>(StringComparer.Ordinal);
        bool completed = false;
        try
        {
            foreach (var element in elements)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var shape = Create(slide, element, shapes, warnings);
                if (shape is null)
                    continue;
                shapes[element.Key] = shape;
                Decorate(shape, element, appIdFor(element.Key));
            }
            completed = true;
            return shapes;
        }
        finally
        {
            if (!completed)
            {
                foreach (var shape in shapes.Values)
                {
                    var item = shape;
                    ComUtilities.Release(ref item!);
                }
            }
        }
    }

    /// <summary>Fits text shapes that must each fit their own box (shared proportional shrink); returns overflow notes.</summary>
    internal static List<string> FitIndividually(IReadOnlyList<string> keys, Dictionary<string, PowerPoint.Shape> shapes, float minFontSize, bool canShrink)
    {
        var region = new FitRegion { Key = "diagram", Box = default, Mode = "each", Keys = keys, MinFontSize = minFontSize };
        var results = keys.Where(shapes.ContainsKey).ToDictionary(key => key, key => new RenderedElement { Key = key, Role = "", ShapeId = 0, AppId = "", Box = [] });
        var rendered = new RenderedSlide { SlideId = 0, SlideIndex = 0, AppId = "", Elements = [], Overflow = [], Warnings = [] };
        FitEach(region, keys.Where(shapes.ContainsKey).ToList(), shapes, results, canShrink, rendered);
        return rendered.Overflow;
    }

    private static PowerPoint.Shape? Create(PowerPoint.Slide slide, PlannedElement element, Dictionary<string, PowerPoint.Shape> created, List<string> warnings)
    {
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            var box = element.Box;
            PowerPoint.Shape? shape = element.Type switch
            {
                "text" when element.TitlePlaceholder => TitlePlaceholder(shapes, box, warnings),
                "text" => shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, box.Left, box.Top, Math.Max(1, box.Width), Math.Max(1, box.Height)),
                "shape" => shapes.AddShape(Geometry(element.Geometry), box.Left, box.Top, Math.Max(1, box.Width), Math.Max(1, box.Height)),
                "line" => shapes.AddLine(box.Left, box.Top, box.Right, box.Bottom),
                "table" => CreateTable(shapes, element),
                "chart" => CreateChart(shapes, element, warnings),
                "picture" => CreatePicture(shapes, element),
                "connector" => CreateConnector(shapes, element, created, warnings),
                _ => null,
            };
            if (shape is null)
                return null;

            if (element.Type is "text" or "shape" && element.Paragraphs is { Count: > 0 } paragraphs)
                WriteText(shape, element, paragraphs);
            return shape;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    private static PowerPoint.Shape TitlePlaceholder(PowerPoint.Shapes shapes, Box box, List<string> warnings)
    {
        PowerPoint.Shape title;
        if (shapes.HasTitle == Office.MsoTriState.msoTrue)
        {
            title = shapes.Title;
        }
        else
        {
            // The layout has no title placeholder (custom template): use a text box tagged as the title.
            warnings.Add("The slide layout has no title placeholder; the title is a text box, so screen readers may not announce it as the slide title.");
            title = shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, box.Left, box.Top, box.Width, box.Height);
        }
        title.Left = box.Left;
        title.Top = box.Top;
        title.Width = Math.Max(1, box.Width);
        title.Height = Math.Max(1, box.Height);
        return title;
    }

    private static Office.MsoAutoShapeType Geometry(string geometry) => geometry switch
    {
        "rounded-rectangle" => Office.MsoAutoShapeType.msoShapeRoundedRectangle,
        "oval" => Office.MsoAutoShapeType.msoShapeOval,
        "chevron" => Office.MsoAutoShapeType.msoShapeChevron,
        "pentagon" => Office.MsoAutoShapeType.msoShapePentagon,
        "triangle" => Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
        "diamond" => Office.MsoAutoShapeType.msoShapeDiamond,
        "hexagon" => Office.MsoAutoShapeType.msoShapeHexagon,
        "cylinder" => Office.MsoAutoShapeType.msoShapeCan,
        "parallelogram" => Office.MsoAutoShapeType.msoShapeParallelogram,
        "document" => Office.MsoAutoShapeType.msoShapeFlowchartDocument,
        "terminator" => Office.MsoAutoShapeType.msoShapeFlowchartTerminator,
        _ => Office.MsoAutoShapeType.msoShapeRectangle,
    };

    private static void Decorate(PowerPoint.Shape shape, PlannedElement element, string appId)
    {
        if (!element.TitlePlaceholder)
            shape.Name = element.Key;
        var tags = shape.Tags;
        try
        {
            tags.Add(DeckRoles.IdTag, appId);
            tags.Add(DeckRoles.RoleTag, element.Role);
            if (element.Component is not null)
                tags.Add(DeckRoles.ComponentTag, element.Component);
            foreach (var (name, value) in element.Tags ?? new Dictionary<string, string>())
                tags.Add(name, value);
        }
        finally
        {
            ComUtilities.Release(ref tags);
        }

        if (element.AltText is not null)
            shape.AlternativeText = element.AltText;
        if (element.Type is "shape")
        {
            ApplyFill(shape, element.Fill);
            ApplyLine(shape, element.LineColor, element.LineWidth);
            if (element.Geometry == "rounded-rectangle" && element.Radius > 0)
                SetRadius(shape, element);
        }
        else if (element.Type is "line" or "connector")
        {
            ApplyLine(shape, element.LineColor ?? "#808080", element.LineWidth > 0 ? element.LineWidth : 1f);
            if (element.EndArrow)
                SetEndArrow(shape);
        }
        else if (element.Type == "text" && !element.TitlePlaceholder)
        {
            ApplyFill(shape, element.Fill);
            ApplyLine(shape, element.LineColor, element.LineWidth);
        }
        if (Math.Abs(element.Rotation) > 0.01f)
            shape.Rotation = element.Rotation;
    }

    private static void ApplyFill(PowerPoint.Shape shape, string? color)
    {
        PowerPoint.FillFormat? fill = null;
        PowerPoint.ColorFormat? fore = null;
        try
        {
            fill = shape.Fill;
            if (color is null)
            {
                fill.Visible = Office.MsoTriState.msoFalse;
                return;
            }
            fill.Visible = Office.MsoTriState.msoTrue;
            fill.Solid();
            fore = fill.ForeColor;
            fore.RGB = ChartStyler.Bgr(color);
            fill.Transparency = 0f;
        }
        finally
        {
            if (fore is not null) ComUtilities.Release(ref fore);
            if (fill is not null) ComUtilities.Release(ref fill);
        }
    }

    private static void ApplyLine(PowerPoint.Shape shape, string? color, float width)
    {
        PowerPoint.LineFormat? line = null;
        PowerPoint.ColorFormat? fore = null;
        try
        {
            line = shape.Line;
            if (color is null || width <= 0)
            {
                line.Visible = Office.MsoTriState.msoFalse;
                return;
            }
            line.Visible = Office.MsoTriState.msoTrue;
            fore = line.ForeColor;
            fore.RGB = ChartStyler.Bgr(color);
            line.Weight = width;
        }
        finally
        {
            if (fore is not null) ComUtilities.Release(ref fore);
            if (line is not null) ComUtilities.Release(ref line);
        }
    }

    private static void SetEndArrow(PowerPoint.Shape shape)
    {
        PowerPoint.LineFormat? line = null;
        try
        {
            line = shape.Line;
            line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;
        }
        finally
        {
            if (line is not null) ComUtilities.Release(ref line);
        }
    }

    private static void SetRadius(PowerPoint.Shape shape, PlannedElement element)
    {
        PowerPoint.Adjustments? adjustments = null;
        try
        {
            adjustments = shape.Adjustments;
            if (adjustments.Count >= 1)
                adjustments[1] = Math.Clamp(element.Radius / Math.Max(1f, Math.Min(element.Box.Width, element.Box.Height)), 0f, 0.5f);
        }
        finally
        {
            if (adjustments is not null) ComUtilities.Release(ref adjustments);
        }
    }

    internal static void WriteText(PowerPoint.Shape shape, PlannedElement element, IReadOnlyList<PlannedParagraph> paragraphs)
    {
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        PowerPoint.TextFrame2? frame2 = null;
        try
        {
            frame = shape.TextFrame;
            frame.WordWrap = Office.MsoTriState.msoTrue;
            frame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
            frame.MarginLeft = element.Padding[0];
            frame.MarginTop = element.Padding[1];
            frame.MarginRight = element.Padding[2];
            frame.MarginBottom = element.Padding[3];
            frame.VerticalAnchor = element.VAlign switch
            {
                "middle" => Office.MsoVerticalAnchor.msoAnchorMiddle,
                "bottom" => Office.MsoVerticalAnchor.msoAnchorBottom,
                _ => Office.MsoVerticalAnchor.msoAnchorTop,
            };
            range = frame.TextRange;
            // Paragraphs are separated by CR; a newline inside one paragraph becomes a line break (VT).
            range.Text = string.Join("\r", paragraphs.Select(paragraph =>
                paragraph.Text.Replace("\r\n", "\v", StringComparison.Ordinal).Replace('\n', '\v').Replace('\r', '\v')));
            frame2 = shape.TextFrame2;
            for (int index = 1; index <= paragraphs.Count; index++)
                FormatParagraph(range, frame2, index, paragraphs[index - 1], element.LineSpacing);
        }
        finally
        {
            if (frame2 is not null) ComUtilities.Release(ref frame2);
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    private static void FormatParagraph(PowerPoint.TextRange range, PowerPoint.TextFrame2 frame2, int index, PlannedParagraph paragraph, float lineSpacing)
    {
        PowerPoint.TextRange? target = null;
        PowerPoint.Font? font = null;
        PowerPoint.ColorFormat? color = null;
        PowerPoint.ParagraphFormat? format = null;
        PowerPoint.BulletFormat? bullet = null;
        Office.TextRange2? all2 = null;
        Office.TextRange2? target2 = null;
        Office.ParagraphFormat2? format2 = null;
        try
        {
            target = range.Paragraphs(index, 1);
            font = target.Font;
            font.Name = paragraph.Font;
            font.Size = paragraph.Size;
            font.Bold = paragraph.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            font.Italic = paragraph.Italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            color = font.Color;
            color.RGB = ChartStyler.Bgr(paragraph.Color);

            format = target.ParagraphFormat;
            format.Alignment = paragraph.Align switch
            {
                "center" => PowerPoint.PpParagraphAlignment.ppAlignCenter,
                "right" => PowerPoint.PpParagraphAlignment.ppAlignRight,
                _ => PowerPoint.PpParagraphAlignment.ppAlignLeft,
            };
            format.LineRuleWithin = Office.MsoTriState.msoTrue;
            format.SpaceWithin = lineSpacing;
            format.LineRuleAfter = Office.MsoTriState.msoFalse;
            format.SpaceAfter = paragraph.SpaceAfter;
            format.LineRuleBefore = Office.MsoTriState.msoFalse;
            format.SpaceBefore = 0;
            bullet = format.Bullet;
            if (paragraph.Bullet == "none")
            {
                bullet.Visible = Office.MsoTriState.msoFalse;
            }
            else
            {
                bullet.Visible = Office.MsoTriState.msoTrue;
                bullet.Type = paragraph.Bullet == "number" ? PowerPoint.PpBulletType.ppBulletNumbered : PowerPoint.PpBulletType.ppBulletUnnumbered;
                if (paragraph.Bullet == "number")
                    bullet.Style = PowerPoint.PpNumberedBulletStyle.ppBulletArabicPeriod;
                else
                    bullet.Character = 8226; // U+2022 bullet
            }

            // Indents (hanging bullets) are only on the TextFrame2 paragraph format.
            all2 = frame2.TextRange;
            target2 = all2.Paragraphs[index, 1];
            format2 = target2.ParagraphFormat;
            var levelIndent = paragraph.Level * 18f;
            if (paragraph.Bullet == "none")
            {
                format2.LeftIndent = levelIndent;
                format2.FirstLineIndent = 0;
            }
            else
            {
                var hanging = Math.Max(12f, paragraph.Size * 1.1f);
                format2.LeftIndent = levelIndent + hanging;
                format2.FirstLineIndent = -hanging;
            }

            if (paragraph.Hyperlink is not null)
                SetHyperlink(target, paragraph.Hyperlink);
        }
        finally
        {
            if (format2 is not null) ComUtilities.Release(ref format2);
            if (target2 is not null) ComUtilities.Release(ref target2);
            if (all2 is not null) ComUtilities.Release(ref all2);
            if (bullet is not null) ComUtilities.Release(ref bullet);
            if (format is not null) ComUtilities.Release(ref format);
            if (color is not null) ComUtilities.Release(ref color);
            if (font is not null) ComUtilities.Release(ref font);
            if (target is not null) ComUtilities.Release(ref target);
        }
    }

    private static void SetHyperlink(PowerPoint.TextRange target, string address)
    {
        PowerPoint.ActionSettings? settings = null;
        PowerPoint.ActionSetting? click = null;
        PowerPoint.Hyperlink? link = null;
        try
        {
            settings = target.ActionSettings;
            click = settings[PowerPoint.PpMouseActivation.ppMouseClick];
            link = click.Hyperlink;
            link.Address = address;
        }
        finally
        {
            if (link is not null) ComUtilities.Release(ref link);
            if (click is not null) ComUtilities.Release(ref click);
            if (settings is not null) ComUtilities.Release(ref settings);
        }
    }

    private static PowerPoint.Shape CreatePicture(PowerPoint.Shapes shapes, PlannedElement element)
    {
        var picture = element.Picture!;
        var placement = picture.Placement;
        if (picture.Fit == "contain" || !placement.Cropped)
        {
            var frame = placement.Frame;
            var shape = shapes.AddPicture(picture.Path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, frame.Left, frame.Top, frame.Width, frame.Height);
            TagPixels(shape, picture);
            return shape;
        }

        var cover = shapes.AddPicture(picture.Path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, placement.Frame.Left, placement.Frame.Top, placement.PictureWidth, placement.PictureHeight);
        PowerPoint.PictureFormat? format = null;
        Office.Crop? crop = null;
        try
        {
            format = cover.PictureFormat;
            crop = format.Crop;
            crop.PictureWidth = placement.PictureWidth;
            crop.PictureHeight = placement.PictureHeight;
            crop.ShapeLeft = placement.Frame.Left;
            crop.ShapeTop = placement.Frame.Top;
            crop.ShapeWidth = placement.Frame.Width;
            crop.ShapeHeight = placement.Frame.Height;
            crop.PictureOffsetX = placement.OffsetX;
            crop.PictureOffsetY = placement.OffsetY;
        }
        finally
        {
            if (crop is not null) ComUtilities.Release(ref crop);
            if (format is not null) ComUtilities.Release(ref format);
        }
        TagPixels(cover, picture);
        return cover;
    }

    private static void TagPixels(PowerPoint.Shape shape, PicturePlan picture)
    {
        var tags = shape.Tags;
        try
        {
            tags.Add(DeckRoles.ImagePixelsTag, $"{picture.PixelWidth.ToString(CultureInfo.InvariantCulture)}x{picture.PixelHeight.ToString(CultureInfo.InvariantCulture)}");
            if (picture.Attribution is not null)
                tags.Add(DeckRoles.AttributionTag, picture.Attribution);
            tags.Add("PPTMCP_IMG_SOURCE", picture.Path);
        }
        finally
        {
            ComUtilities.Release(ref tags);
        }
    }

    private static PowerPoint.Shape CreateTable(PowerPoint.Shapes shapes, PlannedElement element)
    {
        var plan = element.Table!;
        bool hasHeader = plan.Header.Count > 0;
        int rows = plan.Rows.Count + (hasHeader ? 1 : 0);
        int columns = hasHeader ? plan.Header.Count : plan.Rows[0].Count;
        var box = element.Box;
        var shape = shapes.AddTable(rows, columns, box.Left, box.Top, box.Width, Math.Max(rows * 12f, 10f));
        TableStyler.Fill(shape, plan);
        return shape;
    }

    private static PowerPoint.Shape? CreateChart(PowerPoint.Shapes shapes, PlannedElement element, List<string> warnings)
    {
        var plan = element.Chart!;
        var box = element.Box;
        int type = ChartDataWriter.ChartTypes[plan.Spec.Type];
        // PIA gap: Shapes.AddChart2 is not exposed on the embedded PIA's Shapes interface.
        PowerPoint.Shape shape = ((dynamic)shapes).AddChart2(-1, type, box.Left, box.Top, box.Width, box.Height, true);
        PowerPoint.Chart? chart = null;
        try
        {
            chart = shape.Chart;
            ChartDataWriter.Write(chart, plan.Spec.Categories, plan.Spec.Series.Select(series => new ChartSeriesData(series.Name, series.Values)).ToList(), plan.Spec.NumberFormat);
            var notes = ChartStyler.Apply(chart, new ChartLook(plan.Palette, plan.Font, plan.FontSize, plan.TextColor, plan.Gridlines, plan.DataLabels,
                plan.Legend, plan.Spec.NumberFormat, plan.Spec.ValueAxisTitle, plan.Spec.CategoryAxisTitle), plan.Spec.Type is "pie" or "doughnut");
            warnings.AddRange(notes);
            var bindingTags = shape.Tags;
            try
            {
                bindingTags.Add(DeckRoles.BindingTag, "inline:composition");
            }
            finally
            {
                ComUtilities.Release(ref bindingTags);
            }
        }
        finally
        {
            if (chart is not null) ComUtilities.Release(ref chart);
        }
        return shape;
    }

    private static PowerPoint.Shape? CreateConnector(PowerPoint.Shapes shapes, PlannedElement element, Dictionary<string, PowerPoint.Shape> created, List<string> warnings)
    {
        if (element.FromKey is null || element.ToKey is null || !created.TryGetValue(element.FromKey, out var from) || !created.TryGetValue(element.ToKey, out var to))
        {
            warnings.Add($"Connector {element.Key} skipped: its end shapes were not created.");
            return null;
        }
        var kind = element.ConnectorKind == "straight" ? Office.MsoConnectorType.msoConnectorStraight
            : element.ConnectorKind == "curve" ? Office.MsoConnectorType.msoConnectorCurve
            : Office.MsoConnectorType.msoConnectorElbow;
        var connector = shapes.AddConnector(kind, 0, 0, 10, 10);
        PowerPoint.ConnectorFormat? format = null;
        try
        {
            format = connector.ConnectorFormat;
            format.BeginConnect(from, Math.Clamp(element.FromSite, 1, Math.Max(1, from.ConnectionSiteCount)));
            format.EndConnect(to, Math.Clamp(element.ToSite, 1, Math.Max(1, to.ConnectionSiteCount)));
            // Site 0 means "let PowerPoint choose the closest sites".
            if (element.FromSite == 0 || element.ToSite == 0)
                connector.RerouteConnections();
        }
        finally
        {
            if (format is not null) ComUtilities.Release(ref format);
        }
        return connector;
    }

    private static void Fit(
        FitRegion region,
        PlannedSlide planned,
        Dictionary<string, PowerPoint.Shape> shapes,
        Dictionary<string, RenderedElement> results,
        SpecFit fit,
        RenderedSlide rendered)
    {
        var keys = region.Keys.Where(shapes.ContainsKey).ToList();
        if (keys.Count == 0)
            return;
        bool canShrink = fit.Policy.Contains("shrink");
        switch (region.Mode)
        {
            case "each":
                FitEach(region, keys, shapes, results, canShrink, rendered);
                break;
            case "stack":
                FitStack(region, planned, keys, shapes, results, canShrink, rendered);
                break;
            case "paragraphs":
                FitParagraphs(region, keys[0], shapes, results, canShrink, rendered);
                break;
            case "table":
                FitTable(region, keys[0], planned.Elements.First(element => element.Key == keys[0]).Table?.Header.Count > 0, shapes, results, rendered);
                break;
        }
    }

    private static void FitEach(FitRegion region, List<string> keys, Dictionary<string, PowerPoint.Shape> shapes, Dictionary<string, RenderedElement> results, bool canShrink, RenderedSlide rendered)
    {
        bool AllFit() => keys.All(key => (TextRuns.NeededHeight(shapes[key]) ?? 0) <= shapes[key].Height + Tolerance);
        if (AllFit())
        {
            Mark(keys, shapes, results, "fits");
            return;
        }
        if (canShrink && Shrink(Capture(keys, shapes), region.MinFontSize, AllFit) is { } factor)
        {
            Mark(keys, shapes, results, "shrunk", factor);
            return;
        }
        foreach (var key in keys)
        {
            var needed = TextRuns.NeededHeight(shapes[key]) ?? 0;
            var shape = shapes[key];
            if (needed > shape.Height + Tolerance)
            {
                results[key].Fit = "overflow";
                results[key].MeasuredHeight = Round(needed);
                rendered.Overflow.Add($"{key}: text needs {Fmt(needed)} pt but the box is {Fmt(shape.Height)} pt tall at the {Fmt(region.MinFontSize)} pt minimum. Nothing was cut; shorten the text, raise the box, or allow a split.");
            }
            else
            {
                results[key].Fit = canShrink ? "shrunk" : "fits";
                results[key].MeasuredHeight = Round(needed);
            }
        }
    }

    private static void FitStack(FitRegion region, PlannedSlide planned, List<string> keys, Dictionary<string, PowerPoint.Shape> shapes, Dictionary<string, RenderedElement> results, bool canShrink, RenderedSlide rendered)
    {
        float Restack()
        {
            var y = region.Box.Top;
            foreach (var key in keys)
            {
                var shape = shapes[key];
                var needed = MathF.Ceiling(TextRuns.NeededHeight(shape) ?? shape.Height);
                shape.Top = y;
                shape.Height = Math.Max(needed, 1f);
                foreach (var anchored in planned.Elements.Where(element => element.AnchorKey == key && shapes.ContainsKey(element.Key)))
                    shapes[anchored.Key].Top = y + anchored.AnchorOffset;
                y += needed + region.Gap;
            }
            return y - region.Gap;
        }

        if (Restack() <= region.Box.Bottom + Tolerance)
        {
            Mark(keys, shapes, results, "restacked");
            return;
        }
        var snapshot = Capture(keys, shapes);
        if (canShrink && Shrink(snapshot, region.MinFontSize, () => Restack() <= region.Box.Bottom + Tolerance) is { } factor)
        {
            Mark(keys, shapes, results, "shrunk", factor);
            return;
        }

        if (region.Splittable && keys.Count > 1)
        {
            // Restore planned sizes, then count how many items fit when measured.
            snapshot.Apply(1f);
            Restack();
            int fit = keys.TakeWhile(key => shapes[key].Top + shapes[key].Height <= region.Box.Bottom + Tolerance).Count();
            rendered.SplitAfter = Math.Max(1, fit);
            return;
        }

        var bottom = Restack();
        Mark(keys, shapes, results, "overflow");
        rendered.Overflow.Add($"{region.Key}: the stacked items need {Fmt(bottom - region.Box.Top)} pt but {Fmt(region.Box.Height)} pt are available at the {Fmt(region.MinFontSize)} pt minimum. Nothing was cut; allow \"split\" in fit.policy or shorten the items.");
    }

    private static void FitParagraphs(FitRegion region, string key, Dictionary<string, PowerPoint.Shape> shapes, Dictionary<string, RenderedElement> results, bool canShrink, RenderedSlide rendered)
    {
        var shape = shapes[key];
        bool Fits() => (TextRuns.NeededHeight(shape) ?? 0) <= shape.Height + Tolerance;
        if (Fits())
        {
            Mark([key], shapes, results, "fits");
            return;
        }
        var snapshot = Capture([key], shapes);
        if (canShrink && Shrink(snapshot, region.MinFontSize, Fits) is { } factor)
        {
            Mark([key], shapes, results, "shrunk", factor);
            return;
        }
        if (region.Splittable)
        {
            snapshot.Apply(1f);
            rendered.SplitAfter = Math.Max(1, TextRuns.WithRange(shape, (frame, range) =>
            {
                int count = range.Paragraphs().Count;
                int fit = 0;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.TextRange? paragraph = null;
                    try
                    {
                        paragraph = range.Paragraphs(index, 1);
                        if (paragraph.BoundTop + paragraph.BoundHeight + frame.MarginBottom > shape.Top + shape.Height + Tolerance)
                            break;
                        fit++;
                    }
                    finally
                    {
                        if (paragraph is not null) ComUtilities.Release(ref paragraph);
                    }
                }
                return fit;
            }));
            return;
        }
        Mark([key], shapes, results, "overflow");
        rendered.Overflow.Add($"{key}: text needs {Fmt(TextRuns.NeededHeight(shape) ?? 0)} pt but {Fmt(shape.Height)} pt are available. Nothing was cut.");
    }

    private static void FitTable(FitRegion region, string key, bool hasHeader, Dictionary<string, PowerPoint.Shape> shapes, Dictionary<string, RenderedElement> results, RenderedSlide rendered)
    {
        var shape = shapes[key];
        results[key].MeasuredHeight = Round(shape.Height);
        if (shape.Top + shape.Height <= region.Box.Bottom + Tolerance)
        {
            results[key].Fit = "fits";
            return;
        }
        if (region.Splittable)
        {
            var (fit, total) = TableStyler.RowsThatFit(shape, region.Box.Bottom, hasHeader);
            if (fit < total)
            {
                rendered.SplitAfter = Math.Max(1, fit);
                return;
            }
        }
        results[key].Fit = "overflow";
        rendered.Overflow.Add($"{key}: the table is {Fmt(shape.Height)} pt tall and ends {Fmt(shape.Top + shape.Height - region.Box.Bottom)} pt below its area. Nothing was cut; allow \"split\" or reduce rows.");
    }

    /// <summary>Original run sizes of a set of shapes, so fitting can scale from and restore them.</summary>
    private sealed class SizeSnapshot
    {
        public List<(PowerPoint.Shape Shape, List<float> Sizes)> Entries { get; } = [];

        public float Smallest => Entries.SelectMany(entry => entry.Sizes).Where(size => size > 0).DefaultIfEmpty(0).Min();

        public void Apply(float factor)
        {
            foreach (var (shape, sizes) in Entries)
                TextRuns.WithRange(shape, (_, range) => { TextRuns.WriteSizes(range, sizes, factor); return true; });
        }
    }

    private static SizeSnapshot Capture(IEnumerable<string> keys, Dictionary<string, PowerPoint.Shape> shapes)
    {
        var snapshot = new SizeSnapshot();
        foreach (var key in keys)
        {
            var shape = shapes[key];
            if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                continue;
            if (TextRuns.WithRange(shape, (_, range) => TextRuns.ReadSizes(range)) is { } sizes)
                snapshot.Entries.Add((shape, sizes));
        }
        return snapshot;
    }

    /// <summary>
    /// Scales all runs by one factor, stepping down 2.5% at a time until <paramref name="fits"/>
    /// holds or the smallest text reaches the floor. Returns the factor, or null when even the
    /// floor does not fit (sizes are then left at the floor; callers may restore with Apply(1)).
    /// </summary>
    private static float? Shrink(SizeSnapshot snapshot, float floor, Func<bool> fits)
    {
        var smallest = snapshot.Smallest;
        if (snapshot.Entries.Count == 0 || smallest <= 0)
            return null;
        var minimum = Math.Min(1f, floor / smallest);
        for (var factor = 0.975f; factor >= minimum - 0.0001f; factor -= 0.025f)
        {
            snapshot.Apply(factor);
            if (fits())
                return factor;
        }
        snapshot.Apply(minimum);
        return fits() ? minimum : null;
    }

    private static void Mark(IEnumerable<string> keys, Dictionary<string, PowerPoint.Shape> shapes, Dictionary<string, RenderedElement> results, string fit, float? factor = null)
    {
        foreach (var key in keys)
        {
            var shape = shapes[key];
            results[key].Fit = fit;
            results[key].MeasuredHeight = Round(TextRuns.NeededHeight(shape) ?? shape.Height);
            if (factor is not null && shape.HasTextFrame == Office.MsoTriState.msoTrue)
                results[key].MinFontSize = TextRuns.WithRange(shape, (_, range) => TextRuns.ReadSizes(range)?.Where(size => size > 0).DefaultIfEmpty(0).Min());
        }
    }

    private static void GroupParts(PowerPoint.Slide slide, PlannedSlide planned, Dictionary<string, PowerPoint.Shape> shapes, List<string> warnings)
    {
        foreach (var group in planned.Elements.Where(element => element.Group is not null && shapes.ContainsKey(element.Key)).GroupBy(element => element.Group!))
        {
            var members = group.ToList();
            if (members.Count < 2)
                continue;
            PowerPoint.Shapes? collection = null;
            PowerPoint.ShapeRange? range = null;
            PowerPoint.Shape? grouped = null;
            try
            {
                collection = slide.Shapes;
                range = collection.Range(members.Select(member => (object)shapes[member.Key].Name).ToArray());
                grouped = range.Group();
                grouped.Name = group.Key;
                var tags = grouped.Tags;
                try
                {
                    tags.Add(DeckRoles.IdTag, AppIdFor(planned, group.Key));
                    tags.Add(DeckRoles.RoleTag, members[0].Role);
                    if (members[0].Component is { } component)
                        tags.Add(DeckRoles.ComponentTag, component);
                }
                finally
                {
                    ComUtilities.Release(ref tags);
                }
            }
            catch (COMException ex)
            {
                warnings.Add($"Group {group.Key} was not created: {ex.Message}");
            }
            finally
            {
                if (grouped is not null) ComUtilities.Release(ref grouped);
                if (range is not null) ComUtilities.Release(ref range);
                if (collection is not null) ComUtilities.Release(ref collection);
            }
        }
    }

    private static void TagSlide(PowerPoint.Slide slide, PlannedSlide planned, CompositionPlan plan, ResolvedProfile profile)
    {
        var tags = slide.Tags;
        try
        {
            tags.Add(DeckRoles.IdTag, planned.AppId);
            tags.Add("PPTMCP_COMPOSITION", $"{plan.Spec.Kind}@1");
            tags.Add("PPTMCP_PROFILE", $"{profile.Name}@{profile.Version}");
            tags.Add("PPTMCP_PART", $"{(planned.Sequence + 1).ToString(CultureInfo.InvariantCulture)}/{plan.Slides.Count.ToString(CultureInfo.InvariantCulture)}");
            if (planned.Sequence == 0 && plan.Spec.OriginalJson is { } json)
                tags.Add("PPTMCP_SPEC", json);
            foreach (var (key, value) in plan.Spec.Metadata ?? new Dictionary<string, string>())
                tags.Add($"PPTMCP_META_{key.ToUpperInvariant()}", value);
        }
        finally
        {
            ComUtilities.Release(ref tags);
        }
    }

    private static void WriteNotes(PowerPoint.Slide slide, string notes, List<string> warnings)
    {
        PowerPoint.SlideRange? page = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            page = slide.NotesPage;
            shapes = page.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.PlaceholderFormat? format = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        continue;
                    format = shape.PlaceholderFormat;
                    if (format.Type != PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        continue;
                    TextRuns.WithRange(shape, (_, range) => { range.Text = notes; return true; });
                    return;
                }
                finally
                {
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
            warnings.Add("The notes page has no body placeholder; speaker notes were not written.");
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (page is not null) ComUtilities.Release(ref page);
        }
    }

    private static void ApplyFooter(PowerPoint.Slide slide, ResolvedProfile profile, List<string> warnings)
    {
        PowerPoint.HeadersFooters? footers = null;
        PowerPoint.HeaderFooter? footer = null;
        PowerPoint.HeaderFooter? number = null;
        try
        {
            footers = slide.HeadersFooters;
            if (profile.Footer.Text is { } text)
            {
                footer = footers.Footer;
                footer.Visible = Office.MsoTriState.msoTrue;
                footer.Text = text;
            }
            number = footers.SlideNumber;
            number.Visible = profile.Footer.SlideNumber ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
        }
        catch (COMException ex)
        {
            warnings.Add($"Footer not applied (the layout may have no footer placeholders): {ex.Message}");
        }
        finally
        {
            if (number is not null) ComUtilities.Release(ref number);
            if (footer is not null) ComUtilities.Release(ref footer);
            if (footers is not null) ComUtilities.Release(ref footers);
        }
    }

    private static float Round(float value) => MathF.Round(value * 10f) / 10f;

    private static string Fmt(float value) => value.ToString("0.#", CultureInfo.InvariantCulture);
}
