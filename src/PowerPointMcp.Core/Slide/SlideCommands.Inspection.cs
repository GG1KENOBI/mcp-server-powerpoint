extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

public sealed partial class SlideCommands
{
    // Font sizes are read per text run; very long text is sampled up to this many runs.
    private const int MaxFontRunsRead = 200;

    /// <inheritdoc/>
    public SlideOperationResult Inspect(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.PageSetup? pageSetup = null;
            try
            {
                slides = ctx.Presentation.Slides;
                var validation = ValidateSlideIndex(slides.Count, slideIndex);
                if (validation is not null) return validation;

                pageSetup = ctx.Presentation.PageSetup;
                slide = slides[slideIndex];
                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    SlideCount = slides.Count,
                    SlideWidth = pageSetup.SlideWidth,
                    SlideHeight = pageSetup.SlideHeight,
                    Shapes = ReadSlideShapes(slide, ct),
                };
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
                if (pageSetup is not null) ComUtilities.Release(ref pageSetup);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult CheckLayout(IPresentationBatch batch, int? slideIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.PageSetup? pageSetup = null;
            try
            {
                slides = ctx.Presentation.Slides;
                int slideCount = slides.Count;
                if (slideIndex is { } requested)
                {
                    var validation = ValidateSlideIndex(slideCount, requested);
                    if (validation is not null) return validation;
                }

                pageSetup = ctx.Presentation.PageSetup;
                float slideWidth = pageSetup.SlideWidth;
                float slideHeight = pageSetup.SlideHeight;
                int first = slideIndex ?? 1;
                int last = slideIndex ?? slideCount;
                var issues = new List<SlideLayoutIssue>();
                for (int index = first; index <= last; index++)
                {
                    PowerPoint.Slide? slide = null;
                    try
                    {
                        slide = slides[index];
                        issues.AddRange(SlideLayoutAnalyzer.Analyze(index, slideWidth, slideHeight, ReadSlideShapes(slide, ct)));
                    }
                    finally
                    {
                        if (slide is not null) ComUtilities.Release(ref slide);
                    }
                }

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    SlideCount = slideCount,
                    SlideWidth = slideWidth,
                    SlideHeight = slideHeight,
                    LayoutIssues = issues,
                    LayoutOk = !issues.Any(issue => issue.Severity is "error" or "warning"),
                };
            }
            finally
            {
                if (pageSetup is not null) ComUtilities.Release(ref pageSetup);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    private static List<SlideShapeInfo> ReadSlideShapes(PowerPoint.Slide slide, CancellationToken cancellationToken)
    {
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            var result = new List<SlideShapeInfo>(shapes.Count);
            for (int shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = shapes[shapeIndex];
                    result.Add(ReadShape(shape, shapeIndex));
                }
                finally
                {
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }

            return result;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    private static SlideShapeInfo ReadShape(PowerPoint.Shape shape, int shapeIndex)
    {
        Office.MsoShapeType type = shape.Type;
        bool isPlaceholder = type == Office.MsoShapeType.msoPlaceholder;
        var text = ReadText(shape);
        var (rows, columns) = ReadTableSize(shape);
        return new SlideShapeInfo
        {
            ShapeIndex = shapeIndex,
            Name = shape.Name,
            Kind = KindOf(type, shape.Connector == Office.MsoTriState.msoTrue),
            PlaceholderType = isPlaceholder ? ReadPlaceholderType(shape) : null,
            Left = shape.Left,
            Top = shape.Top,
            Width = shape.Width,
            Height = shape.Height,
            Rotation = shape.Rotation,
            ZOrder = shape.ZOrderPosition,
            Visible = shape.Visible == Office.MsoTriState.msoTrue,
            HasText = text.Text is not null,
            Text = text.Text,
            MinFontSize = text.MinFontSize,
            MaxFontSize = text.MaxFontSize,
            AutoSize = text.AutoSize,
            TextLeft = text.Left,
            TextTop = text.Top,
            TextWidth = text.Width,
            TextHeight = text.Height,
            HasPicture = isPlaceholder ? HasPictureFill(shape) : null,
            HasChart = shape.HasChart == Office.MsoTriState.msoTrue ? true : null,
            TableRows = rows,
            TableColumns = columns,
        };
    }

    private static string KindOf(Office.MsoShapeType type, bool isConnector) => type switch
    {
        Office.MsoShapeType.msoPlaceholder => "placeholder",
        Office.MsoShapeType.msoTextBox => "text-box",
        Office.MsoShapeType.msoAutoShape when isConnector => "connector",
        Office.MsoShapeType.msoAutoShape or Office.MsoShapeType.msoCallout => "auto-shape",
        Office.MsoShapeType.msoLine => "line",
        Office.MsoShapeType.msoPicture or Office.MsoShapeType.msoLinkedPicture => "picture",
        Office.MsoShapeType.msoTable => "table",
        Office.MsoShapeType.msoChart => "chart",
        Office.MsoShapeType.msoSmartArt => "smart-art",
        Office.MsoShapeType.msoGroup => "group",
        Office.MsoShapeType.msoMedia => "media",
        Office.MsoShapeType.msoTextEffect => "text-effect",
        Office.MsoShapeType.msoFreeform => "freeform",
        _ => "other",
    };

    private static string ReadPlaceholderType(PowerPoint.Shape shape)
    {
        PowerPoint.PlaceholderFormat? format = null;
        try
        {
            format = shape.PlaceholderFormat;
            return format.Type.ToString();
        }
        finally
        {
            if (format is not null) ComUtilities.Release(ref format);
        }
    }

    private static bool HasPictureFill(PowerPoint.Shape shape)
    {
        PowerPoint.FillFormat? fill = null;
        try
        {
            fill = shape.Fill;
            return fill.Type == Office.MsoFillType.msoFillPicture;
        }
        finally
        {
            if (fill is not null) ComUtilities.Release(ref fill);
        }
    }

    private static (int? Rows, int? Columns) ReadTableSize(PowerPoint.Shape shape)
    {
        if (shape.HasTable != Office.MsoTriState.msoTrue)
            return (null, null);

        PowerPoint.Table? table = null;
        PowerPoint.Rows? rows = null;
        PowerPoint.Columns? columns = null;
        try
        {
            table = shape.Table;
            rows = table.Rows;
            columns = table.Columns;
            return (rows.Count, columns.Count);
        }
        finally
        {
            if (columns is not null) ComUtilities.Release(ref columns);
            if (rows is not null) ComUtilities.Release(ref rows);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static ShapeTextSnapshot ReadText(PowerPoint.Shape shape)
    {
        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
            return default;

        PowerPoint.TextFrame? textFrame = null;
        PowerPoint.TextFrame2? textFrame2 = null;
        PowerPoint.TextRange? textRange = null;
        try
        {
            textFrame = shape.TextFrame;
            if (textFrame.HasText != Office.MsoTriState.msoTrue)
                return default;

            textRange = textFrame.TextRange;
            string text = textRange.Text;
            if (string.IsNullOrWhiteSpace(text))
                return default;

            textFrame2 = shape.TextFrame2;
            var (minFont, maxFont) = ReadFontSizeRange(textRange);
            return new ShapeTextSnapshot(
                text,
                minFont,
                maxFont,
                AutoSizeName(textFrame2.AutoSize),
                textRange.BoundLeft,
                textRange.BoundTop,
                textRange.BoundWidth,
                textRange.BoundHeight);
        }
        finally
        {
            if (textRange is not null) ComUtilities.Release(ref textRange);
            if (textFrame2 is not null) ComUtilities.Release(ref textFrame2);
            if (textFrame is not null) ComUtilities.Release(ref textFrame);
        }
    }

    private static (float? Min, float? Max) ReadFontSizeRange(PowerPoint.TextRange textRange)
    {
        PowerPoint.TextRange? allRuns = null;
        float? min = null;
        float? max = null;
        try
        {
            allRuns = textRange.Runs();
            int runCount = Math.Min(allRuns.Count, MaxFontRunsRead);
            for (int runIndex = 1; runIndex <= runCount; runIndex++)
            {
                PowerPoint.TextRange? run = null;
                PowerPoint.Font? font = null;
                try
                {
                    run = textRange.Runs(runIndex, 1);
                    font = run.Font;
                    float size = font.Size;
                    if (size > 0)
                    {
                        min = min is { } currentMin ? Math.Min(currentMin, size) : size;
                        max = max is { } currentMax ? Math.Max(currentMax, size) : size;
                    }
                }
                finally
                {
                    if (font is not null) ComUtilities.Release(ref font);
                    if (run is not null) ComUtilities.Release(ref run);
                }
            }

            return (min, max);
        }
        finally
        {
            if (allRuns is not null) ComUtilities.Release(ref allRuns);
        }
    }

    private static string AutoSizeName(Office.MsoAutoSize autoSize) => autoSize switch
    {
        Office.MsoAutoSize.msoAutoSizeNone => "none",
        Office.MsoAutoSize.msoAutoSizeShapeToFitText => "shape-to-fit-text",
        Office.MsoAutoSize.msoAutoSizeTextToFitShape => "shrink-text-on-overflow",
        _ => "mixed",
    };

    private readonly record struct ShapeTextSnapshot(
        string? Text,
        float? MinFontSize,
        float? MaxFontSize,
        string? AutoSize,
        float? Left,
        float? Top,
        float? Width,
        float? Height);
}
