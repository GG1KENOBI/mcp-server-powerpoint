extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Deck;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>A table, chart, or component instance found by the scanner.</summary>
internal sealed record ScannedObject(int SlideIndex, int SlideId, int ShapeId, string ShapeName, string Kind, string? Component, bool HeaderRow, bool TotalRow, bool Banded, bool IsPie);

/// <summary>Everything the design commands read from a deck.</summary>
internal sealed class DeckStyleScan
{
    public List<StyledRun> Runs { get; } = [];
    public List<StyledFill> Fills { get; } = [];
    public List<StyledSlide> Slides { get; } = [];
    public List<ScannedObject> Objects { get; } = [];
    public List<(float Left, float Top, float Width, float Height)> TitleBoxes { get; } = [];
    public List<(float Left, float Top, float Width, float Height)> ContentBoxes { get; } = [];
    public int SkippedRuns { get; set; }
}

/// <summary>
/// Reads formatting runs (font, size, color and whether the color comes from the theme), solid
/// fills and outlines, slide backgrounds and profile tags, tables, charts, and component tags
/// from every slide (optionally only some slides or shapes). Group members and table cells are
/// included. STA thread only.
/// </summary>
internal static class DesignScanner
{
    internal static DeckStyleScan Scan(PowerPoint.Presentation presentation, Func<int, int, bool>? includeShape, Func<int, bool>? includeSlide, CancellationToken cancellationToken)
    {
        var scan = new DeckStyleScan();
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (int slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPoint.Slide? slide = null;
                PowerPoint.Shapes? shapes = null;
                PowerPoint.Tags? tags = null;
                try
                {
                    slide = slides[slideIndex];
                    int slideId = slide.SlideID;
                    if (includeSlide is not null && !includeSlide(slideId))
                        continue;
                    tags = slide.Tags;
                    var profileTag = tags["PPTMCP_PROFILE"];
                    scan.Slides.Add(new StyledSlide(slideIndex, slideId, string.IsNullOrEmpty(profileTag) ? null : profileTag, ReadBackground(slide)));
                    shapes = slide.Shapes;
                    for (int index = 1; index <= shapes.Count; index++)
                    {
                        PowerPoint.Shape? shape = null;
                        try
                        {
                            shape = shapes[index];
                            ScanShape(shape, slideIndex, slideId, includeShape, null, scan, cancellationToken);
                        }
                        finally
                        {
                            if (shape is not null) ComUtilities.Release(ref shape);
                        }
                    }
                }
                finally
                {
                    if (tags is not null) ComUtilities.Release(ref tags);
                    if (shapes is not null) ComUtilities.Release(ref shapes);
                    if (slide is not null) ComUtilities.Release(ref slide);
                }
            }
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
        return scan;
    }

    private static void ScanShape(PowerPoint.Shape shape, int slideIndex, int slideId, Func<int, int, bool>? includeShape, string? groupComponent, DeckStyleScan scan, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        int shapeId = shape.Id;
        bool included = includeShape is null || includeShape(slideId, shapeId);
        var (role, component) = ReadRoleAndComponent(shape);
        component ??= groupComponent;
        string name = shape.Name;

        if (shape.Type == Office.MsoShapeType.msoGroup)
        {
            if (included && component is not null)
                scan.Objects.Add(new ScannedObject(slideIndex, slideId, shapeId, name, "component", component, false, false, false, false));
            PowerPoint.GroupShapes? members = null;
            try
            {
                members = shape.GroupItems;
                for (int index = 1; index <= members.Count; index++)
                {
                    PowerPoint.Shape? member = null;
                    try
                    {
                        member = members[index];
                        ScanShape(member, slideIndex, slideId, included ? null : includeShape, component, scan, cancellationToken);
                    }
                    finally
                    {
                        if (member is not null) ComUtilities.Release(ref member);
                    }
                }
            }
            finally
            {
                if (members is not null) ComUtilities.Release(ref members);
            }
            return;
        }
        if (!included)
            return;

        if (component is not null && groupComponent is null)
            scan.Objects.Add(new ScannedObject(slideIndex, slideId, shapeId, name, "component", component, false, false, false, false));

        if (role is "title" or "subtitle")
            scan.TitleBoxes.Add((shape.Left, shape.Top, shape.Width, shape.Height));
        else if (role is "body" or "object" or "table" or "chart" or "picture" && groupComponent is null)
            scan.ContentBoxes.Add((shape.Left, shape.Top, shape.Width, shape.Height));

        if (shape.HasTable == Office.MsoTriState.msoTrue)
        {
            ScanTable(shape, slideIndex, slideId, shapeId, name, scan, cancellationToken);
            return;
        }
        if (shape.HasChart == Office.MsoTriState.msoTrue)
        {
            ScanChart(shape, slideIndex, slideId, shapeId, name, scan);
            return;
        }
        if (shape.Type is not (Office.MsoShapeType.msoPicture or Office.MsoShapeType.msoLinkedPicture or Office.MsoShapeType.msoMedia))
            ScanFillAndLine(shape, slideIndex, slideId, shapeId, name, component, scan);
        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
            ScanText(shape, new RunLocation(slideIndex, slideId, shapeId, name, role, null, null), scan);
    }

    private sealed record RunLocation(int SlideIndex, int SlideId, int ShapeId, string ShapeName, string? Role, int? Row, int? Column);

    private static void ScanText(PowerPoint.Shape shape, RunLocation location, DeckStyleScan scan)
    {
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        PowerPoint.TextRange? runs = null;
        try
        {
            frame = shape.TextFrame;
            if (frame.HasText != Office.MsoTriState.msoTrue)
                return;
            range = frame.TextRange;
            runs = range.Runs();
            int count = runs.Count;
            if (count > TextRuns.MaxRuns)
            {
                scan.SkippedRuns += count;
                return;
            }
            for (int index = 1; index <= count; index++)
            {
                PowerPoint.TextRange? run = null;
                PowerPoint.Font? font = null;
                PowerPoint.ColorFormat? color = null;
                try
                {
                    run = range.Runs(index, 1);
                    font = run.Font;
                    color = font.Color;
                    bool fromTheme = color.Type == Office.MsoColorType.msoColorTypeScheme || color.ObjectThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor;
                    string text = run.Text;
                    scan.Runs.Add(new StyledRun(location.SlideIndex, location.SlideId, location.ShapeId, location.ShapeName, location.Role, location.Row, location.Column,
                        run.Start, run.Length, font.Name, font.Size, DeckSnapshotReader.Hex(color.RGB), fromTheme, text));
                }
                finally
                {
                    if (color is not null) ComUtilities.Release(ref color);
                    if (font is not null) ComUtilities.Release(ref font);
                    if (run is not null) ComUtilities.Release(ref run);
                }
            }
        }
        finally
        {
            if (runs is not null) ComUtilities.Release(ref runs);
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    private static void ScanTable(PowerPoint.Shape shape, int slideIndex, int slideId, int shapeId, string name, DeckStyleScan scan, CancellationToken cancellationToken)
    {
        PowerPoint.Table? table = null;
        PowerPoint.Rows? rows = null;
        PowerPoint.Columns? columns = null;
        try
        {
            table = shape.Table;
            rows = table.Rows;
            columns = table.Columns;
            scan.Objects.Add(new ScannedObject(slideIndex, slideId, shapeId, name, "table", null, table.FirstRow, table.LastRow, table.HorizBanding, false));
            for (int row = 1; row <= rows.Count; row++)
            {
                for (int column = 1; column <= columns.Count; column++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    PowerPoint.Cell? cell = null;
                    PowerPoint.Shape? cellShape = null;
                    try
                    {
                        cell = table.Cell(row, column);
                        cellShape = cell.Shape;
                        ScanText(cellShape, new RunLocation(slideIndex, slideId, shapeId, name, "table", row, column), scan);
                    }
                    finally
                    {
                        if (cellShape is not null) ComUtilities.Release(ref cellShape);
                        if (cell is not null) ComUtilities.Release(ref cell);
                    }
                }
            }
        }
        finally
        {
            if (columns is not null) ComUtilities.Release(ref columns);
            if (rows is not null) ComUtilities.Release(ref rows);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static void ScanChart(PowerPoint.Shape shape, int slideIndex, int slideId, int shapeId, string name, DeckStyleScan scan)
    {
        PowerPoint.Chart? chart = null;
        try
        {
            chart = shape.Chart;
            bool isPie = (int)chart.ChartType is 5 or -4120 or 69 or 70;
            scan.Objects.Add(new ScannedObject(slideIndex, slideId, shapeId, name, "chart", null, false, false, false, isPie));
        }
        finally
        {
            if (chart is not null) ComUtilities.Release(ref chart);
        }
    }

    private static void ScanFillAndLine(PowerPoint.Shape shape, int slideIndex, int slideId, int shapeId, string name, string? component, DeckStyleScan scan)
    {
        PowerPoint.FillFormat? fill = null;
        PowerPoint.LineFormat? line = null;
        PowerPoint.ColorFormat? color = null;
        try
        {
            fill = shape.Fill;
            if (fill.Visible == Office.MsoTriState.msoTrue && fill.Type == Office.MsoFillType.msoFillSolid)
            {
                color = fill.ForeColor;
                scan.Fills.Add(new StyledFill(slideIndex, slideId, shapeId, name, "fill", DeckSnapshotReader.Hex(color.RGB), IsTheme(color), component));
                ComUtilities.Release(ref color!);
            }
            line = shape.Line;
            if (line.Visible == Office.MsoTriState.msoTrue)
            {
                color = line.ForeColor;
                scan.Fills.Add(new StyledFill(slideIndex, slideId, shapeId, name, "line", DeckSnapshotReader.Hex(color.RGB), IsTheme(color), component));
            }
        }
        finally
        {
            if (color is not null) ComUtilities.Release(ref color);
            if (line is not null) ComUtilities.Release(ref line);
            if (fill is not null) ComUtilities.Release(ref fill);
        }
    }

    private static bool IsTheme(PowerPoint.ColorFormat color) =>
        color.Type == Office.MsoColorType.msoColorTypeScheme || color.ObjectThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor;

    private static string? ReadBackground(PowerPoint.Slide slide)
    {
        if (slide.FollowMasterBackground == Office.MsoTriState.msoTrue)
            return null;
        PowerPoint.ShapeRange? background = null;
        PowerPoint.FillFormat? fill = null;
        PowerPoint.ColorFormat? color = null;
        try
        {
            background = slide.Background;
            fill = background.Fill;
            if (fill.Type != Office.MsoFillType.msoFillSolid)
                return null;
            color = fill.ForeColor;
            return DeckSnapshotReader.Hex(color.RGB);
        }
        finally
        {
            if (color is not null) ComUtilities.Release(ref color);
            if (fill is not null) ComUtilities.Release(ref fill);
            if (background is not null) ComUtilities.Release(ref background);
        }
    }

    private static (string? Role, string? Component) ReadRoleAndComponent(PowerPoint.Shape shape)
    {
        PowerPoint.Tags? tags = null;
        PowerPoint.PlaceholderFormat? placeholder = null;
        try
        {
            tags = shape.Tags;
            var role = tags[DeckRoles.RoleTag];
            var component = tags[DeckRoles.ComponentTag];
            if (string.IsNullOrEmpty(role) && shape.Type == Office.MsoShapeType.msoPlaceholder)
            {
                placeholder = shape.PlaceholderFormat;
                role = DeckRoles.FromPlaceholder(placeholder.Type.ToString());
            }
            return (string.IsNullOrEmpty(role) ? null : role, string.IsNullOrEmpty(component) ? null : component);
        }
        finally
        {
            if (placeholder is not null) ComUtilities.Release(ref placeholder);
            if (tags is not null) ComUtilities.Release(ref tags);
        }
    }
}
