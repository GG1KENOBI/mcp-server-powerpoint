extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Data;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Table;

public sealed partial class TableCommands
{
    /// <inheritdoc/>
    public TableOperationResult ApplyStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, string? profile = null, bool headerRow = true, bool totalRow = false,
        bool banded = true, float? fontSize = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (fontSize is < 6 or > 72)
            return TableFail("font_size must be between 6 and 72 points.");
        var (resolved, error, _) = ProfileResolver.Resolve(batch, profile);
        if (error is not null)
            return TableFail(error);
        return WithTable(batch, slideIndex, shapeIndex, (table, shape, rows, columns) =>
        {
            TableStyler.StyleExisting(table, resolved!.Table, resolved.BodyFont, fontSize ?? resolved.Table.FontSize, headerRow, totalRow, banded, null);
            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, RowCount = rows, ColumnCount = columns, CellsChanged = rows * columns };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult SetColumnWidths(IPresentationBatch batch, int slideIndex, int shapeIndex, IReadOnlyList<double> widths, string mode = "points")
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(widths);
        if (mode is not ("points" or "weights"))
            return TableFail("mode must be points or weights.");
        if (widths.Any(width => width <= 0))
            return TableFail("widths must all be positive.");
        return WithTable(batch, slideIndex, shapeIndex, (table, shape, rows, columns) =>
        {
            if (widths.Count != columns)
                return TableFail($"widths has {widths.Count} values but the table has {columns} columns.");
            var total = widths.Sum();
            var tableWidth = shape.Width;
            PowerPoint.Columns? all = null;
            try
            {
                all = table.Columns;
                for (int index = 1; index <= columns; index++)
                {
                    PowerPoint.Column? column = null;
                    try
                    {
                        column = all[index];
                        column.Width = (float)(mode == "weights" ? tableWidth * widths[index - 1] / total : widths[index - 1]);
                    }
                    finally
                    {
                        if (column is not null) ComUtilities.Release(ref column);
                    }
                }
            }
            finally
            {
                if (all is not null) ComUtilities.Release(ref all);
            }
            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, ColumnCount = columns };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult FormatNumbers(IPresentationBatch batch, int slideIndex, int shapeIndex, int column, string format, string culture = "invariant", int firstRow = 2)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(format))
            return TableFail("format is required, e.g. #,##0.0 or 0%.");
        CultureInfo cultureInfo;
        try
        {
            cultureInfo = DataLoader.Culture(culture);
        }
        catch (ArgumentException ex)
        {
            return TableFail(ex.Message);
        }
        return WithTable(batch, slideIndex, shapeIndex, (table, shape, rows, columns) =>
        {
            if (column < 1 || column > columns)
                return TableFail($"Column {column} is out of range (1-{columns}).");
            int changed = 0;
            var skipped = new List<string>();
            for (int row = Math.Max(1, firstRow); row <= rows; row++)
            {
                var text = DeckSnapshotReader.ReadCellText(table, row, column);
                if (string.IsNullOrWhiteSpace(text))
                    continue;
                if (!DataTyping.TryParseNumber(text, cultureInfo, out var value, out _))
                {
                    skipped.Add($"R{row}C{column}: \"{text}\"");
                    continue;
                }
                var formatted = NumberFormatter.Format(value, format, cultureInfo);
                if (formatted == text)
                    continue;
                WithCellRange(table, row, column, range => range.Text = formatted);
                changed++;
            }
            return new TableOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                CellsChanged = changed,
                Skipped = skipped.Count > 0 ? skipped : null,
                Warnings = skipped.Count > 0 ? [$"{skipped.Count} cell(s) are not numbers in culture {cultureInfo.Name} and were left unchanged."] : null,
            };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult ConditionalFormat(IPresentationBatch batch, int slideIndex, int shapeIndex, int column, string rule, double? threshold = null, int count = 3,
        string? textColor = null, string? fillColor = null, bool? bold = null, string culture = "invariant", int firstRow = 2)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (rule is not ("negative" or "positive" or "above" or "below" or "equals" or "top" or "bottom"))
            return TableFail("rule must be negative, positive, above, below, equals, top, or bottom.");
        if (rule is "above" or "below" or "equals" && threshold is null)
            return TableFail($"rule {rule} needs threshold.");
        if (count < 1)
            return TableFail("count must be 1 or greater.");
        if (textColor is null && fillColor is null && bold is null)
            return TableFail("Pass text_color, fill_color, and/or bold.");
        foreach (var color in new[] { textColor, fillColor }.OfType<string>())
        {
            if (!ResolvedProfile.IsHex(color))
                return TableFail($"'{color}' is not #RRGGBB.");
        }
        CultureInfo cultureInfo;
        try
        {
            cultureInfo = DataLoader.Culture(culture);
        }
        catch (ArgumentException ex)
        {
            return TableFail(ex.Message);
        }

        return WithTable(batch, slideIndex, shapeIndex, (table, shape, rows, columns) =>
        {
            if (column < 1 || column > columns)
                return TableFail($"Column {column} is out of range (1-{columns}).");
            var values = new List<(int Row, double Value)>();
            var skipped = new List<string>();
            for (int row = Math.Max(1, firstRow); row <= rows; row++)
            {
                var text = DeckSnapshotReader.ReadCellText(table, row, column);
                if (string.IsNullOrWhiteSpace(text))
                    continue;
                if (DataTyping.TryParseNumber(text, cultureInfo, out var value, out _))
                    values.Add((row, value));
                else
                    skipped.Add($"R{row}C{column}: \"{text}\"");
            }
            var matches = rule switch
            {
                "negative" => values.Where(entry => entry.Value < 0),
                "positive" => values.Where(entry => entry.Value > 0),
                "above" => values.Where(entry => entry.Value > threshold),
                "below" => values.Where(entry => entry.Value < threshold),
                "equals" => values.Where(entry => Math.Abs(entry.Value - threshold!.Value) < 1e-9),
                "top" => values.OrderByDescending(entry => entry.Value).Take(count),
                _ => values.OrderBy(entry => entry.Value).Take(count),
            };
            int changed = 0;
            foreach (var (row, _) in matches.ToList())
            {
                StyleCell(table, row, column, fillColor, textColor, bold, null, null, null, null);
                changed++;
            }
            return new TableOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                CellsChanged = changed,
                Skipped = skipped.Count > 0 ? skipped : null,
            };
        });
    }

    /// <inheritdoc/>
    public TableOperationResult SetRangeStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, int firstRow, int lastRow, int firstColumn, int lastColumn,
        string? fillColor = null, string? textColor = null, bool? bold = null, float? fontSize = null, string? align = null, string? verticalAlign = null, float? padding = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        foreach (var color in new[] { textColor, fillColor }.OfType<string>())
        {
            if (!ResolvedProfile.IsHex(color))
                return TableFail($"'{color}' is not #RRGGBB.");
        }
        if (align is not null and not ("left" or "center" or "right"))
            return TableFail("align must be left, center, or right.");
        if (verticalAlign is not null and not ("top" or "middle" or "bottom"))
            return TableFail("vertical_align must be top, middle, or bottom.");
        if (fontSize is < 6 or > 72)
            return TableFail("font_size must be between 6 and 72 points.");
        if (padding is < 0 or > 72)
            return TableFail("padding must be between 0 and 72 points.");
        return WithTable(batch, slideIndex, shapeIndex, (table, shape, rows, columns) =>
        {
            if (firstRow < 1 || lastRow > rows || firstRow > lastRow || firstColumn < 1 || lastColumn > columns || firstColumn > lastColumn)
                return TableFail($"The block R{firstRow}C{firstColumn}:R{lastRow}C{lastColumn} is outside the {rows}x{columns} table.");
            int changed = 0;
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int column = firstColumn; column <= lastColumn; column++)
                {
                    StyleCell(table, row, column, fillColor, textColor, bold, fontSize, align, verticalAlign, padding);
                    changed++;
                }
            }
            return new TableOperationResult { Success = true, ShapeIndex = shapeIndex, CellsChanged = changed };
        });
    }

    private static void StyleCell(PowerPoint.Table table, int row, int column, string? fill, string? textColor, bool? bold, float? fontSize, string? align, string? verticalAlign, float? padding)
    {
        PowerPoint.Cell? cell = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.FillFormat? fillFormat = null;
        PowerPoint.ColorFormat? fillColor = null;
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        PowerPoint.Font? font = null;
        PowerPoint.ColorFormat? fontColor = null;
        PowerPoint.ParagraphFormat? paragraph = null;
        try
        {
            cell = table.Cell(row, column);
            shape = cell.Shape;
            if (fill is not null)
            {
                fillFormat = shape.Fill;
                fillFormat.Visible = Office.MsoTriState.msoTrue;
                fillFormat.Solid();
                fillColor = fillFormat.ForeColor;
                fillColor.RGB = ChartStyler.Bgr(fill);
            }
            frame = shape.TextFrame;
            if (padding is { } inset)
            {
                frame.MarginLeft = inset;
                frame.MarginRight = inset;
                frame.MarginTop = inset;
                frame.MarginBottom = inset;
            }
            if (verticalAlign is not null)
            {
                frame.VerticalAnchor = verticalAlign switch
                {
                    "middle" => Office.MsoVerticalAnchor.msoAnchorMiddle,
                    "bottom" => Office.MsoVerticalAnchor.msoAnchorBottom,
                    _ => Office.MsoVerticalAnchor.msoAnchorTop,
                };
            }
            range = frame.TextRange;
            font = range.Font;
            if (textColor is not null)
            {
                fontColor = font.Color;
                fontColor.RGB = ChartStyler.Bgr(textColor);
            }
            if (bold is { } isBold)
                font.Bold = isBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            if (fontSize is { } size)
                font.Size = size;
            if (align is not null)
            {
                paragraph = range.ParagraphFormat;
                paragraph.Alignment = align switch
                {
                    "center" => PowerPoint.PpParagraphAlignment.ppAlignCenter,
                    "right" => PowerPoint.PpParagraphAlignment.ppAlignRight,
                    _ => PowerPoint.PpParagraphAlignment.ppAlignLeft,
                };
            }
        }
        finally
        {
            if (paragraph is not null) ComUtilities.Release(ref paragraph);
            if (fontColor is not null) ComUtilities.Release(ref fontColor);
            if (font is not null) ComUtilities.Release(ref font);
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
            if (fillColor is not null) ComUtilities.Release(ref fillColor);
            if (fillFormat is not null) ComUtilities.Release(ref fillFormat);
            if (shape is not null) ComUtilities.Release(ref shape);
            if (cell is not null) ComUtilities.Release(ref cell);
        }
    }

    private static void WithCellRange(PowerPoint.Table table, int row, int column, Action<PowerPoint.TextRange> action)
    {
        PowerPoint.Cell? cell = null;
        PowerPoint.Shape? shape = null;
        try
        {
            cell = table.Cell(row, column);
            shape = cell.Shape;
            TextRuns.WithRange(shape, (_, range) => { action(range); return true; });
        }
        finally
        {
            if (shape is not null) ComUtilities.Release(ref shape);
            if (cell is not null) ComUtilities.Release(ref cell);
        }
    }

    /// <summary>Locates a table by slide and shape index and runs the operation with its size.</summary>
    private static TableOperationResult WithTable(IPresentationBatch batch, int slideIndex, int shapeIndex, Func<PowerPoint.Table, PowerPoint.Shape, int, int, TableOperationResult> operation) =>
        batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.Table? table = null;
            PowerPoint.Rows? rows = null;
            PowerPoint.Columns? columns = null;
            try
            {
                slides = ctx.Presentation.Slides;
                if (slideIndex < 1 || slideIndex > slides.Count)
                    return TableFail($"Slide index {slideIndex} is out of range. The presentation has {slides.Count} slide(s) (valid range: 1-{slides.Count}).");
                slide = slides[slideIndex];
                shapes = slide.Shapes;
                if (shapeIndex < 1 || shapeIndex > shapes.Count)
                    return TableFail($"Shape index {shapeIndex} is out of range. The slide has {shapes.Count} shape(s) (valid range: 1-{shapes.Count}).");
                shape = shapes[shapeIndex];
                if (shape.HasTable != Office.MsoTriState.msoTrue)
                    return TableFail($"Shape {shapeIndex} on slide {slideIndex} does not contain a table.");
                table = shape.Table;
                rows = table.Rows;
                columns = table.Columns;
                return operation(table, shape, rows.Count, columns.Count);
            }
            finally
            {
                if (columns is not null) ComUtilities.Release(ref columns);
                if (rows is not null) ComUtilities.Release(ref rows);
                if (table is not null) ComUtilities.Release(ref table);
                if (shape is not null) ComUtilities.Release(ref shape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });

    private static TableOperationResult TableFail(string message) => new() { Success = false, ErrorMessage = message };
}
