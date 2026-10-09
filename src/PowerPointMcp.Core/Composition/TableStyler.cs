extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Chart;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>
/// Fills and styles a native table from a <see cref="TablePlan"/>: column widths, cell text exactly
/// as given, header/body/banded/total/highlight fills, alignment, padding, and bottom borders.
/// STA thread only.
/// </summary>
internal static class TableStyler
{
    // "No Style, No Grid": a clean base so the profile's fills and borders are the only styling.
    private const string NoStyleNoGrid = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";

    internal static void Fill(PowerPoint.Shape shape, TablePlan plan)
    {
        PowerPoint.Table? table = null;
        try
        {
            table = shape.Table;
            table.ApplyStyle(NoStyleNoGrid, false);
            bool hasHeader = plan.Header.Count > 0;
            int columns = hasHeader ? plan.Header.Count : plan.Rows[0].Count;
            var totalWidth = plan.ColumnWidths.Sum();
            for (int column = 1; column <= columns; column++)
            {
                PowerPoint.Column? item = null;
                PowerPoint.Columns? all = null;
                try
                {
                    all = table.Columns;
                    item = all[column];
                    item.Width = shape.Width * plan.ColumnWidths[column - 1] / totalWidth;
                }
                finally
                {
                    if (item is not null) ComUtilities.Release(ref item);
                    if (all is not null) ComUtilities.Release(ref all);
                }
            }

            int row = 1;
            if (hasHeader)
            {
                for (int column = 1; column <= columns; column++)
                    StyleCell(table, row, column, plan.Header[column - 1], plan, plan.Style.HeaderFill, plan.Style.HeaderText, bold: true, plan.Align[column - 1], isHeader: true);
                row++;
            }
            for (int body = 0; body < plan.Rows.Count; body++, row++)
            {
                bool isTotal = plan.TotalRow && body == plan.Rows.Count - 1;
                bool isHighlight = plan.HighlightRows.Contains(body + 1);
                bool isBand = plan.Style.BandFill is not null && (plan.FirstRowNumber + body) % 2 == 0;
                var fill = isTotal || isHighlight ? plan.Style.TotalFill : isBand ? plan.Style.BandFill! : plan.Style.BodyFill;
                for (int column = 1; column <= columns; column++)
                    StyleCell(table, row, column, plan.Rows[body][column - 1], plan, fill, plan.Style.BodyText, bold: isTotal || isHighlight, plan.Align[column - 1], isHeader: false);
            }
        }
        finally
        {
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static void StyleCell(PowerPoint.Table table, int row, int column, string text, TablePlan plan, string fill, string color, bool bold, string align, bool isHeader)
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
        PowerPoint.Borders? borders = null;
        PowerPoint.LineFormat? bottom = null;
        PowerPoint.ColorFormat? borderColor = null;
        try
        {
            cell = table.Cell(row, column);
            shape = cell.Shape;
            fillFormat = shape.Fill;
            fillFormat.Visible = Office.MsoTriState.msoTrue;
            fillFormat.Solid();
            fillColor = fillFormat.ForeColor;
            fillColor.RGB = ChartStyler.Bgr(fill);

            frame = shape.TextFrame;
            frame.MarginLeft = plan.Style.CellPadding + 2;
            frame.MarginRight = plan.Style.CellPadding + 2;
            frame.MarginTop = plan.Style.CellPadding;
            frame.MarginBottom = plan.Style.CellPadding;
            frame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
            range = frame.TextRange;
            range.Text = text.Replace("\r\n", "\v", StringComparison.Ordinal).Replace('\n', '\v');
            font = range.Font;
            font.Name = plan.Font;
            font.Size = plan.FontSize;
            font.Bold = bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            fontColor = font.Color;
            fontColor.RGB = ChartStyler.Bgr(color);
            paragraph = range.ParagraphFormat;
            paragraph.Alignment = align switch
            {
                "right" => PowerPoint.PpParagraphAlignment.ppAlignRight,
                "center" => PowerPoint.PpParagraphAlignment.ppAlignCenter,
                _ => PowerPoint.PpParagraphAlignment.ppAlignLeft,
            };

            borders = cell.Borders;
            bottom = borders[PowerPoint.PpBorderType.ppBorderBottom];
            if (plan.Style.BorderWidth > 0)
            {
                bottom.Visible = Office.MsoTriState.msoTrue;
                bottom.Weight = isHeader ? plan.Style.BorderWidth * 1.5f : plan.Style.BorderWidth;
                borderColor = bottom.ForeColor;
                borderColor.RGB = ChartStyler.Bgr(plan.Style.Border);
            }
            else
            {
                bottom.Visible = Office.MsoTriState.msoFalse;
            }
        }
        finally
        {
            if (borderColor is not null) ComUtilities.Release(ref borderColor);
            if (bottom is not null) ComUtilities.Release(ref bottom);
            if (borders is not null) ComUtilities.Release(ref borders);
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

    /// <summary>Counts body rows (after the header row, if any) whose bottom is at or above <paramref name="bottom"/>.</summary>
    internal static (int Fit, int Total) RowsThatFit(PowerPoint.Shape shape, float bottom, bool hasHeader)
    {
        PowerPoint.Table? table = null;
        PowerPoint.Rows? rows = null;
        try
        {
            table = shape.Table;
            rows = table.Rows;
            int count = rows.Count;
            float y = shape.Top;
            int fit = 0;
            for (int index = 1; index <= count; index++)
            {
                PowerPoint.Row? row = null;
                try
                {
                    row = rows[index];
                    y += row.Height;
                    if (index == 1 && hasHeader)
                        continue;
                    if (y > bottom + 1f)
                        break;
                    fit++;
                }
                finally
                {
                    if (row is not null) ComUtilities.Release(ref row);
                }
            }
            return (fit, hasHeader ? count - 1 : count);
        }
        finally
        {
            if (rows is not null) ComUtilities.Release(ref rows);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }
}
