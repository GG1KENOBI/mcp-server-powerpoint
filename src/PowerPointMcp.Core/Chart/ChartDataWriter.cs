using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>One chart series: name and one value per category (NaN = missing value, left empty).</summary>
public sealed record ChartSeriesData(string Name, IReadOnlyList<double> Values);

/// <summary>
/// Writes chart data into the chart's embedded workbook and points the chart at that range, so
/// the data stays editable in PowerPoint (Edit Data) and later refreshes update the same cells.
/// STA thread only (inside IPresentationBatch.Execute).
/// </summary>
internal static class ChartDataWriter
{
    // XlRowCol.xlColumns: each series is a worksheet column.
    private const int PlotByColumns = 2;

    /// <summary>Chart type names compositions and data commands accept, mapped to XlChartType values.</summary>
    internal static readonly IReadOnlyDictionary<string, int> ChartTypes = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        ["column"] = 51, // xlColumnClustered
        ["stacked-column"] = 52, // xlColumnStacked
        ["bar"] = 57, // xlBarClustered
        ["stacked-bar"] = 58, // xlBarStacked
        ["line"] = 65, // xlLineMarkers
        ["pie"] = 5, // xlPie
        ["doughnut"] = -4120, // xlDoughnut
        ["area"] = 1, // xlArea
    };

    /// <summary>
    /// Replaces the worksheet contents with a category column and one column per series, applies
    /// an optional Excel number format to the values, resizes the data table, and calls
    /// SetSourceData. Returns the A1 range written (e.g. 'Sheet1'!$A$1:$C$5).
    /// </summary>
    internal static string Write(
        PowerPoint.Chart chart,
        IReadOnlyList<string> categories,
        IReadOnlyList<ChartSeriesData> series,
        string? numberFormat)
    {
        ArgumentNullException.ThrowIfNull(chart);
        int rows = categories.Count + 1;
        int columns = series.Count + 1;
        var cells = new object?[rows, columns];
        cells[0, 0] = "";
        for (int s = 0; s < series.Count; s++)
            cells[0, s + 1] = series[s].Name;
        for (int c = 0; c < categories.Count; c++)
        {
            cells[c + 1, 0] = categories[c];
            for (int s = 0; s < series.Count; s++)
            {
                var value = series[s].Values[c];
                cells[c + 1, s + 1] = double.IsNaN(value) ? null : value;
            }
        }

        string reference = "";
        ChartCommands.WithChartWorkbook(chart, workbook =>
        {
            // PIA gap: the chart data workbook is an Excel object; Excel interop types are not referenced.
            dynamic? worksheets = null;
            dynamic? sheet = null;
            dynamic? used = null;
            dynamic? range = null;
            dynamic? values = null;
            dynamic? tables = null;
            dynamic? table = null;
            try
            {
                worksheets = workbook.Worksheets;
                sheet = worksheets.Item(1);
                used = sheet.UsedRange;
                used.ClearContents();
                range = sheet.Range("A1").Resize(rows, columns);
                range.Value2 = cells;
                if (!string.IsNullOrWhiteSpace(numberFormat) && categories.Count > 0)
                {
                    values = sheet.Range("B2").Resize(categories.Count, series.Count);
                    values.NumberFormat = numberFormat;
                }

                // The default chart workbook defines its data as an Excel table; keep it in step.
                tables = sheet.ListObjects;
                if ((int)tables.Count >= 1)
                {
                    table = tables.Item(1);
                    table.Resize(range);
                }

                string sheetName = sheet.Name;
                reference = $"='{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!$A$1:${ColumnName(columns)}${rows.ToString(CultureInfo.InvariantCulture)}";
                return true;
            }
            finally
            {
                if (table != null) ComUtilities.Release(ref table!);
                if (tables != null) ComUtilities.Release(ref tables!);
                if (values != null) ComUtilities.Release(ref values!);
                if (range != null) ComUtilities.Release(ref range!);
                if (used != null) ComUtilities.Release(ref used!);
                if (sheet != null) ComUtilities.Release(ref sheet!);
                if (worksheets != null) ComUtilities.Release(ref worksheets!);
            }
        });

        chart.SetSourceData(reference, PlotByColumns);
        return reference;
    }

    /// <summary>Excel column letters for a 1-based column number.</summary>
    internal static string ColumnName(int column)
    {
        var name = "";
        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            name = (char)('A' + remainder) + name;
            column = (column - 1) / 26;
        }
        return name;
    }
}
