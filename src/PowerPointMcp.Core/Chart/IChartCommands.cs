using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>
/// Chart lifecycle, data, and quick-formatting operations.
/// </summary>
[ServiceCategory("chart", "Chart")]
[McpTool("chart", Title = "Chart Operations", Destructive = true, Category = "content",
    Description = "Add a native chart shape, edit data, titles, legend, built-in style, color style, and data-table visibility in an open presentation session.")]
[McpReadOnlyActions("get-chart-data", "get-chart-title", "get-axis-title", "get-legend-visibility",
    "get-style", "get-color-style", "get-data-table", "get-details")]
public interface IChartCommands
{
    /// <summary>
    /// Adds a native chart shape (bar, line, or pie) to the given slide with categories and a
    /// single data series, and returns the new shape's index.
    /// </summary>
    /// <param name="batch">The active presentation batch.</param>
    /// <param name="slideIndex">1-based slide index.</param>
    /// <param name="chartType">Chart type: "bar", "line", or "pie".</param>
    /// <param name="left">Left position in points.</param>
    /// <param name="top">Top position in points.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    /// <param name="categories">Category labels (x-axis / pie slice labels).</param>
    /// <param name="seriesName">Name of the single data series.</param>
    /// <param name="values">Data values, one per category.</param>
    ChartOperationResult AddChart(IPresentationBatch batch, int slideIndex, string chartType, float left, float top, float width, float height, IReadOnlyList<string> categories, string seriesName, IReadOnlyList<double> values);

    /// <summary>
    /// Gets the category and series counts of an existing chart shape's data.
    /// </summary>
    ChartOperationResult GetChartData(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Adds another data series to an existing chart (created via <see cref="AddChart"/>),
    /// producing a multi-series chart. <paramref name="values"/> must have the same count as the
    /// chart's existing categories. Returns the new total <c>seriesCount</c>.
    /// </summary>
    ChartOperationResult AddSeries(IPresentationBatch batch, int slideIndex, int shapeIndex, string seriesName, IReadOnlyList<double> values);

    /// <summary>Sets the chart's main title text (and turns the title on).</summary>
    ChartOperationResult SetChartTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string title);

    /// <summary>Gets the chart's main title text and whether it currently has a title (<c>hasTitle</c>).</summary>
    ChartOperationResult GetChartTitle(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets an axis title (and turns that axis's title on). <paramref name="axisType"/> is
    /// <c>"category"</c> or <c>"value"</c>.
    /// </summary>
    ChartOperationResult SetAxisTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string axisType, string title);

    /// <summary>
    /// Gets an axis's title text and whether it currently has a title (<c>hasTitle</c>).
    /// <paramref name="axisType"/> is <c>"category"</c> or <c>"value"</c>.
    /// </summary>
    ChartOperationResult GetAxisTitle(IPresentationBatch batch, int slideIndex, int shapeIndex, string axisType);

    /// <summary>Shows or hides the chart's legend.</summary>
    /// <param name="visible">True to show the chart element; false to hide it.</param>
    ChartOperationResult SetLegendVisibility(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible);

    /// <summary>Gets whether the chart's legend is visible.</summary>
    ChartOperationResult GetLegendVisibility(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Replaces ALL of an existing chart's data (categories and every series) in one call — the
    /// chart's previous categories and series are discarded and replaced wholesale. Unlike
    /// <see cref="AddSeries"/> (which appends one series to the existing categories),
    /// <paramref name="categories"/> here can also change the category count/labels.
    /// <paramref name="seriesValues"/> is a single flat list laid out series-major: all values
    /// for <c>seriesNames[0]</c> (one per category, in category order), then all values for
    /// <c>seriesNames[1]</c>, and so on. Its length must equal
    /// <c>categories.Count * seriesNames.Count</c>.
    /// </summary>
    ChartOperationResult ReplaceChartData(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        IReadOnlyList<string> categories,
        IReadOnlyList<string> seriesNames,
        IReadOnlyList<double> seriesValues);

    /// <summary>Gets the chart's built-in visual style number.</summary>
    ChartOperationResult GetStyle(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets the chart's built-in visual style and returns the accepted value. Values outside the
    /// range verified against installed PowerPoint are rejected before COM.
    /// </summary>
    /// <param name="style">Built-in chart style number, verified against PowerPoint from 1 through 48.</param>
    ChartOperationResult SetStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, int style);

    /// <summary>Gets the chart's built-in color style number.</summary>
    ChartOperationResult GetColorStyle(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets the chart's built-in color style and returns the accepted value. Values outside the
    /// range verified against installed PowerPoint are rejected before COM.
    /// </summary>
    /// <param name="colorStyle">Built-in chart color style number, verified against PowerPoint from 1 through 26.</param>
    ChartOperationResult SetColorStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, int colorStyle);

    /// <summary>Gets whether the chart's data table is visible.</summary>
    ChartOperationResult GetDataTable(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Shows or hides the chart's data table and returns its resulting visibility.</summary>
    /// <param name="visible">True to show the chart element; false to hide it.</param>
    ChartOperationResult SetDataTable(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible);

    /// <summary>
    /// Replaces the chart's data through its embedded workbook (still editable with Edit Data):
    /// categories plus one list of values per series, series-major. Empty strings or "null" are
    /// missing values; anything that is not a number is rejected (nothing is converted). The
    /// chart, its type, and its formatting are kept.
    /// </summary>
    /// <param name="valueTexts">Values as text, series-major (categories x series), invariant culture ("1.5"); "" or "null" for missing.</param>
    /// <param name="numberFormat">Excel number format for the data and labels, e.g. 0.0 or 0%.</param>
    ChartOperationResult SetData(IPresentationBatch batch, int slideIndex, int shapeIndex, IReadOnlyList<string> categories, IReadOnlyList<string> seriesNames,
        IReadOnlyList<string> valueTexts, string? numberFormat = null);

    /// <summary>Sets axis bounds, unit, number format, title, log scale, and gridlines. Unset values are left as they are.</summary>
    /// <param name="axis">value (default), category, or secondary-value.</param>
    /// <param name="minimum">Axis minimum.</param>
    /// <param name="maximum">Axis maximum.</param>
    /// <param name="majorUnit">Distance between major ticks.</param>
    /// <param name="axisTitle">Axis title ("" removes it).</param>
    /// <param name="logScale">Logarithmic scale (value axes only).</param>
    /// <param name="gridlines">Show major gridlines.</param>
    /// <param name="autoScale">Reset minimum, maximum, and unit to automatic.</param>
    ChartOperationResult SetAxis(IPresentationBatch batch, int slideIndex, int shapeIndex, string axis = "value", double? minimum = null, double? maximum = null,
        double? majorUnit = null, string? numberFormat = null, [AllowEmptyString] string? axisTitle = null, bool? logScale = null, bool? gridlines = null, bool autoScale = false);

    /// <summary>Shows or hides data labels for all series or one series, with number format and position.</summary>
    /// <param name="visible">Show labels.</param>
    /// <param name="seriesIndex">1-based series (default: all series).</param>
    /// <param name="position">outside-end, inside-end, center, inside-base, above, below, left, right, or best-fit (depends on chart type).</param>
    /// <param name="showPercent">Show percentages (pie and doughnut).</param>
    ChartOperationResult SetDataLabels(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible, int? seriesIndex = null, string? numberFormat = null,
        string? position = null, bool? showPercent = null);

    /// <summary>
    /// Styles one series: color, its own chart type (combo charts, e.g. line over columns), the
    /// secondary value axis, line width, markers, and plot order.
    /// </summary>
    /// <param name="color">#RRGGBB.</param>
    /// <param name="seriesType">column, bar, line, or area for this series only.</param>
    /// <param name="secondaryAxis">Plot on the secondary value axis.</param>
    /// <param name="lineWidth">Line width in points (line series).</param>
    /// <param name="markers">Show markers (line series).</param>
    /// <param name="plotOrder">1-based position among the series.</param>
    ChartOperationResult SetSeriesStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, int seriesIndex, string? color = null, string? seriesType = null,
        bool? secondaryAxis = null, float? lineWidth = null, bool? markers = null, int? plotOrder = null);

    /// <summary>How empty cells plot: gap (default PowerPoint behaviour), zero, or connect (lines bridge the gap).</summary>
    /// <param name="mode">gap, zero, or connect.</param>
    ChartOperationResult SetMissingValues(IPresentationBatch batch, int slideIndex, int shapeIndex, string mode);

    /// <summary>Applies a design profile's chart style: series palette, font, legend, gridlines, data labels.</summary>
    /// <param name="profile">Design profile (default theme).</param>
    ChartOperationResult ApplyStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, string? profile = null);

    /// <summary>Returns chart type, categories, series (name, values, type, axis, color), axes, labels, and the missing-value mode.</summary>
    ChartOperationResult GetDetails(IPresentationBatch batch, int slideIndex, int shapeIndex);
}
