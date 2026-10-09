using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Composition;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>One chart series as read back from PowerPoint.</summary>
public sealed record ChartSeriesDetails(int Index, string Name, IReadOnlyList<double?> Values, int ChartType, bool SecondaryAxis, string? Color, int PlotOrder);

/// <summary>Chart details for get-details.</summary>
public sealed record ChartDetails(int ChartType, IReadOnlyList<string> Categories, IReadOnlyList<ChartSeriesDetails> Series, bool HasLegend, string MissingValues,
    IReadOnlyDictionary<string, string> Axes);

public sealed partial class ChartCommands
{
    private const int XlValue = 2;
    private const int XlCategory = 1;
    private const int XlPrimary = 1;
    private const int XlSecondary = 2;

    private static readonly Dictionary<string, int> SeriesTypes = new(StringComparer.Ordinal)
    {
        ["column"] = 51,
        ["bar"] = 57,
        ["line"] = 65,
        ["area"] = 1,
    };

    private static readonly Dictionary<string, int> LabelPositions = new(StringComparer.Ordinal)
    {
        ["outside-end"] = 2, // xlLabelPositionOutsideEnd
        ["inside-end"] = 3, // xlLabelPositionInsideEnd
        ["center"] = -4108, // xlLabelPositionCenter
        ["inside-base"] = 4, // xlLabelPositionInsideBase
        ["above"] = 0, // xlLabelPositionAbove
        ["below"] = 1, // xlLabelPositionBelow
        ["left"] = -4131, // xlLabelPositionLeft
        ["right"] = -4152, // xlLabelPositionRight
        ["best-fit"] = 5, // xlLabelPositionBestFit
    };

    /// <inheritdoc/>
    public ChartOperationResult SetData(IPresentationBatch batch, int slideIndex, int shapeIndex, IReadOnlyList<string> categories, IReadOnlyList<string> seriesNames,
        IReadOnlyList<string> valueTexts, string? numberFormat = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(categories);
        ArgumentNullException.ThrowIfNull(seriesNames);
        ArgumentNullException.ThrowIfNull(valueTexts);
        var values = valueTexts;
        if (categories.Count == 0 || seriesNames.Count == 0)
            return ChartFail("At least one category and one series are required.");
        if (values.Count != categories.Count * seriesNames.Count)
            return ChartFail($"value_texts has {values.Count} entries but categories x series_names needs {categories.Count * seriesNames.Count} (series-major).");
        var numbers = new double[values.Count];
        var bad = new List<string>();
        for (int i = 0; i < values.Count; i++)
        {
            var text = values[i]?.Trim() ?? "";
            if (text.Length == 0 || string.Equals(text, "null", StringComparison.OrdinalIgnoreCase))
                numbers[i] = double.NaN;
            else if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out numbers[i]))
                bad.Add($"#{i + 1} \"{text}\"");
        }
        if (bad.Count > 0)
            return ChartFail($"These values are not numbers (invariant culture, e.g. 1.5): {string.Join(", ", bad.Take(10))}. Nothing was changed.");

        var series = seriesNames.Select((name, s) => new ChartSeriesData(name, numbers.Skip(s * categories.Count).Take(categories.Count).ToList())).ToList();
        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            ChartDataWriter.Write(chart, categories, series, numberFormat);
            return new ChartOperationResult { Success = true, ShapeIndex = shapeIndex, CategoryCount = categories.Count, SeriesCount = seriesNames.Count };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetAxis(IPresentationBatch batch, int slideIndex, int shapeIndex, string axis = "value", double? minimum = null, double? maximum = null,
        double? majorUnit = null, string? numberFormat = null, string? axisTitle = null, bool? logScale = null, bool? gridlines = null, bool autoScale = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var (type, group) = axis switch
        {
            "value" => (XlValue, XlPrimary),
            "category" => (XlCategory, XlPrimary),
            "secondary-value" => (XlValue, XlSecondary),
            _ => (0, 0),
        };
        if (type == 0)
            return ChartFail("axis must be value, category, or secondary-value.");
        if (minimum is not null && maximum is not null && minimum >= maximum)
            return ChartFail("minimum must be less than maximum.");
        if (majorUnit is <= 0)
            return ChartFail("major_unit must be positive.");

        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            var notes = new List<string>();
            // PIA gap: Chart.Axes returns an Excel Axis surfaced as object.
            dynamic? target = null;
            try
            {
                try
                {
                    target = chart.Axes(type, (PowerPoint.XlAxisGroup)group);
                }
                catch (COMException)
                {
                    return ChartFail(axis == "secondary-value"
                        ? "The chart has no secondary value axis; put a series on it first with set-series-style secondary_axis=true."
                        : $"This chart type has no {axis} axis.");
                }
                if (autoScale)
                {
                    target.MinimumScaleIsAuto = true;
                    target.MaximumScaleIsAuto = true;
                    target.MajorUnitIsAuto = true;
                }
                if (type == XlCategory && (minimum is not null || maximum is not null || majorUnit is not null || logScale is not null))
                    notes.Add("minimum, maximum, major_unit, and log_scale apply to value axes; ignored for the category axis.");
                else
                {
                    if (minimum is { } min) target.MinimumScale = min;
                    if (maximum is { } max) target.MaximumScale = max;
                    if (majorUnit is { } unit) target.MajorUnit = unit;
                    if (logScale is { } log) target.ScaleType = log ? -4133 : -4132; // xlScaleLogarithmic / xlScaleLinear
                }
                if (numberFormat is not null)
                {
                    dynamic? labels = null;
                    try
                    {
                        labels = target.TickLabels;
                        labels.NumberFormat = numberFormat;
                    }
                    finally
                    {
                        if (labels != null) ComUtilities.Release(ref labels!);
                    }
                }
                if (axisTitle is not null)
                {
                    target.HasTitle = axisTitle.Length > 0;
                    if (axisTitle.Length > 0)
                    {
                        dynamic? title = null;
                        try
                        {
                            title = target.AxisTitle;
                            title.Text = axisTitle;
                        }
                        finally
                        {
                            if (title != null) ComUtilities.Release(ref title!);
                        }
                    }
                }
                if (gridlines is { } grid)
                    target.HasMajorGridlines = grid;
                return new ChartOperationResult { Success = true, ShapeIndex = shapeIndex, AxisType = axis, Warnings = notes.Count > 0 ? notes : null };
            }
            finally
            {
                if (target != null) ComUtilities.Release(ref target!);
            }
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetDataLabels(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible, int? seriesIndex = null, string? numberFormat = null,
        string? position = null, bool? showPercent = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (position is not null && !LabelPositions.ContainsKey(position))
            return ChartFail($"position must be one of {string.Join(", ", LabelPositions.Keys)}.");
        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            var notes = new List<string>();
            int changed = ForEachSeries(chart, seriesIndex, series =>
            {
                series.HasDataLabels = visible;
                if (!visible)
                    return;
                dynamic? labels = null;
                try
                {
                    labels = series.DataLabels();
                    if (numberFormat is not null)
                        labels.NumberFormat = numberFormat;
                    if (showPercent is { } percent)
                        labels.ShowPercentage = percent;
                    if (position is not null)
                    {
                        try
                        {
                            labels.Position = LabelPositions[position];
                        }
                        catch (COMException)
                        {
                            notes.Add($"Label position {position} is not available for this chart type; left unchanged.");
                        }
                    }
                }
                finally
                {
                    if (labels != null) ComUtilities.Release(ref labels!);
                }
            }, out var error);
            return error is not null
                ? ChartFail(error)
                : new ChartOperationResult { Success = true, ShapeIndex = shapeIndex, SeriesCount = changed, Warnings = notes.Distinct().ToList() is { Count: > 0 } list ? list : null };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetSeriesStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, int seriesIndex, string? color = null, string? seriesType = null,
        bool? secondaryAxis = null, float? lineWidth = null, bool? markers = null, int? plotOrder = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (color is not null && !Design.ResolvedProfile.IsHex(color))
            return ChartFail("color must be #RRGGBB.");
        if (seriesType is not null && !SeriesTypes.ContainsKey(seriesType))
            return ChartFail($"series_type must be one of {string.Join(", ", SeriesTypes.Keys)}.");
        if (lineWidth is <= 0 or > 20)
            return ChartFail("line_width must be between 0 and 20 points.");
        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            int count = ForEachSeries(chart, seriesIndex, series =>
            {
                if (seriesType is not null)
                    series.ChartType = SeriesTypes[seriesType];
                if (secondaryAxis is { } secondary)
                    series.AxisGroup = secondary ? XlSecondary : XlPrimary;
                if (color is not null || lineWidth is not null)
                {
                    dynamic? format = null;
                    dynamic? fill = null;
                    dynamic? fillColor = null;
                    dynamic? line = null;
                    dynamic? lineColor = null;
                    try
                    {
                        format = series.Format;
                        if (color is not null)
                        {
                            fill = format.Fill;
                            fill.Visible = -1;
                            fill.Solid();
                            fillColor = fill.ForeColor;
                            fillColor.RGB = ChartStyler.Bgr(color);
                        }
                        line = format.Line;
                        if (color is not null)
                        {
                            lineColor = line.ForeColor;
                            lineColor.RGB = ChartStyler.Bgr(color);
                        }
                        if (lineWidth is { } width)
                            line.Weight = width;
                    }
                    finally
                    {
                        if (lineColor != null) ComUtilities.Release(ref lineColor!);
                        if (line != null) ComUtilities.Release(ref line!);
                        if (fillColor != null) ComUtilities.Release(ref fillColor!);
                        if (fill != null) ComUtilities.Release(ref fill!);
                        if (format != null) ComUtilities.Release(ref format!);
                    }
                }
                if (markers is { } show)
                    series.MarkerStyle = show ? -4105 : -4142; // xlMarkerStyleAutomatic / xlMarkerStyleNone
                if (plotOrder is { } order)
                    series.PlotOrder = order;
            }, out var error);
            return error is not null ? ChartFail(error) : new ChartOperationResult { Success = true, ShapeIndex = shapeIndex, SeriesCount = count };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult SetMissingValues(IPresentationBatch batch, int slideIndex, int shapeIndex, string mode)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var value = mode switch
        {
            "gap" => PowerPoint.XlDisplayBlanksAs.xlNotPlotted,
            "zero" => PowerPoint.XlDisplayBlanksAs.xlZero,
            "connect" => PowerPoint.XlDisplayBlanksAs.xlInterpolated,
            _ => (PowerPoint.XlDisplayBlanksAs?)null,
        };
        if (value is null)
            return ChartFail("mode must be gap, zero, or connect.");
        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            chart.DisplayBlanksAs = value.Value;
            return new ChartOperationResult { Success = true, ShapeIndex = shapeIndex };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult ApplyStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, string? profile = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var (resolved, error, _) = ProfileResolver.Resolve(batch, profile);
        if (error is not null)
            return ChartFail(error);
        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            bool isPie = (int)chart.ChartType is 5 or -4120 or 69 or 70;
            var notes = ChartStyler.Apply(chart, new ChartLook(resolved!.Chart.Palette, resolved.BodyFont, resolved.Chart.FontSize, resolved.Color("text"),
                resolved.Chart.Gridlines, resolved.Chart.DataLabels, resolved.Chart.Legend, null, null, null), isPie);
            return new ChartOperationResult { Success = true, ShapeIndex = shapeIndex, Warnings = notes.Count > 0 ? notes : null };
        });
    }

    /// <inheritdoc/>
    public ChartOperationResult GetDetails(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return ExecuteWithChart(batch, slideIndex, shapeIndex, chart =>
        {
            var series = new List<ChartSeriesDetails>();
            var categories = new List<string>();
            ForEachSeries(chart, null, item =>
            {
                var values = ((Array)RetryTransientChartRead(() => (Array)item.Values)).Cast<object?>()
                    .Select(value => value is null ? (double?)null : Convert.ToDouble(value, CultureInfo.InvariantCulture)).ToList();
                if (categories.Count == 0)
                {
                    categories.AddRange(((Array)ReadNonEmptyXValues(item)).Cast<object?>().Select(value => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""));
                }
                string? color = null;
                dynamic? format = null;
                dynamic? fill = null;
                dynamic? fore = null;
                try
                {
                    format = item.Format;
                    fill = format.Fill;
                    fore = fill.ForeColor;
                    color = Deck.DeckSnapshotReader.Hex((int)fore.RGB);
                }
                catch (COMException)
                {
                    // Some series types (e.g. lines without fill) report no fill color.
                }
                finally
                {
                    if (fore != null) ComUtilities.Release(ref fore!);
                    if (fill != null) ComUtilities.Release(ref fill!);
                    if (format != null) ComUtilities.Release(ref format!);
                }
                series.Add(new ChartSeriesDetails(series.Count + 1, (string)item.Name, values, (int)item.ChartType, (int)item.AxisGroup == XlSecondary, color, (int)item.PlotOrder));
            }, out _);

            var axes = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var (name, type, group) in new[] { ("value", XlValue, XlPrimary), ("category", XlCategory, XlPrimary), ("secondary-value", XlValue, XlSecondary) })
            {
                // PIA gap: Chart.Axes returns an Excel Axis surfaced as object.
                dynamic? axis = null;
                try
                {
                    axis = chart.Axes(type, (PowerPoint.XlAxisGroup)group);
                    if (type == XlValue)
                        axes[name] = $"min={Describe(axis.MinimumScale)}{((bool)axis.MinimumScaleIsAuto ? " (auto)" : "")} max={Describe(axis.MaximumScale)}{((bool)axis.MaximumScaleIsAuto ? " (auto)" : "")} unit={Describe(axis.MajorUnit)} gridlines={(bool)axis.HasMajorGridlines}";
                    else
                        axes[name] = $"title={((bool)axis.HasTitle ? AxisTitleText(axis) : "none")}";
                }
                catch (COMException)
                {
                    // The axis does not exist for this chart type.
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // The axis does not expose this property for this chart type.
                }
                finally
                {
                    if (axis != null) ComUtilities.Release(ref axis!);
                }
            }

            var missing = chart.DisplayBlanksAs switch
            {
                PowerPoint.XlDisplayBlanksAs.xlZero => "zero",
                PowerPoint.XlDisplayBlanksAs.xlInterpolated => "connect",
                _ => "gap",
            };
            return new ChartOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                SeriesCount = series.Count,
                CategoryCount = categories.Count,
                Details = new ChartDetails((int)chart.ChartType, categories, series, chart.HasLegend, missing, axes),
            };
        });
    }

    private static string AxisTitleText(dynamic axis)
    {
        dynamic? title = null;
        try
        {
            title = axis.AxisTitle;
            return (string)title.Text;
        }
        finally
        {
            if (title != null) ComUtilities.Release(ref title!);
        }
    }

    private static string Describe(object value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";

    /// <summary>Runs an action on one series (1-based) or all; returns how many were visited.</summary>
    private static int ForEachSeries(PowerPoint.Chart chart, int? seriesIndex, Action<dynamic> action, out string? error)
    {
        error = null;
        // PIA gap: SeriesCollection and Series are Excel chart types surfaced as object.
        dynamic? collection = null;
        int visited = 0;
        try
        {
            collection = RetryTransientChartRead(() => chart.SeriesCollection());
            int count = (int)collection.Count;
            if (seriesIndex is { } only && (only < 1 || only > count))
            {
                error = $"Series {only} is out of range. The chart has {count} series (valid range: 1-{count}).";
                return 0;
            }
            for (int index = 1; index <= count; index++)
            {
                if (seriesIndex is not null && index != seriesIndex)
                    continue;
                dynamic? series = null;
                try
                {
                    series = collection.Item(index);
                    action(series);
                    visited++;
                }
                finally
                {
                    if (series != null) ComUtilities.Release(ref series!);
                }
            }
            return visited;
        }
        finally
        {
            if (collection != null) ComUtilities.Release(ref collection!);
        }
    }

    private static ChartOperationResult ChartFail(string message) => new() { Success = false, ErrorMessage = message };
}
