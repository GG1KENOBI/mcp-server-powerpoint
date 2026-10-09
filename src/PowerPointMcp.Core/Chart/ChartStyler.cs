using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Chart;

/// <summary>Chart appearance settings applied by <see cref="ChartStyler"/>.</summary>
internal sealed record ChartLook(
    IReadOnlyList<string> Palette,
    string FontName,
    float FontSize,
    string TextColor,
    bool Gridlines,
    bool DataLabels,
    string Legend,
    string? NumberFormat,
    string? ValueAxisTitle,
    string? CategoryAxisTitle);

/// <summary>
/// Applies consistent styling to a chart: series colors from a palette, fonts, legend position,
/// gridlines, data labels with number format, and axis titles. Returns what could not be
/// applied instead of throwing, because some chart types have no axes or gridlines.
/// STA thread only.
/// </summary>
internal static class ChartStyler
{
    private const int XlCategory = 1;
    private const int XlValue = 2;

    internal static List<string> Apply(PowerPoint.Chart chart, ChartLook look, bool isPie)
    {
        var notes = new List<string>();
        Try(notes, "title", () => chart.HasTitle = false);
        Try(notes, "legend", () =>
        {
            chart.HasLegend = look.Legend != "none";
            if (look.Legend != "none")
            {
                PowerPoint.Legend? legend = null;
                try
                {
                    legend = chart.Legend;
                    legend.Position = look.Legend switch
                    {
                        "right" => PowerPoint.XlLegendPosition.xlLegendPositionRight,
                        "top" => PowerPoint.XlLegendPosition.xlLegendPositionTop,
                        "left" => PowerPoint.XlLegendPosition.xlLegendPositionLeft,
                        _ => PowerPoint.XlLegendPosition.xlLegendPositionBottom,
                    };
                }
                finally
                {
                    if (legend is not null) ComUtilities.Release(ref legend);
                }
            }
        });
        Try(notes, "font", () => SetChartFont(chart, look));
        Try(notes, "series colors and labels", () => StyleSeries(chart, look, isPie));
        if (!isPie)
        {
            Try(notes, "gridlines", () => SetGridlines(chart, look.Gridlines));
            if (look.ValueAxisTitle is not null)
                Try(notes, "value axis title", () => SetAxisTitle(chart, XlValue, look.ValueAxisTitle));
            if (look.CategoryAxisTitle is not null)
                Try(notes, "category axis title", () => SetAxisTitle(chart, XlCategory, look.CategoryAxisTitle));
            if (look.NumberFormat is not null)
                Try(notes, "value axis number format", () => SetAxisNumberFormat(chart, look.NumberFormat));
        }
        return notes;
    }

    private static void SetChartFont(PowerPoint.Chart chart, ChartLook look)
    {
        // PIA gap: ChartArea.Format.TextFrame2 is reached late-bound because ChartFormat.TextFrame2
        // is not exposed on the embedded PowerPoint PIA's ChartArea type.
        dynamic? area = null;
        dynamic? format = null;
        dynamic? frame = null;
        dynamic? range = null;
        dynamic? font = null;
        dynamic? fill = null;
        dynamic? color = null;
        try
        {
            area = chart.ChartArea;
            format = area.Format;
            frame = format.TextFrame2;
            range = frame.TextRange;
            font = range.Font;
            font.Name = look.FontName;
            font.Size = look.FontSize;
            fill = font.Fill;
            color = fill.ForeColor;
            color.RGB = Bgr(look.TextColor);
        }
        finally
        {
            if (color != null) ComUtilities.Release(ref color!);
            if (fill != null) ComUtilities.Release(ref fill!);
            if (font != null) ComUtilities.Release(ref font!);
            if (range != null) ComUtilities.Release(ref range!);
            if (frame != null) ComUtilities.Release(ref frame!);
            if (format != null) ComUtilities.Release(ref format!);
            if (area != null) ComUtilities.Release(ref area!);
        }
    }

    private static void StyleSeries(PowerPoint.Chart chart, ChartLook look, bool isPie)
    {
        // PIA gap: SeriesCollection/Series/Points are Excel chart types surfaced as object.
        dynamic? collection = null;
        try
        {
            collection = ChartCommands.RetryTransientChartRead(() => chart.SeriesCollection());
            int count = (int)collection.Count;
            for (int index = 1; index <= count; index++)
            {
                dynamic? series = null;
                try
                {
                    series = collection.Item(index);
                    if (isPie)
                        ColorPoints(series, look.Palette);
                    else
                        ColorSeries(series, look.Palette[(index - 1) % look.Palette.Count]);

                    series.HasDataLabels = look.DataLabels;
                    if (look.DataLabels)
                        StyleLabels(series, look, isPie);
                }
                finally
                {
                    if (series != null) ComUtilities.Release(ref series!);
                }
            }
        }
        finally
        {
            if (collection != null) ComUtilities.Release(ref collection!);
        }
    }

    private static void ColorSeries(dynamic series, string hex)
    {
        dynamic? format = null;
        dynamic? fill = null;
        dynamic? color = null;
        dynamic? line = null;
        dynamic? lineColor = null;
        try
        {
            format = series.Format;
            fill = format.Fill;
            fill.Visible = -1; // msoTrue
            fill.Solid();
            color = fill.ForeColor;
            color.RGB = Bgr(hex);
            line = format.Line;
            lineColor = line.ForeColor;
            lineColor.RGB = Bgr(hex);
        }
        finally
        {
            if (lineColor != null) ComUtilities.Release(ref lineColor!);
            if (line != null) ComUtilities.Release(ref line!);
            if (color != null) ComUtilities.Release(ref color!);
            if (fill != null) ComUtilities.Release(ref fill!);
            if (format != null) ComUtilities.Release(ref format!);
        }
    }

    private static void ColorPoints(dynamic series, IReadOnlyList<string> palette)
    {
        dynamic? points = null;
        try
        {
            points = series.Points();
            int count = (int)points.Count;
            for (int index = 1; index <= count; index++)
            {
                dynamic? point = null;
                dynamic? format = null;
                dynamic? fill = null;
                dynamic? color = null;
                try
                {
                    point = points.Item(index);
                    format = point.Format;
                    fill = format.Fill;
                    fill.Solid();
                    color = fill.ForeColor;
                    color.RGB = Bgr(palette[(index - 1) % palette.Count]);
                }
                finally
                {
                    if (color != null) ComUtilities.Release(ref color!);
                    if (fill != null) ComUtilities.Release(ref fill!);
                    if (format != null) ComUtilities.Release(ref format!);
                    if (point != null) ComUtilities.Release(ref point!);
                }
            }
        }
        finally
        {
            if (points != null) ComUtilities.Release(ref points!);
        }
    }

    private static void StyleLabels(dynamic series, ChartLook look, bool isPie)
    {
        dynamic? labels = null;
        try
        {
            labels = series.DataLabels();
            if (look.NumberFormat is not null)
            {
                labels.NumberFormat = look.NumberFormat;
            }
            if (isPie)
            {
                labels.ShowPercentage = true;
                labels.ShowValue = look.NumberFormat is not null;
            }
        }
        finally
        {
            if (labels != null) ComUtilities.Release(ref labels!);
        }
    }

    private static void SetGridlines(PowerPoint.Chart chart, bool visible)
    {
        // PIA gap: Chart.Axes returns an Excel Axis surfaced as object.
        dynamic? axis = null;
        try
        {
            axis = chart.Axes(XlValue);
            axis.HasMajorGridlines = visible;
            axis.HasMinorGridlines = false;
        }
        finally
        {
            if (axis != null) ComUtilities.Release(ref axis!);
        }
    }

    private static void SetAxisTitle(PowerPoint.Chart chart, int axisType, string title)
    {
        // PIA gap: Chart.Axes returns an Excel Axis surfaced as object.
        dynamic? axis = null;
        dynamic? axisTitle = null;
        try
        {
            axis = chart.Axes(axisType);
            axis.HasTitle = true;
            axisTitle = axis.AxisTitle;
            axisTitle.Text = title;
        }
        finally
        {
            if (axisTitle != null) ComUtilities.Release(ref axisTitle!);
            if (axis != null) ComUtilities.Release(ref axis!);
        }
    }

    private static void SetAxisNumberFormat(PowerPoint.Chart chart, string numberFormat)
    {
        // PIA gap: Chart.Axes returns an Excel Axis surfaced as object.
        dynamic? axis = null;
        dynamic? labels = null;
        try
        {
            axis = chart.Axes(XlValue);
            labels = axis.TickLabels;
            labels.NumberFormat = numberFormat;
        }
        finally
        {
            if (labels != null) ComUtilities.Release(ref labels!);
            if (axis != null) ComUtilities.Release(ref axis!);
        }
    }

    private static void Try(List<string> notes, string what, Action action)
    {
        try
        {
            action();
        }
        catch (COMException ex)
        {
            notes.Add($"Chart {what} not applied: {ex.Message}");
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
        {
            notes.Add($"Chart {what} not applied: {ex.Message}");
        }
    }

    /// <summary>Converts #RRGGBB to the COM BGR integer.</summary>
    internal static int Bgr(string hex)
    {
        var value = int.Parse(hex.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return ((value & 0xFF) << 16) | (value & 0xFF00) | ((value >> 16) & 0xFF);
    }
}
