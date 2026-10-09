extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Run-level font size reading and proportional scaling, plus PowerPoint text measurement.
/// Scaling multiplies every run's own size by one factor, so heading/body hierarchy inside a
/// text frame is preserved. STA thread only.
/// </summary>
internal static class TextRuns
{
    /// <summary>Most runs scaled per shape; larger frames are reported instead of scaled.</summary>
    internal const int MaxRuns = 400;

    /// <summary>Reads every run's font size; null when the frame has more than <see cref="MaxRuns"/> runs.</summary>
    internal static List<float>? ReadSizes(PowerPoint.TextRange range)
    {
        PowerPoint.TextRange? runs = null;
        try
        {
            runs = range.Runs();
            if (runs.Count > MaxRuns)
                return null;
            var sizes = new List<float>(runs.Count);
            for (int index = 1; index <= runs.Count; index++)
            {
                PowerPoint.TextRange? run = null;
                PowerPoint.Font? font = null;
                try
                {
                    run = range.Runs(index, 1);
                    font = run.Font;
                    sizes.Add(font.Size);
                }
                finally
                {
                    if (font is not null) ComUtilities.Release(ref font);
                    if (run is not null) ComUtilities.Release(ref run);
                }
            }
            return sizes;
        }
        finally
        {
            if (runs is not null) ComUtilities.Release(ref runs);
        }
    }

    /// <summary>Sets every run to its original size times <paramref name="factor"/> (rounded to 0.1 pt).</summary>
    internal static void WriteSizes(PowerPoint.TextRange range, List<float> sizes, float factor)
    {
        for (int index = 1; index <= sizes.Count; index++)
        {
            if (sizes[index - 1] <= 0)
                continue;
            PowerPoint.TextRange? run = null;
            PowerPoint.Font? font = null;
            try
            {
                run = range.Runs(index, 1);
                font = run.Font;
                font.Size = MathF.Round(sizes[index - 1] * factor * 10f) / 10f;
            }
            finally
            {
                if (font is not null) ComUtilities.Release(ref font);
                if (run is not null) ComUtilities.Release(ref run);
            }
        }
    }

    /// <summary>
    /// Height the shape's text needs (measured text height plus top and bottom margins), or null
    /// when the shape has no text frame.
    /// </summary>
    internal static float? NeededHeight(PowerPoint.Shape shape)
    {
        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
            return null;
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            frame = shape.TextFrame;
            range = frame.TextRange;
            var text = range.Text;
            if (string.IsNullOrEmpty(text))
                return frame.MarginTop + frame.MarginBottom;
            return range.BoundHeight + frame.MarginTop + frame.MarginBottom;
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    /// <summary>Runs <paramref name="action"/> with the shape's text range.</summary>
    internal static T WithRange<T>(PowerPoint.Shape shape, Func<PowerPoint.TextFrame, PowerPoint.TextRange, T> action)
    {
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            frame = shape.TextFrame;
            range = frame.TextRange;
            return action(frame, range);
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }
}
