using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>
/// The regions of a slide derived from a design profile: safe area inside the margins, a title
/// area at the top, a footer area and an optional source line at the bottom, and the body
/// between them. All values are points, rounded to 0.1 pt so repeated layouts never drift.
/// </summary>
public sealed record SlideFrame(
    float Width,
    float Height,
    Box Safe,
    Box Title,
    Box Body,
    Box Source,
    Box Footer,
    int Columns,
    float Gutter,
    float Gap);

/// <summary>Deterministic box arithmetic for slide layout. Pure; points everywhere.</summary>
public static class LayoutEngine
{
    /// <summary>Computes the slide frame. Without a title the body starts at the top margin; without a source line the body extends to the footer.</summary>
    public static SlideFrame Frame(ResolvedProfile profile, float width, float height, bool hasTitle = true, bool hasSource = false)
    {
        ArgumentNullException.ThrowIfNull(profile);
        var left = profile.Margins["left"];
        var right = width - profile.Margins["right"];
        var top = profile.Margins["top"];
        var bottom = height - profile.Margins["bottom"];
        var safe = R(new Box(left, top, right, bottom));
        var gap = profile.Spacing("md");

        var footerHeight = profile.Areas["footer"];
        var footer = R(new Box(left, bottom - footerHeight, right, bottom));
        var sourceHeight = hasSource ? profile.Areas["source"] : 0f;
        var source = R(new Box(left, footer.Top - sourceHeight, right, footer.Top));
        var title = hasTitle ? R(new Box(left, top, right, top + profile.Areas["title"])) : R(new Box(left, top, right, top));
        var bodyTop = hasTitle ? title.Bottom + gap : top;
        var bodyBottom = (hasSource ? source.Top : footer.Top) - (hasSource ? profile.Spacing("sm") : gap / 2f);
        var body = R(new Box(left, bodyTop, right, Math.Max(bodyTop, bodyBottom)));
        return new SlideFrame(width, height, safe, title, body, source, footer, profile.GridColumns, profile.Gutter, gap);
    }

    /// <summary>Splits an area into columns, optionally weighted (weights are relative).</summary>
    public static IReadOnlyList<Box> Columns(Box area, int count, float gutter, IReadOnlyList<float>? weights = null)
    {
        if (count < 1)
            return [];
        var shares = Shares(count, weights);
        var available = area.Width - (gutter * (count - 1));
        var result = new List<Box>(count);
        var x = area.Left;
        for (int i = 0; i < count; i++)
        {
            var width = available * shares[i];
            var right = i == count - 1 ? area.Right : x + width;
            result.Add(R(new Box(x, area.Top, right, area.Bottom)));
            x = right + gutter;
        }
        return result;
    }

    /// <summary>Splits an area into rows, optionally weighted.</summary>
    public static IReadOnlyList<Box> Rows(Box area, int count, float gap, IReadOnlyList<float>? weights = null)
    {
        if (count < 1)
            return [];
        var shares = Shares(count, weights);
        var available = area.Height - (gap * (count - 1));
        var result = new List<Box>(count);
        var y = area.Top;
        for (int i = 0; i < count; i++)
        {
            var height = available * shares[i];
            var bottom = i == count - 1 ? area.Bottom : y + height;
            result.Add(R(new Box(area.Left, y, area.Right, bottom)));
            y = bottom + gap;
        }
        return result;
    }

    /// <summary>Stacks fixed heights from the top of an area with a gap. The result may extend past the area; compare with <see cref="StackHeight"/>.</summary>
    public static IReadOnlyList<Box> Stack(Box area, IReadOnlyList<float> heights, float gap)
    {
        ArgumentNullException.ThrowIfNull(heights);
        var result = new List<Box>(heights.Count);
        var y = area.Top;
        foreach (var height in heights)
        {
            result.Add(R(new Box(area.Left, y, area.Right, y + height)));
            y += height + gap;
        }
        return result;
    }

    /// <summary>Total height of a stack.</summary>
    public static float StackHeight(IReadOnlyList<float> heights, float gap) =>
        heights.Count == 0 ? 0 : heights.Sum() + (gap * (heights.Count - 1));

    /// <summary>A grid of equal cells, filled row by row.</summary>
    public static IReadOnlyList<Box> Grid(Box area, int items, int columns, float gutter, float rowGap)
    {
        if (items < 1 || columns < 1)
            return [];
        var rows = (items + columns - 1) / columns;
        var rowBoxes = Rows(area, rows, rowGap);
        var result = new List<Box>(items);
        for (int row = 0; row < rows; row++)
        {
            var inRow = Math.Min(columns, items - (row * columns));
            // Last row keeps the same cell width as the others, left-aligned.
            var cells = Columns(rowBoxes[row], columns, gutter);
            result.AddRange(cells.Take(inRow));
        }
        return result;
    }

    /// <summary>Chooses a column count (1..max) so each item is at least <paramref name="minItemWidth"/> wide and rows stay balanced.</summary>
    public static int AutoColumns(int items, float areaWidth, float gutter, float minItemWidth, int max = 4)
    {
        if (items <= 1)
            return 1;
        var fit = Math.Max(1, (int)((areaWidth + gutter) / (minItemWidth + gutter)));
        var columns = Math.Min(Math.Min(max, fit), items);
        // Prefer balanced rows: 4 items -> 2x2 rather than 3+1 when 3 columns would be chosen.
        var rows = (items + columns - 1) / columns;
        while (columns > 1 && (items + columns - 2) / (columns - 1) == rows)
            columns--;
        return columns;
    }

    /// <summary>Shrinks a box by equal padding.</summary>
    public static Box Inset(Box box, float padding) => Inset(box, padding, padding, padding, padding);

    /// <summary>Shrinks a box by per-side padding (never below zero size).</summary>
    public static Box Inset(Box box, float left, float top, float right, float bottom)
    {
        var l = box.Left + left;
        var t = box.Top + top;
        var r = Math.Max(l, box.Right - right);
        var b = Math.Max(t, box.Bottom - bottom);
        return R(new Box(l, t, r, b));
    }

    /// <summary>A box of the given size centered in another.</summary>
    public static Box Center(Box outer, float width, float height) =>
        R(Box.FromSize(outer.Left + ((outer.Width - width) / 2f), outer.Top + ((outer.Height - height) / 2f), width, height));

    /// <summary>Takes a slice from the top of a box.</summary>
    public static (Box Taken, Box Remainder) TakeTop(Box box, float height, float gap)
    {
        var taken = R(new Box(box.Left, box.Top, box.Right, Math.Min(box.Bottom, box.Top + height)));
        var rest = R(new Box(box.Left, Math.Min(box.Bottom, taken.Bottom + gap), box.Right, box.Bottom));
        return (taken, rest);
    }

    /// <summary>Takes a slice from the bottom of a box.</summary>
    public static (Box Taken, Box Remainder) TakeBottom(Box box, float height, float gap)
    {
        var taken = R(new Box(box.Left, Math.Max(box.Top, box.Bottom - height), box.Right, box.Bottom));
        var rest = R(new Box(box.Left, box.Top, box.Right, Math.Max(box.Top, taken.Top - gap)));
        return (taken, rest);
    }

    /// <summary>Spans grid columns <paramref name="first"/> (0-based) through first+span-1 of the frame's body.</summary>
    public static Box Span(SlideFrame frame, Box area, int first, int span)
    {
        ArgumentNullException.ThrowIfNull(frame);
        var columns = Columns(area, frame.Columns, frame.Gutter);
        var last = Math.Min(columns.Count - 1, first + span - 1);
        return R(new Box(columns[first].Left, area.Top, columns[last].Right, area.Bottom));
    }

    /// <summary>Rounds to 0.1 pt.</summary>
    public static Box R(Box box) => new(Round(box.Left), Round(box.Top), Round(box.Right), Round(box.Bottom));

    /// <summary>Rounds to 0.1 pt.</summary>
    public static float Round(float value) => MathF.Round(value * 10f, MidpointRounding.AwayFromZero) / 10f;

    private static float[] Shares(int count, IReadOnlyList<float>? weights)
    {
        if (weights is null || weights.Count != count || weights.Any(weight => weight <= 0))
            return Enumerable.Repeat(1f / count, count).ToArray();
        var total = weights.Sum();
        return weights.Select(weight => weight / total).ToArray();
    }
}
