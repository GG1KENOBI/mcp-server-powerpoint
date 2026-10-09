using System.Globalization;
using System.Text;
using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Preview;

/// <summary>
/// A text wireframe of a slide for models that cannot see images: object outlines on a character
/// grid, labelled A, B, C… with a legend. Geometry only; it is not a rendering. Pure.
/// </summary>
public static class LayoutSketch
{
    /// <summary>Draws the sketch. Backdrops and hidden objects are listed but not outlined.</summary>
    public static string Render(float slideWidth, float slideHeight, IReadOnlyList<DeckObjectInfo> objects, int columns = 64)
    {
        ArgumentNullException.ThrowIfNull(objects);
        if (slideWidth <= 0 || slideHeight <= 0)
            return "";
        // Characters are roughly twice as tall as wide.
        int rows = Math.Max(8, (int)Math.Round(columns * (slideHeight / slideWidth) / 2.0));
        var grid = new char[rows, columns];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < columns; c++)
                grid[r, c] = ' ';

        var legend = new StringBuilder();
        var drawn = objects.Where(item => item.ShapeIndex is not null).OrderBy(item => item.ZOrder).ToList();
        int label = 0;
        foreach (var item in drawn)
        {
            var name = Label(label++);
            var box = DeckGeometry.BoundsOf(item);
            bool backdrop = box.Width * box.Height >= 0.9f * slideWidth * slideHeight;
            var text = item.Text is { Length: > 0 } value ? $" \"{Shorten(value, 50)}\"" : "";
            var role = item.Role is null ? "" : $" role={item.Role}";
            legend.Append(CultureInfo.InvariantCulture, $"{name} = {item.Kind}{role} id={item.ShapeId} at ({P(box.Left)},{P(box.Top)}) {P(box.Width)}x{P(box.Height)}{(item.Visible ? "" : " hidden")}{(backdrop ? " backdrop" : "")}{text}\n");
            if (!item.Visible || backdrop)
                continue;
            int c0 = Clamp((int)Math.Floor(box.Left / slideWidth * columns), columns);
            int c1 = Clamp((int)Math.Ceiling(box.Right / slideWidth * columns) - 1, columns);
            int r0 = Clamp((int)Math.Floor(box.Top / slideHeight * rows), rows);
            int r1 = Clamp((int)Math.Ceiling(box.Bottom / slideHeight * rows) - 1, rows);
            if (c1 < c0 || r1 < r0)
                continue;
            for (int c = c0; c <= c1; c++)
            {
                grid[r0, c] = grid[r0, c] == ' ' ? '-' : '+';
                grid[r1, c] = grid[r1, c] == ' ' ? '-' : '+';
            }
            for (int r = r0; r <= r1; r++)
            {
                grid[r, c0] = grid[r, c0] == ' ' ? '|' : '+';
                grid[r, c1] = grid[r, c1] == ' ' ? '|' : '+';
            }
            if (c0 + 1 <= c1 && r0 + 1 <= r1 || c0 == c1 || r0 == r1)
            {
                var labelRow = r0 + 1 <= r1 ? r0 + 1 : r0;
                var labelColumn = c0 + 1 <= c1 ? c0 + 1 : c0;
                for (int i = 0; i < name.Length && labelColumn + i < columns; i++)
                    grid[labelRow, labelColumn + i] = name[i];
            }
        }

        var sketch = new StringBuilder();
        sketch.Append('+').Append('=', columns).Append("+\n");
        for (int r = 0; r < rows; r++)
        {
            sketch.Append('|');
            for (int c = 0; c < columns; c++)
                sketch.Append(grid[r, c]);
            sketch.Append("|\n");
        }
        sketch.Append('+').Append('=', columns).Append("+\n");
        sketch.Append(CultureInfo.InvariantCulture, $"Slide {P(slideWidth)} x {P(slideHeight)} pt; one column = {P(slideWidth / columns)} pt, one row = {P(slideHeight / rows)} pt.\n");
        sketch.Append(legend);
        return sketch.ToString();
    }

    private static string Label(int index) =>
        index < 26 ? ((char)('A' + index)).ToString() : $"{(char)('A' + (index / 26) - 1)}{(char)('A' + (index % 26))}";

    private static int Clamp(int value, int size) => Math.Clamp(value, 0, size - 1);

    private static string P(float value) => value.ToString("0", CultureInfo.InvariantCulture);

    private static string Shorten(string value, int length)
    {
        var flat = value.Replace('\r', ' ').Replace('\n', ' ').Replace('\v', ' ');
        return flat.Length <= length ? flat : flat[..length] + "…";
    }
}
