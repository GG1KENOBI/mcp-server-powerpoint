using System.Text;

namespace Sbroenne.PowerPointMcp.Core.Table;

/// <summary>
/// Parses table rows written as separated cells, e.g. <c>"Region|Q1|Q2"</c> or a pasted
/// markdown row <c>"| Region | Q1 | Q2 |"</c>. Pure string handling, no COM.
/// </summary>
/// <remarks>
/// Cells are trimmed. One leading and one trailing separator are dropped (markdown style), and
/// markdown header dividers such as <c>|---|:--:|</c> are skipped. Write <c>\|</c> (backslash +
/// separator) for a literal separator inside a cell.
/// </remarks>
public static class TableDataParser
{
    /// <summary>Splits each row into trimmed cells; returns one list per data row.</summary>
    public static IReadOnlyList<IReadOnlyList<string>> Parse(IReadOnlyList<string> rows, string separator)
    {
        ArgumentNullException.ThrowIfNull(rows);
        if (string.IsNullOrEmpty(separator))
            throw new ArgumentException("separator must not be empty.", nameof(separator));

        var result = new List<IReadOnlyList<string>>(rows.Count);
        foreach (var row in rows)
        {
            var cells = Split(row ?? "", separator);
            if (IsMarkdownDivider(cells))
                continue;
            result.Add(cells);
        }
        return result;
    }

    private static List<string> Split(string row, string separator)
    {
        var trimmed = row.Trim();
        var cells = new List<string>();
        var cell = new StringBuilder();
        var escapedSeparator = "\\" + separator;
        for (var i = 0; i < trimmed.Length;)
        {
            if (string.CompareOrdinal(trimmed, i, escapedSeparator, 0, escapedSeparator.Length) == 0)
            {
                cell.Append(separator);
                i += escapedSeparator.Length;
            }
            else if (string.CompareOrdinal(trimmed, i, separator, 0, separator.Length) == 0)
            {
                cells.Add(cell.ToString().Trim());
                cell.Clear();
                i += separator.Length;
            }
            else
            {
                cell.Append(trimmed[i]);
                i++;
            }
        }
        cells.Add(cell.ToString().Trim());

        // Markdown rows start and end with the separator: drop the empty edge cells that creates.
        if (cells.Count > 1 && trimmed.StartsWith(separator, StringComparison.Ordinal))
            cells.RemoveAt(0);
        if (cells.Count > 1 && trimmed.EndsWith(separator, StringComparison.Ordinal) &&
            !trimmed.EndsWith(escapedSeparator, StringComparison.Ordinal))
        {
            cells.RemoveAt(cells.Count - 1);
        }
        return cells;
    }

    private static bool IsMarkdownDivider(List<string> cells) =>
        cells.Any(cell => cell.Contains('-', StringComparison.Ordinal)) &&
        cells.All(cell => cell.Length > 0 && cell.All(character => character is '-' or ':' || char.IsWhiteSpace(character)));
}
