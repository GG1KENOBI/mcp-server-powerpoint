using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>A typed cell: kind (number, percent, date, bool, text, empty, error), numeric value when numeric, display text, and the original text.</summary>
public sealed record DataValue(string Kind, double? Number, string Display, string Raw);

/// <summary>A column with its inferred type. Mixed columns are text, with notes naming the cells that disagree.</summary>
public sealed record DataColumn(int Index, string Name, string Type, string? Format, IReadOnlyList<string> Notes);

/// <summary>Tabular data from CSV or XLSX with inferred column types.</summary>
public sealed class DataTableContent
{
    /// <summary>Columns.</summary>
    public required IReadOnlyList<DataColumn> Columns { get; init; }

    /// <summary>Body rows (header excluded).</summary>
    public required IReadOnlyList<IReadOnlyList<DataValue>> Rows { get; init; }

    /// <summary>Notes on decoding and typing.</summary>
    public required IReadOnlyList<string> Notes { get; init; }

    /// <summary>Index of a column by name (case-insensitive) or 1-based number; -1 when not found.</summary>
    public int ColumnIndex(string nameOrNumber)
    {
        ArgumentNullException.ThrowIfNull(nameOrNumber);
        if (int.TryParse(nameOrNumber, NumberStyles.None, CultureInfo.InvariantCulture, out var number) && number >= 1 && number <= Columns.Count)
            return number - 1;
        return Columns.FirstOrDefault(column => string.Equals(column.Name, nameOrNumber, StringComparison.OrdinalIgnoreCase))?.Index ?? -1;
    }
}

/// <summary>
/// Types raw cells: culture-aware numbers ("1 234,5" in ru-RU, "1,234.5" in en-US), percentages,
/// accounting negatives "(12)", currency symbols, ISO and culture dates, booleans. A value is never
/// coerced silently: a column is numeric only when every non-empty cell parses, otherwise it is
/// text and the disagreeing cells are listed. Pure.
/// </summary>
public static partial class DataTyping
{
    /// <summary>Builds a table from CSV rows.</summary>
    public static DataTableContent FromRows(IReadOnlyList<IReadOnlyList<string>> rows, bool header, CultureInfo culture, IReadOnlyList<string>? notes = null)
    {
        ArgumentNullException.ThrowIfNull(rows);
        var width = rows.Count == 0 ? 0 : rows.Max(row => row.Count);
        var names = header && rows.Count > 0
            ? Enumerable.Range(0, width).Select(i => i < rows[0].Count && !string.IsNullOrWhiteSpace(rows[0][i]) ? rows[0][i].Trim() : $"Column {i + 1}").ToList()
            : Enumerable.Range(0, width).Select(i => $"Column {i + 1}").ToList();
        var body = rows.Skip(header ? 1 : 0).Select(row => Enumerable.Range(0, width).Select(i => Classify(i < row.Count ? row[i] : "", culture)).ToList()).ToList();
        return Finish(names, body, notes ?? [], formats: null);
    }

    /// <summary>Builds a table from an XLSX range (cells already typed by the workbook).</summary>
    public static DataTableContent FromXlsx(XlsxRange range, bool header, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(range);
        var width = range.Rows.Count == 0 ? 0 : range.Rows[0].Count;
        var names = header && range.Rows.Count > 0
            ? range.Rows[0].Select((cell, i) => string.IsNullOrWhiteSpace(Display(cell, culture)) ? $"Column {i + 1}" : Display(cell, culture).Trim()).ToList()
            : Enumerable.Range(0, width).Select(i => $"Column {i + 1}").ToList();
        var formats = new string?[width];
        var body = range.Rows.Skip(header ? 1 : 0).Select(row => row.Select((cell, i) =>
        {
            formats[i] ??= cell.Format;
            return cell.Kind switch
            {
                "number" or "percent" => new DataValue(cell.Kind, cell.Number, Display(cell, culture), cell.Number!.Value.ToString("R", CultureInfo.InvariantCulture)),
                "date" => new DataValue("date", cell.Number, cell.Text!, cell.Text!),
                "bool" => new DataValue("bool", cell.Number, cell.Text!, cell.Text!),
                "error" => new DataValue("error", null, cell.Text ?? "#ERROR", cell.Text ?? ""),
                "text" => Classify(cell.Text ?? "", culture) is { Kind: "text" } classified ? classified : new DataValue("text", null, cell.Text ?? "", cell.Text ?? ""),
                _ => new DataValue("empty", null, "", ""),
            };
        }).ToList()).ToList();
        return Finish(names, body, [$"Read {range.Sheet}!{range.Address}. Formulas are read as the values Excel last calculated."], formats);
    }

    private static DataTableContent Finish(List<string> names, List<List<DataValue>> body, IReadOnlyList<string> notes, string?[]? formats)
    {
        var columns = new List<DataColumn>();
        for (int i = 0; i < names.Count; i++)
        {
            var values = body.Select(row => row[i]).ToList();
            var kinds = values.Where(value => value.Kind != "empty").Select(value => value.Kind).Distinct().ToList();
            string type;
            var columnNotes = new List<string>();
            if (kinds.Count == 0)
            {
                type = "empty";
            }
            else if (kinds.Count == 1)
            {
                type = kinds[0];
            }
            else
            {
                type = "text";
                var dominant = values.Where(value => value.Kind != "empty").GroupBy(value => value.Kind).OrderByDescending(group => group.Count()).First().Key;
                var odd = values.Select((value, row) => (value, row)).Where(entry => entry.value.Kind != "empty" && entry.value.Kind != dominant).Take(5)
                    .Select(entry => $"row {entry.row + 1} \"{entry.value.Raw}\" ({entry.value.Kind})");
                columnNotes.Add($"Mixed values: mostly {dominant}, but {string.Join(", ", odd)}; the column is treated as text and nothing was converted.");
            }
            if (values.Any(value => value.Kind == "empty") && kinds.Count > 0)
                columnNotes.Add($"{values.Count(value => value.Kind == "empty")} empty cell(s) are missing values.");
            columns.Add(new DataColumn(i, names[i], type, formats?[i], columnNotes));
        }
        return new DataTableContent { Columns = columns, Rows = body, Notes = notes };
    }

    /// <summary>Classifies one text cell.</summary>
    public static DataValue Classify(string raw, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(raw);
        ArgumentNullException.ThrowIfNull(culture);
        var text = raw.Trim();
        if (text.Length == 0)
            return new DataValue("empty", null, "", raw);
        if (TryParseNumber(text, culture, out var number, out var percent))
            return new DataValue(percent ? "percent" : "number", number, raw, raw);
        if (TryParseDate(text, culture, out var date))
            return new DataValue("date", date.ToOADate(), raw, raw);
        if (string.Equals(text, "true", StringComparison.OrdinalIgnoreCase) || string.Equals(text, "false", StringComparison.OrdinalIgnoreCase))
            return new DataValue("bool", string.Equals(text, "true", StringComparison.OrdinalIgnoreCase) ? 1 : 0, raw, raw);
        return new DataValue("text", null, raw, raw);
    }

    /// <summary>
    /// Parses a number in the culture: group separators (including spaces and non-breaking spaces),
    /// the culture's decimal separator, leading +/−/-, accounting parentheses, currency symbols, and a
    /// trailing % (the value is then the fraction, e.g. "12%" → 0.12).
    /// </summary>
    public static bool TryParseNumber(string text, CultureInfo culture, out double value, out bool percent)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(culture);
        value = 0;
        percent = false;
        var s = text.Trim().Replace('−', '-').Replace('–', '-');
        bool negative = false;
        if (s.StartsWith('(') && s.EndsWith(')'))
        {
            negative = true;
            s = s[1..^1].Trim();
        }
        if (s.EndsWith('%'))
        {
            percent = true;
            s = s[..^1].Trim();
        }
        s = CurrencyAffix().Replace(s, "").Trim();
        if (s.StartsWith('+'))
            s = s[1..];
        else if (s.StartsWith('-'))
        {
            negative = !negative;
            s = s[1..];
        }
        s = s.Trim();
        if (s.Length == 0 || !Digits().IsMatch(s))
            return false;

        var format = culture.NumberFormat;
        var decimalSeparator = format.NumberDecimalSeparator;
        var group = format.NumberGroupSeparator;
        // Normalise group separators: the culture's, plus spaces and NBSPs that many cultures print.
        var cleaned = s.Replace(" ", " ", StringComparison.Ordinal).Replace(" ", " ", StringComparison.Ordinal);
        if (group.Trim().Length == 0 || group == " ")
            cleaned = cleaned.Replace(" ", "", StringComparison.Ordinal);
        else
        {
            if (cleaned.Contains(' ', StringComparison.Ordinal))
                return false;
            // A group separator must be followed by exactly three digits.
            if (cleaned.Contains(group, StringComparison.Ordinal) && !GroupedDigits(group, decimalSeparator).IsMatch(cleaned))
                return false;
            cleaned = cleaned.Replace(group, "", StringComparison.Ordinal);
        }
        if (decimalSeparator != ".")
        {
            if (cleaned.Contains('.', StringComparison.Ordinal))
                return false;
            cleaned = cleaned.Replace(decimalSeparator, ".", StringComparison.Ordinal);
        }
        if (!double.TryParse(cleaned, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value))
            return false;
        if (negative)
            value = -value;
        if (percent)
            value /= 100.0;
        return true;
    }

    /// <summary>Parses ISO dates (yyyy-MM-dd, optional time) and the culture's short date.</summary>
    public static bool TryParseDate(string text, CultureInfo culture, out DateTime value)
    {
        ArgumentNullException.ThrowIfNull(culture);
        string[] iso = ["yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm:ss"];
        if (DateTime.TryParseExact(text, iso, CultureInfo.InvariantCulture, DateTimeStyles.None, out value))
            return true;
        return culture != CultureInfo.InvariantCulture &&
            DateTime.TryParseExact(text, [culture.DateTimeFormat.ShortDatePattern, culture.DateTimeFormat.ShortDatePattern.Replace("yyyy", "yy", StringComparison.Ordinal)], culture, DateTimeStyles.None, out value);
    }

    /// <summary>Display text of an XLSX cell using its number format.</summary>
    public static string Display(XlsxCell cell, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(cell);
        return cell.Kind switch
        {
            "number" or "percent" => NumberFormatter.Format(cell.Number!.Value, cell.Format, culture),
            "empty" => "",
            _ => cell.Text ?? "",
        };
    }

    [GeneratedRegex(@"^[$€£¥₽]|[$€£¥₽]$|\s?(руб\.?|р\.|RUB|USD|EUR)$", RegexOptions.CultureInvariant)]
    private static partial Regex CurrencyAffix();

    [GeneratedRegex(@"^[0-9][0-9\s  .,']*$", RegexOptions.CultureInvariant)]
    private static partial Regex Digits();

    private static Regex GroupedDigits(string group, string decimalSeparator) =>
        new($@"^[0-9]{{1,3}}({Regex.Escape(group)}[0-9]{{3}})*({Regex.Escape(decimalSeparator)}[0-9]+)?$", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
}

/// <summary>
/// A subset of Excel number formats: sections (positive;negative;zero), 0 and # digits, thousands
/// grouping with ",", decimals, %, quoted literals, backslash escapes, and literal currency
/// symbols; [color] and [$-xxx] tags, "_" padding, and "*" fill are ignored. "General" or null
/// prints the shortest round-trip number. Culture supplies the separators. Pure.
/// </summary>
public static class NumberFormatter
{
    /// <summary>Formats a number.</summary>
    public static string Format(double value, string? code, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(culture);
        if (double.IsNaN(value))
            return "";
        if (string.IsNullOrWhiteSpace(code) || string.Equals(code, "General", StringComparison.OrdinalIgnoreCase) || code == "@")
            return General(value, culture);

        var sections = SplitSections(code);
        string section;
        bool useAbsolute = false;
        if (value < 0 && sections.Count >= 2)
        {
            section = sections[1];
            useAbsolute = true;
        }
        else if (value == 0 && sections.Count >= 3)
        {
            section = sections[2];
        }
        else
        {
            section = sections[0];
        }
        var number = useAbsolute ? Math.Abs(value) : value;

        var literalPrefix = new StringBuilder();
        var literalSuffix = new StringBuilder();
        var pattern = new StringBuilder();
        bool percent = false;
        bool inNumber = false;
        bool afterNumber = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            string? literal = null;
            if (c == '"')
            {
                int end = section.IndexOf('"', i + 1);
                literal = end < 0 ? section[(i + 1)..] : section[(i + 1)..end];
                i = end < 0 ? section.Length : end;
            }
            else if (c == '\\' && i + 1 < section.Length)
            {
                literal = section[++i].ToString();
            }
            else if (c == '[')
            {
                int end = section.IndexOf(']', i + 1);
                i = end < 0 ? section.Length : end;
                continue;
            }
            else if (c is '_' or '*')
            {
                i++;
                continue;
            }
            else if (c is '0' or '#' or '?' || (c is ',' or '.' && (inNumber || (i + 1 < section.Length && section[i + 1] is '0' or '#'))))
            {
                if (afterNumber)
                {
                    literalSuffix.Append(c);
                    continue;
                }
                inNumber = true;
                pattern.Append(c == '?' ? '#' : c);
                continue;
            }
            else if (c == '%')
            {
                percent = true;
                literal = culture.NumberFormat.PercentSymbol;
            }
            else
            {
                literal = c.ToString();
            }

            if (inNumber)
                afterNumber = true;
            if (afterNumber)
                literalSuffix.Append(literal);
            else
                literalPrefix.Append(literal);
        }

        if (percent)
            number *= 100;
        var numberPattern = pattern.ToString();
        if (numberPattern.Length == 0)
            return literalPrefix.ToString() + literalSuffix;
        bool grouping = numberPattern.Contains(',', StringComparison.Ordinal) && numberPattern.IndexOf(',', StringComparison.Ordinal) < (numberPattern.Contains('.', StringComparison.Ordinal) ? numberPattern.IndexOf('.', StringComparison.Ordinal) : numberPattern.Length);
        var decimals = numberPattern.Contains('.', StringComparison.Ordinal) ? numberPattern[(numberPattern.IndexOf('.', StringComparison.Ordinal) + 1)..].Count(ch => ch is '0' or '#') : 0;
        var required = numberPattern.Contains('.', StringComparison.Ordinal) ? numberPattern[(numberPattern.IndexOf('.', StringComparison.Ordinal) + 1)..].Count(ch => ch == '0') : 0;
        var netPattern = (grouping ? "#,##0" : "0") + (decimals > 0 ? "." + new string('0', required) + new string('#', decimals - required) : "");
        var formatted = number.ToString(netPattern, culture);
        if (!useAbsolute && number < 0 && sections.Count == 1)
            formatted = formatted.Replace(culture.NumberFormat.NegativeSign, "-", StringComparison.Ordinal);
        return literalPrefix + formatted + literalSuffix;
    }

    private static string General(double value, CultureInfo culture)
    {
        var rounded = Math.Round(value, 10);
        return rounded.ToString("0.##########", culture);
    }

    private static List<string> SplitSections(string code)
    {
        var sections = new List<string>();
        var current = new StringBuilder();
        bool quoted = false;
        foreach (var c in code)
        {
            if (c == '"')
                quoted = !quoted;
            if (c == ';' && !quoted)
            {
                sections.Add(current.ToString());
                current.Clear();
                continue;
            }
            current.Append(c);
        }
        sections.Add(current.ToString());
        return sections;
    }
}
