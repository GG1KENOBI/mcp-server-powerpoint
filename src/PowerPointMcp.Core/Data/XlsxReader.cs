using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>A worksheet cell value read from an .xlsx file.</summary>
public sealed record XlsxCell(string Kind, double? Number, string? Text, string? Format);

/// <summary>A rectangular block of cells.</summary>
public sealed record XlsxRange(string Sheet, string Address, IReadOnlyList<IReadOnlyList<XlsxCell>> Rows, IReadOnlyList<string> SheetNames);

/// <summary>
/// Reads .xlsx workbooks directly (Office Open XML in a zip) without Excel: sheet names, a cell
/// range or the used range, shared and inline strings, booleans, errors, numbers, and the number
/// format of each cell so dates and percentages are recognised. Formulas are read as their cached
/// values (no recalculation). Pure and offline.
/// </summary>
public static partial class XlsxReader
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>Built-in number formats that are dates or times.</summary>
    private static readonly HashSet<int> DateFormats = [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47];

    /// <summary>Built-in percent formats.</summary>
    private static readonly HashSet<int> PercentFormats = [9, 10];

    /// <summary>Reads a range ("A1:D20", default: the used range) from a sheet (default: the first).</summary>
    public static XlsxRange Read(string path, string? sheet = null, string? range = null)
    {
        using var archive = ZipFile.OpenRead(path);
        var workbook = Load(archive, "xl/workbook.xml") ?? throw new InvalidDataException("The file has no xl/workbook.xml; it is not an .xlsx workbook.");
        bool date1904 = (string?)workbook.Root!.Element(Main + "workbookPr")?.Attribute("date1904") is "1" or "true";
        var sheets = workbook.Root!.Element(Main + "sheets")?.Elements(Main + "sheet").ToList() ?? [];
        var names = sheets.Select(element => (string?)element.Attribute("name") ?? "").ToList();
        if (sheets.Count == 0)
            throw new InvalidDataException("The workbook has no sheets.");
        var chosen = sheet is null ? sheets[0] : sheets.FirstOrDefault(element => string.Equals((string?)element.Attribute("name"), sheet, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"Sheet '{sheet}' not found. Sheets: {string.Join(", ", names)}.");
        var relationId = (string?)chosen.Attribute(Relationships + "id");
        var rels = Load(archive, "xl/_rels/workbook.xml.rels");
        var target = rels?.Root?.Elements(PackageRelationships + "Relationship").FirstOrDefault(element => (string?)element.Attribute("Id") == relationId)?.Attribute("Target")?.Value
            ?? throw new InvalidDataException("The sheet's part could not be found.");
        var sheetPath = target.StartsWith('/') ? target.TrimStart('/') : "xl/" + target;
        var sheetXml = Load(archive, sheetPath) ?? throw new InvalidDataException($"Missing {sheetPath}.");

        var shared = Load(archive, "xl/sharedStrings.xml")?.Root?.Elements(Main + "si").Select(StringOf).ToList() ?? [];
        var styles = ReadStyles(Load(archive, "xl/styles.xml"));

        var cells = new Dictionary<(int Row, int Column), XlsxCell>();
        int maxRow = 0, maxColumn = 0, minRow = int.MaxValue, minColumn = int.MaxValue;
        foreach (var c in sheetXml.Root!.Descendants(Main + "c"))
        {
            var reference = (string?)c.Attribute("r");
            if (reference is null || !TryParseCell(reference, out var row, out var column))
                continue;
            var cell = ReadCell(c, shared, styles, date1904);
            if (cell.Kind == "empty")
                continue;
            cells[(row, column)] = cell;
            maxRow = Math.Max(maxRow, row);
            maxColumn = Math.Max(maxColumn, column);
            minRow = Math.Min(minRow, row);
            minColumn = Math.Min(minColumn, column);
        }

        int top, left, bottom, right;
        if (range is not null)
        {
            var parts = range.Replace("$", "", StringComparison.Ordinal).Split(':');
            if (parts.Length != 2 || !TryParseCell(parts[0], out top, out left) || !TryParseCell(parts[1], out bottom, out right) || bottom < top || right < left)
                throw new ArgumentException($"Range '{range}' is not like A1:D20.");
        }
        else if (cells.Count == 0)
        {
            return new XlsxRange((string)chosen.Attribute("name")!, "", [], names);
        }
        else
        {
            (top, left, bottom, right) = (minRow, minColumn, maxRow, maxColumn);
        }
        if ((long)(bottom - top + 1) * (right - left + 1) > 200_000)
            throw new ArgumentException("The range has more than 200,000 cells; pass a smaller range.");

        var rows = new List<IReadOnlyList<XlsxCell>>();
        for (int r = top; r <= bottom; r++)
        {
            var row = new List<XlsxCell>();
            for (int column = left; column <= right; column++)
                row.Add(cells.TryGetValue((r, column), out var cell) ? cell : new XlsxCell("empty", null, null, null));
            rows.Add(row);
        }
        return new XlsxRange((string)chosen.Attribute("name")!, $"{ColumnName(left)}{top}:{ColumnName(right)}{bottom}", rows, names);
    }

    private static XDocument? Load(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        if (entry is null)
            return null;
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static string StringOf(XElement si) =>
        si.Element(Main + "t") is { } plain ? plain.Value : string.Concat(si.Elements(Main + "r").Select(run => run.Element(Main + "t")?.Value ?? ""));

    private sealed record StyleInfo(int NumberFormatId, string? Code);

    private static List<StyleInfo> ReadStyles(XDocument? styles)
    {
        if (styles?.Root is null)
            return [];
        var custom = styles.Root.Element(Main + "numFmts")?.Elements(Main + "numFmt")
            .ToDictionary(element => (int)element.Attribute("numFmtId")!, element => (string?)element.Attribute("formatCode")) ?? [];
        return styles.Root.Element(Main + "cellXfs")?.Elements(Main + "xf").Select(xf =>
        {
            var id = (int?)xf.Attribute("numFmtId") ?? 0;
            return new StyleInfo(id, custom.GetValueOrDefault(id) ?? BuiltInCode(id));
        }).ToList() ?? [];
    }

    private static string? BuiltInCode(int id) => id switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        14 => "yyyy-mm-dd",
        _ => null,
    };

    private static XlsxCell ReadCell(XElement c, List<string> shared, List<StyleInfo> styles, bool date1904)
    {
        var type = (string?)c.Attribute("t") ?? "n";
        var value = c.Element(Main + "v")?.Value;
        var style = (int?)c.Attribute("s") is { } index && index < styles.Count ? styles[index] : null;
        switch (type)
        {
            case "s":
                return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var stringIndex) && stringIndex < shared.Count
                    ? new XlsxCell("text", null, shared[stringIndex], null)
                    : new XlsxCell("empty", null, null, null);
            case "inlineStr":
                return new XlsxCell("text", null, c.Element(Main + "is") is { } inline ? StringOf(inline) : "", null);
            case "str":
                return new XlsxCell("text", null, value ?? "", null);
            case "b":
                return new XlsxCell("bool", value == "1" ? 1 : 0, value == "1" ? "TRUE" : "FALSE", null);
            case "e":
                return new XlsxCell("error", null, value, null);
            default:
                if (value is null || !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    return new XlsxCell("empty", null, null, null);
                if (style is not null && IsDate(style))
                {
                    var date = FromSerial(number, date1904);
                    var text = date.TimeOfDay == TimeSpan.Zero ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
                    return new XlsxCell("date", number, text, style.Code);
                }
                if (style is not null && (PercentFormats.Contains(style.NumberFormatId) || (style.Code?.Contains('%', StringComparison.Ordinal) ?? false)))
                    return new XlsxCell("percent", number, null, style.Code);
                return new XlsxCell("number", number, null, style?.Code);
        }
    }

    private static bool IsDate(StyleInfo style)
    {
        if (DateFormats.Contains(style.NumberFormatId))
            return true;
        if (style.Code is null)
            return false;
        // Strip quoted literals and bracketed sections ([Red], [$-419]) before looking for date tokens.
        var code = QuotedOrBracketed().Replace(style.Code, "");
        return DateToken().IsMatch(code);
    }

    /// <summary>Converts an Excel serial date to a DateTime (1900 system with the 1900 leap-year bug, or 1904).</summary>
    public static DateTime FromSerial(double serial, bool date1904 = false)
    {
        if (date1904)
            return new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(serial);
        // Serial 60 is the non-existent 1900-02-29; dates after it are shifted by one day.
        var days = serial < 60 ? serial : serial - 1;
        return new DateTime(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified).AddDays(days);
    }

    /// <summary>Parses "B12" into 1-based row and column.</summary>
    public static bool TryParseCell(string reference, out int row, out int column)
    {
        row = column = 0;
        var match = CellReference().Match(reference ?? "");
        if (!match.Success)
            return false;
        foreach (var letter in match.Groups["column"].Value.ToUpperInvariant())
            column = (column * 26) + (letter - 'A' + 1);
        return int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out row) && row > 0 && column > 0 && column <= 16384;
    }

    /// <summary>Column letters for a 1-based column.</summary>
    public static string ColumnName(int column)
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

    [GeneratedRegex("^(?<column>[A-Za-z]{1,3})(?<row>[0-9]{1,7})$", RegexOptions.CultureInvariant)]
    private static partial Regex CellReference();

    [GeneratedRegex("\"[^\"]*\"|\\[[^\\]]*\\]", RegexOptions.CultureInvariant)]
    private static partial Regex QuotedOrBracketed();

    [GeneratedRegex("[dmyhs]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex DateToken();
}
