// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using System.IO.Compression;
using System.Text;
using Sbroenne.PowerPointMcp.Core.Data;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Data;

/// <summary>CSV and XLSX reading, typing, and number formatting (no PowerPoint, no Excel).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Data")]
public sealed class DataReaderTests
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    [Fact]
    public void Csv_ParsesQuotesEmbeddedDelimitersAndLineBreaks()
    {
        var rows = CsvReader.Parse("name,note\r\n\"Smith, J.\",\"said \"\"hi\"\"\nthen left\"\r\nplain,\r\n", ',');
        Assert.Equal(3, rows.Count);
        Assert.Equal(["Smith, J.", "said \"hi\"\nthen left"], rows[1]);
        Assert.Equal(["plain", ""], rows[2]);
    }

    [Theory]
    [InlineData("a;b;c\n1;2;3\n4;5;6", ';')]
    [InlineData("a,b,c\n1,2,3", ',')]
    [InlineData("a\tb\n1\t2", '\t')]
    [InlineData("\"x;y\",b\n\"1;2\",3", ',')]
    [InlineData("Регион;Выручка, млн\nМосква;1,5\nКазань;2,0", ';')]
    public void Csv_DetectsDelimiter(string text, char expected) => Assert.Equal(expected, CsvReader.Detect(text));

    [Fact]
    public void Csv_DecodesBomUtf8AndWindows1251()
    {
        var utf8 = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes("Город;Сумма\nМосква;10")).ToArray();
        var content = CsvReader.Read(utf8);
        Assert.Equal(("utf-8", ';', "Москва"), (content.Encoding, content.Delimiter, content.Rows[1][0]));

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var cp1251 = Encoding.GetEncoding(1251).GetBytes("Город;Сумма\nКазань;20");
        var decoded = CsvReader.Read(cp1251);
        Assert.Equal("windows-1251", decoded.Encoding);
        Assert.Equal("Казань", decoded.Rows[1][0]);
        Assert.Contains(decoded.Notes, note => note.Contains("not valid UTF-8", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("1 234,5", "ru-RU", 1234.5, false)]
    [InlineData("1 234,5", "ru-RU", 1234.5, false)]
    [InlineData("1,234.5", "en-US", 1234.5, false)]
    [InlineData("12%", "en-US", 0.12, true)]
    [InlineData("12,5 %", "ru-RU", 0.125, true)]
    [InlineData("(12.5)", "en-US", -12.5, false)]
    [InlineData("−3", "en-US", -3, false)]
    [InlineData("$4.2", "en-US", 4.2, false)]
    [InlineData("4 200 ₽", "ru-RU", 4200, false)]
    [InlineData("+7", "en-US", 7, false)]
    public void TryParseNumber_AcceptsCultureForms(string text, string culture, double expected, bool percent)
    {
        Assert.True(DataTyping.TryParseNumber(text, CultureInfo.GetCultureInfo(culture), out var value, out var isPercent));
        Assert.Equal(expected, value, 6);
        Assert.Equal(percent, isPercent);
    }

    [Theory]
    [InlineData("1,5", "en-US")]
    [InlineData("1.5", "ru-RU")]
    [InlineData("12,34,5", "en-US")]
    [InlineData("Q3", "en-US")]
    [InlineData("2025-01", "en-US")]
    [InlineData("", "en-US")]
    public void TryParseNumber_RejectsAmbiguousOrNonNumbers(string text, string culture)
    {
        Assert.False(DataTyping.TryParseNumber(text, CultureInfo.GetCultureInfo(culture), out _, out _));
    }

    [Fact]
    public void Typing_MixedColumnsBecomeTextWithNotes_NothingConverted()
    {
        var table = DataTyping.FromRows(
        [
            ["Region", "Revenue", "Share", "Date", "Note"],
            ["EMEA", "3,1", "12%", "2025-03-01", "ok"],
            ["APAC", "n/a", "8%", "2025-03-02", ""],
            ["LATAM", "2,4", "", "01.04.2025", "x"],
        ], header: true, Ru);
        Assert.Equal(["text", "text", "percent", "date", "text"], table.Columns.Select(column => column.Type));
        Assert.Contains(table.Columns[1].Notes, note => note.Contains("row 2 \"n/a\" (text)", StringComparison.Ordinal));
        Assert.Contains(table.Columns[2].Notes, note => note.Contains("1 empty cell", StringComparison.Ordinal));
        Assert.Equal("3,1", table.Rows[0][1].Display);
        Assert.Equal("n/a", table.Rows[1][1].Display);
        Assert.Equal(0.12, table.Rows[0][2].Number!.Value, 6);
        Assert.Equal(2, table.ColumnIndex("share"));
        Assert.Equal(0, table.ColumnIndex("1"));
        Assert.Equal(-1, table.ColumnIndex("missing"));
    }

    [Theory]
    [InlineData(1234.567, "#,##0.00", "en-US", "1,234.57")]
    [InlineData(1234.567, "#,##0.0", "ru-RU", "1 234,6")]
    [InlineData(0.256, "0%", "en-US", "26%")]
    [InlineData(0.256, "0.0%", "en-US", "25.6%")]
    [InlineData(-5, "0;(0)", "en-US", "(5)")]
    [InlineData(0, "0;-0;\"—\"", "en-US", "—")]
    [InlineData(4.2, "\"$\"0.0\"M\"", "en-US", "$4.2M")]
    [InlineData(1500, "#,##0 \"₽\"", "ru-RU", "1 500 ₽")]
    [InlineData(12.5, "[Red]0.0", "en-US", "12.5")]
    [InlineData(3.14159, null, "en-US", "3.14159")]
    [InlineData(1.5, "General", "ru-RU", "1,5")]
    public void NumberFormatter_RendersCommonExcelFormats(double value, string? code, string culture, string expected)
    {
        Assert.Equal(expected, NumberFormatter.Format(value, code, CultureInfo.GetCultureInfo(culture)));
    }

    [Fact]
    public void Xlsx_ReadsSharedStringsNumbersPercentsDatesAndRanges()
    {
        var path = WriteWorkbook();
        try
        {
            var used = XlsxReader.Read(path);
            Assert.Equal(("Data", "A1:D3"), (used.Sheet, used.Address));
            Assert.Equal(["Data", "Notes"], used.SheetNames);
            Assert.Equal("Регион", used.Rows[0][0].Text);
            Assert.Equal(("number", 3.1), (used.Rows[1][1].Kind, used.Rows[1][1].Number!.Value));
            Assert.Equal("percent", used.Rows[1][2].Kind);
            Assert.Equal(("date", "2025-03-01"), (used.Rows[1][3].Kind, used.Rows[1][3].Text));
            Assert.Equal("empty", used.Rows[2][2].Kind);

            var table = DataTyping.FromXlsx(used, header: true, Ru);
            Assert.Equal(["text", "number", "percent", "date"], table.Columns.Select(column => column.Type));
            Assert.Equal("12,5%", table.Rows[0][2].Display);
            Assert.Equal("3,1", table.Rows[0][1].Display);

            var part = XlsxReader.Read(path, sheet: "data", range: "B2:C3");
            Assert.Equal("B2:C3", part.Address);
            Assert.Equal(2, part.Rows.Count);
            Assert.Throws<ArgumentException>(() => XlsxReader.Read(path, sheet: "Missing"));
            Assert.Throws<ArgumentException>(() => XlsxReader.Read(path, range: "C3:A1"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(45717, false, "2025-03-01")]
    [InlineData(1, false, "1900-01-01")]
    [InlineData(61, false, "1900-03-01")]
    [InlineData(0, true, "1904-01-01")]
    public void FromSerial_HandlesBothDateSystems(double serial, bool date1904, string expected)
    {
        Assert.Equal(expected, XlsxReader.FromSerial(serial, date1904).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("Z10", 10, 26)]
    [InlineData("AA3", 3, 27)]
    [InlineData("xfd1", 1, 16384)]
    public void CellReferences(string reference, int row, int column)
    {
        Assert.True(XlsxReader.TryParseCell(reference, out var r, out var c));
        Assert.Equal((row, column), (r, c));
        Assert.Equal(reference.ToUpperInvariant().TrimEnd("0123456789".ToCharArray()), XlsxReader.ColumnName(column));
    }

    /// <summary>A minimal valid workbook: shared strings, a percent style, a date style, and two sheets.</summary>
    internal static string WriteWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptmcp-{Guid.NewGuid():N}.xlsx");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string content)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Add("[Content_Types].xml", """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>""");
        Add("xl/workbook.xml", $"""<?xml version="1.0"?><workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>""");
        Add("xl/_rels/workbook.xml.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="worksheet" Target="worksheets/sheet2.xml"/></Relationships>""");
        Add("xl/sharedStrings.xml", $"""<?xml version="1.0"?><sst xmlns="{ns}"><si><t>Регион</t></si><si><t>Выручка</t></si><si><r><t>Do</t></r><r><t>la</t></r></si><si><t>Дата</t></si><si><t>Москва</t></si><si><t>Казань</t></si></sst>""");
        Add("xl/styles.xml", $"""<?xml version="1.0"?><styleSheet xmlns="{ns}"><numFmts><numFmt numFmtId="164" formatCode="0.0%"/><numFmt numFmtId="165" formatCode="dd.mm.yyyy"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs></styleSheet>""");
        Add("xl/worksheets/sheet1.xml", $"""<?xml version="1.0"?><worksheet xmlns="{ns}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row><row r="2"><c r="A2" t="s"><v>4</v></c><c r="B2"><v>3.1</v></c><c r="C2" s="1"><v>0.125</v></c><c r="D2" s="2"><v>45717</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Казань</t></is></c><c r="B3"><f>B2*2</f><v>6.2</v></c><c r="D3" s="2"><v>45718</v></c></row></sheetData></worksheet>""");
        Add("xl/worksheets/sheet2.xml", $"""<?xml version="1.0"?><worksheet xmlns="{ns}"><sheetData/></worksheet>""");
        return path;
    }
}
