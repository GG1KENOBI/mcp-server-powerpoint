using System.Text;

namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>A CSV file decoded into rows of strings.</summary>
public sealed record CsvContent(IReadOnlyList<IReadOnlyList<string>> Rows, char Delimiter, string Encoding, IReadOnlyList<string> Notes);

/// <summary>
/// RFC 4180 CSV reader: quoted fields with embedded delimiters, quotes, and line breaks; delimiter
/// detection (comma, semicolon, tab, pipe); encoding detection (BOM, strict UTF-8, otherwise
/// Windows-1251 for Cyrillic bytes or Windows-1252). Pure and offline.
/// </summary>
public static class CsvReader
{
    private static readonly char[] Candidates = [',', ';', '\t', '|'];

    static CsvReader()
    {
        System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>Reads a file; <paramref name="delimiter"/> and <paramref name="encoding"/> override detection.</summary>
    public static CsvContent ReadFile(string path, char? delimiter = null, string? encoding = null) =>
        Read(File.ReadAllBytes(path), delimiter, encoding);

    /// <summary>Decodes and parses bytes.</summary>
    public static CsvContent Read(byte[] bytes, char? delimiter = null, string? encoding = null)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        var notes = new List<string>();
        var (text, encodingName) = Decode(bytes, encoding, notes);
        var separator = delimiter ?? Detect(text);
        if (delimiter is null)
            notes.Add($"Delimiter detected as {Describe(separator)}.");
        return new CsvContent(Parse(text, separator), separator, encodingName, notes);
    }

    /// <summary>Parses text with a known delimiter.</summary>
    public static IReadOnlyList<IReadOnlyList<string>> Parse(string text, char delimiter)
    {
        ArgumentNullException.ThrowIfNull(text);
        var rows = new List<IReadOnlyList<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool quoted = false;
        bool fieldStarted = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (quoted)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        field.Append('"');
                        i++;
                    }
                    else
                    {
                        quoted = false;
                    }
                }
                else
                {
                    field.Append(c);
                }
                continue;
            }

            if (c == '"' && !fieldStarted)
            {
                quoted = true;
                fieldStarted = true;
            }
            else if (c == delimiter)
            {
                row.Add(field.ToString());
                field.Clear();
                fieldStarted = false;
            }
            else if (c is '\r' or '\n')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i++;
                row.Add(field.ToString());
                field.Clear();
                fieldStarted = false;
                if (!(row.Count == 1 && row[0].Length == 0))
                    rows.Add(row);
                row = [];
            }
            else
            {
                field.Append(c);
                fieldStarted = true;
            }
        }
        if (fieldStarted || field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            if (!(row.Count == 1 && row[0].Length == 0))
                rows.Add(row);
        }
        return rows;
    }

    /// <summary>The delimiter that splits the first lines most consistently (outside quotes).</summary>
    public static char Detect(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var lines = SampleLines(text, 20);
        char best = ',';
        double bestScore = -1;
        foreach (var candidate in Candidates)
        {
            var counts = lines.Select(line => CountOutsideQuotes(line, candidate)).ToList();
            if (counts.Count == 0 || counts.All(count => count == 0))
                continue;
            // Prefer delimiters that appear the same number of times on every line.
            var mode = counts.GroupBy(count => count).OrderByDescending(group => group.Count()).First();
            double score = mode.Key * (mode.Count() / (double)counts.Count);
            // Commas between digits are usually decimal separators (1,5), not delimiters.
            if (candidate == ',')
                score *= 1 - (0.5 * DigitFlankedShare(lines));
            if (score > bestScore)
            {
                bestScore = score;
                best = candidate;
            }
        }
        return best;
    }

    private static List<string> SampleLines(string text, int count)
    {
        var lines = new List<string>();
        var current = new StringBuilder();
        bool quoted = false;
        foreach (var c in text)
        {
            if (c == '"')
                quoted = !quoted;
            if ((c == '\n') && !quoted)
            {
                lines.Add(current.ToString().TrimEnd('\r'));
                current.Clear();
                if (lines.Count >= count)
                    break;
                continue;
            }
            current.Append(c);
        }
        if (current.Length > 0 && lines.Count < count)
            lines.Add(current.ToString());
        return lines.Where(line => line.Length > 0).ToList();
    }

    private static double DigitFlankedShare(List<string> lines)
    {
        int total = 0, flanked = 0;
        foreach (var line in lines)
        {
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] != ',')
                    continue;
                total++;
                if (i > 0 && i + 1 < line.Length && char.IsDigit(line[i - 1]) && char.IsDigit(line[i + 1]))
                    flanked++;
            }
        }
        return total == 0 ? 0 : flanked / (double)total;
    }

    private static int CountOutsideQuotes(string line, char delimiter)
    {
        int count = 0;
        bool quoted = false;
        foreach (var c in line)
        {
            if (c == '"')
                quoted = !quoted;
            else if (c == delimiter && !quoted)
                count++;
        }
        return count;
    }

    private static (string Text, string Encoding) Decode(byte[] bytes, string? requested, List<string> notes)
    {
        if (requested is not null)
        {
            var named = System.Text.Encoding.GetEncoding(requested);
            return (StripBom(named.GetString(bytes)), named.WebName);
        }
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return (System.Text.Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3), "utf-8");
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return (System.Text.Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2), "utf-16");
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return (System.Text.Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2), "utf-16BE");
        try
        {
            return (new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes), "utf-8");
        }
        catch (DecoderFallbackException)
        {
            // Not UTF-8: Russian Excel exports are usually Windows-1251.
            int high = bytes.Count(value => value >= 0xC0);
            int total = bytes.Count(value => value >= 0x80);
            var fallback = total > 0 && high >= total * 0.6 ? System.Text.Encoding.GetEncoding(1251) : System.Text.Encoding.GetEncoding(1252);
            notes.Add($"The file is not valid UTF-8; decoded as {fallback.WebName}. Pass encoding to override.");
            return (fallback.GetString(bytes), fallback.WebName);
        }
    }

    private static string StripBom(string text) => text.Length > 0 && text[0] == '﻿' ? text[1..] : text;

    private static string Describe(char delimiter) => delimiter switch
    {
        '\t' => "tab",
        ';' => "semicolon",
        '|' => "pipe",
        _ => "comma",
    };
}
