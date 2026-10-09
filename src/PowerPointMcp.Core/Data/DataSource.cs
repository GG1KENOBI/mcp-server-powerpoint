using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>
/// Where a table or chart gets its data: stored as JSON in the object's PPTMCP_BINDING tag so
/// refresh can re-read the same file, sheet, and range later.
/// </summary>
public sealed record DataBinding
{
    /// <summary>Full path of the CSV or XLSX file.</summary>
    public required string Source { get; init; }

    /// <summary>Worksheet name (XLSX).</summary>
    public string? Sheet { get; init; }

    /// <summary>Cell range (XLSX), e.g. A1:D20.</summary>
    public string? Range { get; init; }

    /// <summary>Whether the first row is a header.</summary>
    public bool Header { get; init; } = true;

    /// <summary>CSV delimiter override.</summary>
    public string? Delimiter { get; init; }

    /// <summary>CSV encoding override.</summary>
    public string? Encoding { get; init; }

    /// <summary>Culture for parsing and formatting numbers ("invariant", "ru-RU", "en-US", ...).</summary>
    public string Culture { get; init; } = "invariant";

    /// <summary>table or chart.</summary>
    public required string Kind { get; init; }

    /// <summary>Chart: category column (name or 1-based number).</summary>
    public string? CategoryColumn { get; init; }

    /// <summary>Chart: value columns.</summary>
    public IReadOnlyList<string>? ValueColumns { get; init; }

    /// <summary>Excel number format for chart values.</summary>
    public string? NumberFormat { get; init; }

    /// <summary>Table split across slides: 0-based first data row of this part.</summary>
    public int RowOffset { get; init; }

    /// <summary>Table split across slides: data rows in this part (null = all remaining).</summary>
    public int? RowCount { get; init; }

    /// <summary>SHA-256 (first 16 hex digits) of the source file when last read.</summary>
    public string? Hash { get; init; }

    /// <summary>When the data was last written (UTC, ISO 8601).</summary>
    public string? RefreshedAt { get; init; }

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>Serializes the binding.</summary>
    public string ToJson() => JsonSerializer.Serialize(this, Options);

    /// <summary>Parses a binding tag; null when it is not a data binding.</summary>
    public static DataBinding? Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json) || !json.TrimStart().StartsWith('{'))
            return null;
        try
        {
            return JsonSerializer.Deserialize<DataBinding>(json, Options);
        }
        catch (JsonException)
        {
            return null;
        }
    }
}

/// <summary>Loads bound data. Pure apart from reading the source file.</summary>
public static class DataLoader
{
    /// <summary>Resolves a culture name ("invariant" or an IETF tag).</summary>
    public static CultureInfo Culture(string? name)
    {
        if (string.IsNullOrWhiteSpace(name) || string.Equals(name, "invariant", StringComparison.OrdinalIgnoreCase))
            return CultureInfo.InvariantCulture;
        try
        {
            return CultureInfo.GetCultureInfo(name);
        }
        catch (CultureNotFoundException)
        {
            throw new ArgumentException($"Unknown culture '{name}'. Use invariant, ru-RU, en-US, de-DE, ...");
        }
    }

    /// <summary>Reads the source described by a binding.</summary>
    public static (DataTableContent Table, string Hash) Load(DataBinding binding)
    {
        ArgumentNullException.ThrowIfNull(binding);
        if (!File.Exists(binding.Source))
            throw new FileNotFoundException($"The data file '{binding.Source}' does not exist on this computer.", binding.Source);
        var culture = Culture(binding.Culture);
        var hash = HashOf(binding.Source);
        var extension = Path.GetExtension(binding.Source).ToLowerInvariant();
        if (extension is ".xlsx" or ".xlsm")
        {
            var range = XlsxReader.Read(binding.Source, binding.Sheet, binding.Range);
            return (DataTyping.FromXlsx(range, binding.Header, culture), hash);
        }
        if (extension == ".xls")
            throw new ArgumentException("Legacy .xls workbooks are not supported; save the file as .xlsx or export CSV.");
        if (binding.Sheet is not null || binding.Range is not null)
            throw new ArgumentException("sheet and range apply to .xlsx files only.");
        char? delimiter = binding.Delimiter switch
        {
            null => null,
            "tab" or "\\t" or "\t" => '\t',
            { Length: 1 } single => single[0],
            _ => throw new ArgumentException("delimiter must be one character or 'tab'."),
        };
        var csv = CsvReader.ReadFile(binding.Source, delimiter, binding.Encoding);
        var notes = csv.Notes.Append($"Encoding {csv.Encoding}.").ToList();
        return (DataTyping.FromRows(csv.Rows, binding.Header, culture, notes), hash);
    }

    /// <summary>First 16 hex digits of the file's SHA-256.</summary>
    public static string HashOf(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream))[..16].ToLowerInvariant();
    }
}
