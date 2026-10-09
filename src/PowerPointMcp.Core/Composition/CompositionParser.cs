using System.Text.Json;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>Result of parsing a composition.</summary>
public sealed class CompositionParseResult
{
    /// <summary>The validated spec when there are no errors.</summary>
    public CompositionSpec? Spec { get; init; }

    /// <summary>Problems as "$.path: message".</summary>
    public required IReadOnlyList<string> Errors { get; init; }

    /// <summary>Ignored properties and accessibility or style concerns.</summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}

/// <summary>
/// Parses and validates composition JSON with JSONPath-style error locations. Nothing is
/// shortened, reformatted, or coerced: wrong types are errors, not conversions.
/// </summary>
public static partial class CompositionParser
{
    private static readonly HashSet<string> RootKeys = new(StringComparer.Ordinal)
    {
        "schema", "kind", "id", "title", "subtitle", "insert_at", "profile", "notes", "source", "takeaway", "date", "author",
        "section_number", "points", "columns", "cards", "kpis", "image", "table", "chart", "milestones", "steps", "tree",
        "quote", "references", "fit", "metadata",
    };

    private static readonly string[] FitPolicies = ["shrink", "split", "report"];
    private static readonly string[] Trends = ["up", "down", "flat"];
    private static readonly string[] Sentiments = ["positive", "negative", "neutral"];
    private static readonly string[] Statuses = ["done", "current", "planned"];
    private static readonly string[] ImageFits = ["cover", "contain"];
    private static readonly string[] ImagePositions = ["left", "right"];
    private static readonly string[] Alignments = ["left", "center", "right"];

    /// <summary>Parses composition JSON.</summary>
    public static CompositionParseResult Parse(string json)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json ?? "", new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
        }
        catch (JsonException ex)
        {
            return new CompositionParseResult
            {
                Errors = [$"$: invalid JSON at line {(ex.LineNumber ?? 0) + 1}, position {(ex.BytePositionInLine ?? 0) + 1}: {ex.Message}"],
                Warnings = [],
            };
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                return new CompositionParseResult { Errors = ["$: the composition must be a JSON object."], Warnings = [] };

            var reader = new JsonSpecReader(errors, warnings);
            foreach (var property in root.EnumerateObject().Where(property => !RootKeys.Contains(property.Name)))
                warnings.Add($"$.{property.Name}: unknown property ignored.");

            var schema = reader.String(root, "$", "schema", required: false);
            if (schema is not null && schema != CompositionSpec.SchemaId)
                errors.Add($"$.schema: expected {CompositionSpec.SchemaId}.");

            var kind = reader.String(root, "$", "kind", required: true);
            if (kind is not null && CompositionKinds.Find(kind) is null)
            {
                errors.Add($"$.kind: unknown kind '{kind}'. Kinds: {string.Join(", ", CompositionKinds.All.Select(item => item.Name))}.");
                kind = null;
            }

            var id = reader.String(root, "$", "id", required: false);
            if (id is not null && !AppId().IsMatch(id))
                errors.Add("$.id: use 1-64 letters, digits, dot, dash, underscore, or tilde.");

            var insertAt = reader.Int(root, "$", "insert_at", 1, 10000);
            var fit = ReadFit(reader, root);
            var metadata = ReadMetadata(reader, root);

            var spec = new CompositionSpec
            {
                Kind = kind ?? "",
                Id = id,
                Title = reader.String(root, "$", "title", required: false),
                Subtitle = reader.String(root, "$", "subtitle", required: false),
                InsertAt = insertAt,
                Profile = reader.String(root, "$", "profile", required: false),
                Notes = reader.String(root, "$", "notes", required: false, allowEmpty: true),
                Source = reader.String(root, "$", "source", required: false),
                Takeaway = reader.String(root, "$", "takeaway", required: false),
                Date = reader.String(root, "$", "date", required: false),
                Author = reader.String(root, "$", "author", required: false),
                SectionNumber = reader.String(root, "$", "section_number", required: false),
                Points = reader.Array(root, "$", "points", (element, path) => ReadPoint(reader, element, path)),
                Columns = reader.Array(root, "$", "columns", (element, path) => ReadColumn(reader, element, path)),
                Cards = reader.Array(root, "$", "cards", (element, path) => ReadCard(reader, element, path)),
                Kpis = reader.Array(root, "$", "kpis", (element, path) => ReadKpi(reader, element, path)),
                Image = reader.Object(root, "$", "image", (element, path) => ReadImage(reader, element, path)),
                Table = reader.Object(root, "$", "table", (element, path) => ReadTable(reader, element, path)),
                Chart = reader.Object(root, "$", "chart", (element, path) => ReadChart(reader, element, path)),
                Milestones = reader.Array(root, "$", "milestones", (element, path) => ReadMilestone(reader, element, path)),
                Steps = reader.Array(root, "$", "steps", (element, path) => ReadStep(reader, element, path)),
                Tree = reader.Object(root, "$", "tree", (element, path) => ReadNode(reader, element, path, 1)),
                Quote = reader.Object(root, "$", "quote", (element, path) => ReadQuote(reader, element, path)),
                References = reader.Array(root, "$", "references", (element, path) => ReadReference(reader, element, path)),
                Fit = fit,
                Metadata = metadata,
                OriginalJson = root.GetRawText(),
            };

            if (kind is not null)
                CompositionKinds.Find(kind)!.Validate(spec, errors, warnings);

            return new CompositionParseResult { Spec = errors.Count == 0 ? spec : null, Errors = errors, Warnings = warnings };
        }
    }

    private static SpecFit ReadFit(JsonSpecReader reader, JsonElement root)
    {
        if (!root.TryGetProperty("fit", out var fit) || fit.ValueKind == JsonValueKind.Null)
            return new SpecFit(["shrink", "split"], null);
        if (fit.ValueKind != JsonValueKind.Object)
        {
            reader.Errors.Add("$.fit: must be an object, e.g. {\"policy\": [\"shrink\", \"split\"], \"min_font_size\": 12}.");
            return new SpecFit(["shrink", "split"], null);
        }
        var policy = reader.StringArray(fit, "$.fit", "policy") ?? ["shrink", "split"];
        foreach (var (value, index) in policy.Select((value, index) => (value, index)))
        {
            if (Array.IndexOf(FitPolicies, value) < 0)
                reader.Errors.Add($"$.fit.policy[{index}]: '{value}' must be one of {string.Join(", ", FitPolicies)}.");
        }
        if (policy.Contains("report") && policy.Count > 1)
            reader.Errors.Add("$.fit.policy: 'report' means no automatic fitting and cannot be combined with other policies.");
        var min = reader.Float(fit, "$.fit", "min_font_size", 6f, 40f);
        return new SpecFit(policy, min);
    }

    private static Dictionary<string, string>? ReadMetadata(JsonSpecReader reader, JsonElement root)
    {
        if (!root.TryGetProperty("metadata", out var metadata) || metadata.ValueKind == JsonValueKind.Null)
            return null;
        if (metadata.ValueKind != JsonValueKind.Object)
        {
            reader.Errors.Add("$.metadata: must be an object of string values.");
            return null;
        }
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in metadata.EnumerateObject())
        {
            if (property.Value.ValueKind != JsonValueKind.String)
                reader.Errors.Add($"$.metadata.{property.Name}: must be a string.");
            else if (!MetaKey().IsMatch(property.Name))
                reader.Errors.Add($"$.metadata.{property.Name}: keys use letters, digits, and underscore (max 40).");
            else
                result[property.Name] = property.Value.GetString()!;
        }
        return result;
    }

    private static SpecPoint? ReadPoint(JsonSpecReader reader, JsonElement element, string path)
    {
        if (element.ValueKind == JsonValueKind.String)
        {
            var text = element.GetString()!;
            if (string.IsNullOrWhiteSpace(text))
            {
                reader.Errors.Add($"{path}: empty point.");
                return null;
            }
            return new SpecPoint(text);
        }
        if (element.ValueKind != JsonValueKind.Object)
        {
            reader.Errors.Add($"{path}: a point is a string or {{\"text\": ..., \"detail\": ..., \"level\": 0}}.");
            return null;
        }
        reader.Unknown(element, path, "text", "detail", "level");
        var value = reader.String(element, path, "text", required: true);
        var detail = reader.String(element, path, "detail", required: false);
        var level = reader.Int(element, path, "level", 0, 4) ?? 0;
        return value is null ? null : new SpecPoint(value, detail, level);
    }

    private static SpecColumn? ReadColumn(JsonSpecReader reader, JsonElement element, string path)
    {
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "heading", "points", "tone");
        var heading = reader.String(element, path, "heading", required: true);
        var points = reader.Array(element, path, "points", (item, itemPath) => ReadPoint(reader, item, itemPath));
        var tone = reader.String(element, path, "tone", required: false);
        if (tone is not null && tone is not ("positive" or "negative" or "neutral"))
            reader.Errors.Add($"{path}.tone: use positive, negative, or neutral.");
        if (points is null or { Count: 0 })
            reader.Errors.Add($"{path}.points: at least one point is required.");
        return heading is null || points is null ? null : new SpecColumn(heading, points, tone);
    }

    private static SpecCard? ReadCard(JsonSpecReader reader, JsonElement element, string path)
    {
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "heading", "body", "badge");
        var heading = reader.String(element, path, "heading", required: true);
        return heading is null ? null : new SpecCard(heading, reader.String(element, path, "body", required: false), reader.String(element, path, "badge", required: false));
    }

    private static SpecKpi? ReadKpi(JsonSpecReader reader, JsonElement element, string path)
    {
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "value", "label", "delta", "trend", "sentiment", "note");
        var value = reader.String(element, path, "value", required: true);
        var label = reader.String(element, path, "label", required: true);
        var trend = reader.Enum(element, path, "trend", Trends);
        var sentiment = reader.Enum(element, path, "sentiment", Sentiments);
        return value is null || label is null
            ? null
            : new SpecKpi(value, label, reader.String(element, path, "delta", required: false), trend, reader.String(element, path, "note", required: false), sentiment);
    }

    private static SpecImage? ReadImage(JsonSpecReader reader, JsonElement element, string path)
    {
        reader.Unknown(element, path, "path", "alt", "fit", "focal_x", "focal_y", "position", "attribution");
        var file = reader.String(element, path, "path", required: true);
        var fit = reader.Enum(element, path, "fit", ImageFits) ?? "cover";
        var position = reader.Enum(element, path, "position", ImagePositions) ?? "left";
        var focalX = reader.Float(element, path, "focal_x", 0f, 1f) ?? 0.5f;
        var focalY = reader.Float(element, path, "focal_y", 0f, 1f) ?? 0.5f;
        var alt = reader.String(element, path, "alt", required: false);
        if (alt is null)
            reader.Warnings.Add($"{path}.alt: no alternative text; screen readers will not describe the image.");
        return file is null ? null : new SpecImage(file, alt, fit, focalX, focalY, position, reader.String(element, path, "attribution", required: false));
    }

    private static SpecTable? ReadTable(JsonSpecReader reader, JsonElement element, string path)
    {
        reader.Unknown(element, path, "header", "rows", "align", "total_row", "highlight", "column_weights", "repeat_header");
        var header = reader.StringArray(element, path + "", "header");
        if (header is null or { Count: 0 })
        {
            reader.Errors.Add($"{path}.header: required, one string per column.");
            return null;
        }
        if (header.Count > 12)
            reader.Errors.Add($"{path}.header: at most 12 columns fit on a slide (got {header.Count}).");

        var rows = new List<IReadOnlyList<string>>();
        if (!element.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array)
        {
            reader.Errors.Add($"{path}.rows: required array of rows (arrays of cell strings or numbers).");
            return null;
        }
        int rowIndex = 0;
        foreach (var row in rowsElement.EnumerateArray())
        {
            var rowPath = $"{path}.rows[{rowIndex}]";
            if (row.ValueKind != JsonValueKind.Array)
            {
                reader.Errors.Add($"{rowPath}: must be an array of cells.");
            }
            else
            {
                var cells = new List<string>();
                int cellIndex = 0;
                foreach (var cell in row.EnumerateArray())
                {
                    cells.Add(cell.ValueKind switch
                    {
                        JsonValueKind.String => cell.GetString()!,
                        // Numbers keep their JSON spelling; no locale formatting is applied.
                        JsonValueKind.Number => cell.GetRawText(),
                        JsonValueKind.Null => "",
                        JsonValueKind.True => "true",
                        JsonValueKind.False => "false",
                        _ => Bad($"{rowPath}[{cellIndex}]: cells are strings or numbers."),
                    });
                    cellIndex++;
                }
                if (cells.Count != header.Count)
                    reader.Errors.Add($"{rowPath}: has {cells.Count} cells but the header has {header.Count} columns. Nothing is padded or dropped.");
                rows.Add(cells);
            }
            rowIndex++;
        }
        if (rows.Count == 0)
            reader.Errors.Add($"{path}.rows: at least one row is required.");

        var align = reader.StringArray(element, path, "align");
        if (align is not null)
        {
            if (align.Count != header.Count)
                reader.Errors.Add($"{path}.align: needs one entry per column ({header.Count}).");
            foreach (var (value, index) in align.Select((value, index) => (value, index)).Where(entry => Array.IndexOf(Alignments, entry.value) < 0))
                reader.Errors.Add($"{path}.align[{index}]: '{value}' must be left, center, or right.");
        }
        var highlight = reader.IntArray(element, path, "highlight");
        foreach (var row in highlight ?? [])
        {
            if (row < 1 || row > rows.Count)
                reader.Errors.Add($"{path}.highlight: row {row} is outside 1-{rows.Count} (1-based body rows).");
        }
        var weights = reader.FloatArray(element, path, "column_weights");
        if (weights is not null && (weights.Count != header.Count || weights.Any(weight => weight <= 0)))
            reader.Errors.Add($"{path}.column_weights: needs {header.Count} positive numbers.");
        var total = reader.Bool(element, path, "total_row") ?? false;
        var repeat = reader.Bool(element, path, "repeat_header") ?? true;
        return new SpecTable(header, rows, align, total, highlight, weights, repeat);

        string Bad(string message)
        {
            reader.Errors.Add(message);
            return "";
        }
    }

    private static SpecChart? ReadChart(JsonSpecReader reader, JsonElement element, string path)
    {
        reader.Unknown(element, path, "type", "categories", "series", "number_format", "value_axis_title", "category_axis_title", "legend", "data_labels");
        var type = reader.Enum(element, path, "type", CompositionKinds.ChartTypes) ?? "column";
        var categories = reader.StringArray(element, path, "categories");
        if (categories is null or { Count: 0 })
        {
            reader.Errors.Add($"{path}.categories: required, one label per category.");
            return null;
        }
        var series = reader.Array(element, path, "series", (item, itemPath) =>
        {
            if (!reader.IsObject(item, itemPath))
                return null;
            reader.Unknown(item, itemPath, "name", "values");
            var name = reader.String(item, itemPath, "name", required: true);
            if (!item.TryGetProperty("values", out var valuesElement) || valuesElement.ValueKind != JsonValueKind.Array)
            {
                reader.Errors.Add($"{itemPath}.values: required array of numbers (null for a missing value).");
                return null;
            }
            var values = new List<double>();
            int index = 0;
            foreach (var value in valuesElement.EnumerateArray())
            {
                if (value.ValueKind == JsonValueKind.Number)
                    values.Add(value.GetDouble());
                else if (value.ValueKind == JsonValueKind.Null)
                    values.Add(double.NaN);
                else
                    reader.Errors.Add($"{itemPath}.values[{index}]: must be a number or null; text such as \"12%\" is not converted.");
                index++;
            }
            if (values.Count != categories.Count)
                reader.Errors.Add($"{itemPath}.values: has {values.Count} values but there are {categories.Count} categories.");
            return name is null ? null : new SpecSeries(name, values);
        });
        if (series is null or { Count: 0 })
        {
            reader.Errors.Add($"{path}.series: at least one series is required.");
            return null;
        }
        if (type is "pie" or "doughnut" && series.Count != 1)
            reader.Errors.Add($"{path}.series: a {type} chart takes exactly one series.");
        if (type is "pie" or "doughnut" && series.SelectMany(item => item.Values).Any(value => value < 0))
            reader.Errors.Add($"{path}.series: {type} charts cannot show negative values.");
        return new SpecChart(type, categories, series,
            reader.String(element, path, "number_format", required: false),
            reader.String(element, path, "value_axis_title", required: false),
            reader.String(element, path, "category_axis_title", required: false),
            reader.Bool(element, path, "legend"),
            reader.Bool(element, path, "data_labels"));
    }

    private static SpecMilestone? ReadMilestone(JsonSpecReader reader, JsonElement element, string path)
    {
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "date", "label", "detail", "status");
        var date = reader.String(element, path, "date", required: true);
        var label = reader.String(element, path, "label", required: true);
        var status = reader.Enum(element, path, "status", Statuses) ?? "planned";
        return date is null || label is null ? null : new SpecMilestone(date, label, reader.String(element, path, "detail", required: false), status);
    }

    private static SpecStep? ReadStep(JsonSpecReader reader, JsonElement element, string path)
    {
        if (element.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(element.GetString()))
            return new SpecStep(element.GetString()!, null);
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "label", "detail");
        var label = reader.String(element, path, "label", required: true);
        return label is null ? null : new SpecStep(label, reader.String(element, path, "detail", required: false));
    }

    private static SpecNode? ReadNode(JsonSpecReader reader, JsonElement element, string path, int depth)
    {
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "label", "detail", "children", "id");
        if (depth > 4)
        {
            reader.Errors.Add($"{path}: hierarchies deeper than 4 levels do not fit on one slide; split the tree.");
            return null;
        }
        var label = reader.String(element, path, "label", required: true);
        var children = reader.Array(element, path, "children", (item, itemPath) => ReadNode(reader, item, itemPath, depth + 1)) ?? [];
        return label is null ? null : new SpecNode(label, reader.String(element, path, "detail", required: false), children, reader.String(element, path, "id", required: false));
    }

    private static SpecQuote? ReadQuote(JsonSpecReader reader, JsonElement element, string path)
    {
        reader.Unknown(element, path, "text", "author", "role");
        var text = reader.String(element, path, "text", required: true);
        var author = reader.String(element, path, "author", required: true);
        return text is null || author is null ? null : new SpecQuote(text, author, reader.String(element, path, "role", required: false));
    }

    private static SpecReference? ReadReference(JsonSpecReader reader, JsonElement element, string path)
    {
        if (element.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(element.GetString()))
            return new SpecReference(element.GetString()!, null, null);
        if (!reader.IsObject(element, path))
            return null;
        reader.Unknown(element, path, "text", "label", "url");
        var text = reader.String(element, path, "text", required: true);
        return text is null ? null : new SpecReference(text, reader.String(element, path, "label", required: false), reader.String(element, path, "url", required: false));
    }

    [GeneratedRegex("^[A-Za-z0-9_.~-]{1,64}$", RegexOptions.CultureInvariant)]
    private static partial Regex AppId();

    [GeneratedRegex("^[A-Za-z0-9_]{1,40}$", RegexOptions.CultureInvariant)]
    private static partial Regex MetaKey();
}
