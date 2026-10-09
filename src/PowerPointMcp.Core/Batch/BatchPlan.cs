using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Sbroenne.PowerPointMcp.Core.Catalog;

namespace Sbroenne.PowerPointMcp.Core.Batch;

/// <summary>One operation of a batch.</summary>
public sealed class BatchOperation
{
    /// <summary>Operation id (from the input, or op1, op2, …).</summary>
    public required string Id { get; init; }

    /// <summary>Command category (tool name without _read).</summary>
    public required string Category { get; init; }

    /// <summary>Action.</summary>
    public required string Action { get; init; }

    /// <summary>Arguments with camelCase names; references to earlier results are kept as "$opId.path" strings.</summary>
    public required JsonObject Arguments { get; init; }

    /// <summary>References this operation makes: argument name to (operation id, result path).</summary>
    public required IReadOnlyDictionary<string, (string Operation, string Path)> References { get; init; }

    /// <summary>category.action.</summary>
    public string Command => $"{Category}.{Action}";
}

/// <summary>A parsed and prevalidated batch.</summary>
public sealed class BatchPlan
{
    /// <summary>Operations in order.</summary>
    public required IReadOnlyList<BatchOperation> Operations { get; init; }

    /// <summary>Problems that prevent running (op id: message).</summary>
    public required IReadOnlyList<string> Errors { get; init; }

    /// <summary>Notes (removed session ids, renamed arguments).</summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}

/// <summary>
/// Parses batch JSON and prevalidates every operation against the generated command registry
/// before anything runs. Pure apart from reading the registry.
/// </summary>
/// <remarks>
/// Input: a JSON array (or {"operations": [...]}) of objects with <c>tool</c> (or <c>category</c>)
/// and <c>action</c>, or <c>command</c> "tool.action"; optional <c>id</c> and <c>args</c>. Argument
/// names may be snake_case (MCP style) or camelCase. A string argument written exactly as
/// <c>$op1.shapeIndex</c> (or <c>$op2.slides[0].slideId</c>) is replaced at run time with that value
/// from an earlier operation's result.
/// </remarks>
public static partial class BatchPlanner
{
    /// <summary>Most operations per batch.</summary>
    public const int MaxOperations = 500;

    // Arguments go to the in-process service only; keep non-ASCII text readable.
    private static readonly JsonSerializerOptions JsonOptions = new() { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    private static readonly HashSet<string> Forbidden = new(StringComparer.Ordinal) { "batch", "service", "session" };

    /// <summary>Parses and prevalidates.</summary>
    public static BatchPlan Parse(string json)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        var operations = new List<BatchOperation>();
        JsonNode? root;
        try
        {
            root = JsonNode.Parse(json ?? "", documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
        }
        catch (JsonException ex)
        {
            return new BatchPlan { Operations = [], Errors = [$"$: invalid JSON: {ex.Message}"], Warnings = [] };
        }

        var array = root switch
        {
            JsonArray items => items,
            JsonObject wrapper when wrapper["operations"] is JsonArray items => items,
            _ => null,
        };
        if (array is null)
            return new BatchPlan { Operations = [], Errors = ["$: expected a JSON array of operations (or {\"operations\": [...]})."], Warnings = [] };
        if (array.Count == 0)
            return new BatchPlan { Operations = [], Errors = ["$: the batch has no operations."], Warnings = [] };
        if (array.Count > MaxOperations)
            return new BatchPlan { Operations = [], Errors = [$"$: {array.Count} operations exceed the limit of {MaxOperations}; split the batch."], Warnings = [] };

        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < array.Count; index++)
        {
            var path = $"$[{index}]";
            if (array[index] is not JsonObject item)
            {
                errors.Add($"{path}: each operation is an object.");
                continue;
            }
            foreach (var key in item.Select(property => property.Key).Where(key => key is not ("id" or "tool" or "category" or "action" or "command" or "args")))
                errors.Add($"{path}.{key}: unknown field; use id, tool, action, args (or command).");

            var id = item["id"]?.GetValueKind() == JsonValueKind.String ? item["id"]!.GetValue<string>() : $"op{index + 1}";
            if (!OpId().IsMatch(id))
            {
                errors.Add($"{path}.id: use letters, digits, dash, or underscore.");
                continue;
            }
            if (!seen.Add(id))
            {
                errors.Add($"{path}.id: '{id}' is used twice.");
                continue;
            }

            string? category = Text(item, "tool") ?? Text(item, "category");
            string? action = Text(item, "action");
            if (Text(item, "command") is { } command)
            {
                var parts = command.Split('.', 2);
                if (parts.Length != 2)
                {
                    errors.Add($"{id}: command must be tool.action, e.g. shape.add-text-box.");
                    continue;
                }
                (category, action) = (parts[0], parts[1]);
            }
            if (category is null || action is null)
            {
                errors.Add($"{id}: needs tool and action (or command).");
                continue;
            }
            if (category.EndsWith("_read", StringComparison.Ordinal))
                category = category[..^"_read".Length];
            if (Forbidden.Contains(category) || category == "presentation")
            {
                errors.Add($"{id}: {category} operations cannot run inside a batch (session lifecycle and nested batches are excluded).");
                continue;
            }

            var arguments = new JsonObject();
            var references = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
            if (item["args"] is JsonObject args)
            {
                foreach (var (name, value) in args)
                {
                    var canonical = CamelCase(name);
                    if (canonical is "presentationSessionId" or "sessionId")
                    {
                        warnings.Add($"{id}: {name} removed; batch operations run in the batch's own session.");
                        continue;
                    }
                    if (value is JsonValue text && text.GetValueKind() == JsonValueKind.String && Reference().Match(text.GetValue<string>()) is { Success: true } match)
                    {
                        var target = match.Groups["op"].Value;
                        if (!operations.Any(operation => operation.Id == target))
                            errors.Add($"{id}.args.{name}: refers to '{target}', which is not an earlier operation.");
                        references[canonical] = (target, match.Groups["path"].Value);
                    }
                    arguments[canonical] = value?.DeepClone();
                }
            }
            else if (item["args"] is not null)
            {
                errors.Add($"{id}.args: must be an object.");
            }

            // Referenced values exist only at run time: probe them as a number, then text, then a boolean.
            if (Prevalidate(category, action, arguments, references.Keys) is { } problem)
                errors.Add($"{id} ({category}.{action}): {problem}");

            operations.Add(new BatchOperation { Id = id, Category = category, Action = action, Arguments = arguments, References = references });
        }

        return new BatchPlan { Operations = operations, Errors = errors, Warnings = warnings };
    }

    private static string? Prevalidate(string category, string action, JsonObject arguments, IEnumerable<string> referenced)
    {
        var names = referenced.ToList();
        string? first = null;
        foreach (JsonNode placeholder in new JsonNode[] { JsonValue.Create(1), JsonValue.Create("x"), JsonValue.Create(true) })
        {
            var probe = (JsonObject)arguments.DeepClone();
            foreach (var name in names)
                probe[name] = placeholder.DeepClone();
            var problem = CommandCatalog.ValidateArguments(category, action, probe.ToJsonString());
            if (problem is null)
                return null;
            first ??= problem;
            if (names.Count == 0)
                break;
        }
        return first;
    }

    /// <summary>Resolves "$op.path" references against earlier results; returns the arguments JSON or an error.</summary>
    public static (string? ArgumentsJson, string? Error) Resolve(BatchOperation operation, IReadOnlyDictionary<string, JsonNode?> results)
    {
        ArgumentNullException.ThrowIfNull(operation);
        ArgumentNullException.ThrowIfNull(results);
        var arguments = (JsonObject)operation.Arguments.DeepClone();
        foreach (var (name, (target, path)) in operation.References)
        {
            if (!results.TryGetValue(target, out var result) || result is null)
                return (null, $"argument {name}: operation '{target}' has no result.");
            var value = Select(result, path);
            if (value is null)
                return (null, $"argument {name}: '{target}' result has no value at '{path}'.");
            arguments[name] = value.DeepClone();
        }
        return (arguments.ToJsonString(JsonOptions), null);
    }

    /// <summary>Selects a value by a dotted path with optional [n] indexes, e.g. slides[0].slideId.</summary>
    public static JsonNode? Select(JsonNode node, string path)
    {
        JsonNode? current = node;
        foreach (Match part in PathPart().Matches(path))
        {
            if (current is null)
                return null;
            if (part.Groups["name"].Success)
                current = current is JsonObject obj ? obj.FirstOrDefault(property => string.Equals(property.Key, part.Groups["name"].Value, StringComparison.OrdinalIgnoreCase)).Value : null;
            else
                current = current is JsonArray list && int.TryParse(part.Groups["index"].Value, out var index) && index < list.Count ? list[index] : null;
        }
        return current;
    }

    /// <summary>snake_case or kebab-case to camelCase; camelCase is unchanged.</summary>
    public static string CamelCase(string name)
    {
        if (!name.Contains('_', StringComparison.Ordinal) && !name.Contains('-', StringComparison.Ordinal))
            return name;
        var builder = new StringBuilder(name.Length);
        bool upper = false;
        foreach (var character in name)
        {
            if (character is '_' or '-')
            {
                upper = builder.Length > 0;
                continue;
            }
            builder.Append(upper ? char.ToUpperInvariant(character) : character);
            upper = false;
        }
        return builder.ToString();
    }

    private static string? Text(JsonObject item, string name) =>
        item[name] is JsonValue value && value.GetValueKind() == JsonValueKind.String ? value.GetValue<string>() : null;

    [GeneratedRegex("^[A-Za-z0-9_-]{1,40}$", RegexOptions.CultureInvariant)]
    private static partial Regex OpId();

    [GeneratedRegex(@"^\$(?<op>[A-Za-z0-9_-]{1,40})\.(?<path>[A-Za-z0-9_]+(\[[0-9]+\])*(\.[A-Za-z0-9_]+(\[[0-9]+\])*)*)$", RegexOptions.CultureInvariant)]
    private static partial Regex Reference();

    [GeneratedRegex(@"(?<name>[A-Za-z0-9_]+)|\[(?<index>[0-9]+)\]", RegexOptions.CultureInvariant)]
    private static partial Regex PathPart();
}
