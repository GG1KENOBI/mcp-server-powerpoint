using System.Collections.Concurrent;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ModelContextProtocol.Protocol;
using Sbroenne.PowerPointMcp.McpServer.Tools;

namespace Sbroenne.PowerPointMcp.McpServer;

/// <summary>
/// Tolerates near-miss tool calls (typically from small/local models) without changing any
/// schema: action names and parameter names are matched to the tool's canonical names by a
/// separator- and case-insensitive key, but only when exactly one canonical name matches.
/// Everything else is rejected with an error that names the closest valid action/parameter,
/// the missing parameters, and a compact example call.
/// </summary>
internal static partial class ToolCallNormalizer
{
    private static readonly ConcurrentDictionary<string, ToolCallMetadata> MetadataCache = new(StringComparer.Ordinal);

    /// <summary>
    /// When true, near-miss names are rewritten to the canonical name. When false (the default,
    /// matching upstream's strict contract) they are rejected with a "Did you mean" suggestion.
    /// Enabled by setting the PPTMCP_LENIENT_ARGUMENTS environment variable to 1 or true, which
    /// is intended for small local models.
    /// </summary>
    internal static bool LenientNames { get; set; } = IsEnabled(Environment.GetEnvironmentVariable("PPTMCP_LENIENT_ARGUMENTS"));

    internal static bool IsEnabled(string? value) =>
        value is not null && (value.Trim() == "1" || value.Trim().Equals("true", StringComparison.OrdinalIgnoreCase));

    // Keeps <presentation_session_id> and Windows paths readable in example calls (no \u003C escapes).
    private static readonly JsonSerializerOptions ExampleJsonOptions = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    internal static ToolCallMetadata GetMetadata(Tool tool) =>
        MetadataCache.GetOrAdd(tool.Name, _ => ToolCallMetadata.Create(tool));

    /// <summary>
    /// Resolves the canonical action and parameter names. Returns the canonical action and the
    /// arguments keyed by canonical names (the same instance when nothing needed renaming).
    /// </summary>
    internal static (string Action, IDictionary<string, JsonElement> Arguments) Normalize(
        ToolCallMetadata metadata,
        IDictionary<string, JsonElement>? arguments)
    {
        if (arguments == null)
            throw new ArgumentException($"The action argument is required. {metadata.DescribeActions()}");
        if (!arguments.TryGetValue("action", out var actionValue) || actionValue.ValueKind != JsonValueKind.String)
            throw new ArgumentException($"The action argument must be a string naming an available action. {metadata.DescribeActions()}");

        var requestedAction = actionValue.GetString()!;
        var action = (LenientNames ? metadata.ResolveAction(requestedAction) : metadata.ResolveActionIgnoringCase(requestedAction))
            ?? throw new ArgumentException(metadata.DescribeUnknownAction(requestedAction));

        var renamed = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var name in arguments.Keys)
        {
            if (name == "action" || metadata.HasParameter(name))
                continue;

            var canonical = LenientNames ? metadata.ResolveParameter(name) : null;
            if (canonical == null)
                throw new ArgumentException(metadata.DescribeUnknownParameter(name, action));
            if (arguments.ContainsKey(canonical) || renamed.ContainsValue(canonical))
                throw new ArgumentException(
                    $"Parameter '{name}' duplicates '{canonical}'. Send only '{canonical}'.");
            renamed[name] = canonical;
        }

        if (renamed.Count == 0 && string.Equals(requestedAction, action, StringComparison.Ordinal))
            return (action, arguments);

        var normalized = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var (name, value) in arguments)
        {
            normalized[renamed.GetValueOrDefault(name, name)] =
                name == "action" ? JsonSerializer.SerializeToElement(action) : value;
        }
        return (action, normalized);
    }

    /// <summary>Rejects the call when parameters the schema marks required for the action are missing.</summary>
    internal static void EnsureRequiredParameters(
        ToolCallMetadata metadata,
        string action,
        IDictionary<string, JsonElement> arguments)
    {
        var required = metadata.RequiredParameters(action);
        var missing = required
            .Where(name => !arguments.TryGetValue(name, out var value) || value.ValueKind == JsonValueKind.Null)
            .ToArray();
        if (missing.Length == 0)
            return;

        var subject = missing.Length == 1
            ? $"{missing[0]} is required"
            : $"{string.Join(", ", missing)} are required";
        throw new ArgumentException(
            $"{subject} for {metadata.ToolName} action \"{action}\". " +
            $"Required parameters for \"{action}\": {string.Join(", ", required)}. " +
            $"Example: {metadata.ExampleCall(action)}");
    }

    /// <summary>Adds the action's valid parameters to an existing "not valid for action" error.</summary>
    internal static string DescribeInapplicableParameters(ToolCallMetadata metadata, string action, string message) =>
        $"{message} Valid parameters for \"{action}\": {string.Join(", ", metadata.ValidParameters(action))}.";

    /// <summary>Lowercase letters and digits only: add-text-box, add_text_box and AddTextBox share a key.</summary>
    internal static string Fold(string name)
    {
        var builder = new StringBuilder(name.Length);
        foreach (var character in name)
        {
            if (char.IsLetterOrDigit(character))
                builder.Append(char.ToLowerInvariant(character));
        }
        return builder.ToString();
    }

    internal static IReadOnlyList<string> Suggest(string requested, IEnumerable<string> candidates, int max = 3)
    {
        var key = Fold(requested);
        if (key.Length == 0)
            return [];

        var threshold = Math.Max(2, key.Length / 4);
        return candidates
            .Select(candidate => (Candidate: candidate, Key: Fold(candidate)))
            .Select(item => (item.Candidate, Distance: key.Length >= 4 && (item.Key.Contains(key, StringComparison.Ordinal) || key.Contains(item.Key, StringComparison.Ordinal))
                ? Math.Min(1, Levenshtein(key, item.Key))
                : Levenshtein(key, item.Key)))
            .Where(item => item.Distance <= threshold)
            .OrderBy(item => item.Distance)
            .ThenBy(item => item.Candidate, StringComparer.Ordinal)
            .Take(max)
            .Select(item => item.Candidate)
            .ToArray();
    }

    private static int Levenshtein(string a, string b)
    {
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
            previous[j] = j;
        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
            }
            (previous, current) = (current, previous);
        }
        return previous[b.Length];
    }

    [GeneratedRegex(@"\((required for|valid for):\s*([^)]*)\)", RegexOptions.IgnoreCase)]
    private static partial Regex GeneratedMarker();

    [GeneratedRegex(@"\([^)]*\)")]
    private static partial Regex Parenthetical();

    [GeneratedRegex(@"(Required for|Used for):\s*([^.]*)\.", RegexOptions.IgnoreCase)]
    private static partial Regex ProseMarker();

    [GeneratedRegex(@",|\s+or\s+")]
    private static partial Regex ListSeparator();

    /// <summary>Per-tool facts derived once from the tool's published input schema.</summary>
    internal sealed class ToolCallMetadata
    {
        private readonly Dictionary<string, string?> _actionsByKey = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string?> _parametersByKey = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _parameterTypes = new(StringComparer.Ordinal);
        private readonly Dictionary<string, List<string>> _required = new(StringComparer.Ordinal);
        private readonly Dictionary<string, List<string>> _valid = new(StringComparer.Ordinal);

        private const string SessionParameter = "presentation_session_id";

        private static readonly HashSet<string> SessionAliases = new(StringComparer.Ordinal)
        {
            "sessionid", "session", "presentationid", "presentationsession",
        };

        private ToolCallMetadata(string toolName) => ToolName = toolName;

        internal string ToolName { get; }

        internal IReadOnlyList<string> Actions { get; private set; } = [];

        internal IReadOnlyList<string> Parameters { get; private set; } = [];

        internal static ToolCallMetadata Create(Tool tool)
        {
            var metadata = new ToolCallMetadata(tool.Name);
            var schema = tool.InputSchema;
            var properties = schema.GetProperty("properties");

            metadata.Actions = properties.TryGetProperty("action", out var actionSchema)
                && actionSchema.TryGetProperty("enum", out var actionEnum)
                    ? actionEnum.EnumerateArray().Select(value => value.GetString()!).ToArray()
                    : [];
            foreach (var action in metadata.Actions)
            {
                AddUnique(metadata._actionsByKey, Fold(action), action);
                metadata._required[action] = [];
                metadata._valid[action] = [];
            }

            var schemaRequired = schema.TryGetProperty("required", out var requiredArray)
                ? requiredArray.EnumerateArray().Select(value => value.GetString()!).ToHashSet(StringComparer.Ordinal)
                : [];

            var parameters = new List<string>();
            foreach (var property in properties.EnumerateObject())
            {
                if (property.Name == "action")
                    continue;
                parameters.Add(property.Name);
                AddUnique(metadata._parametersByKey, Fold(property.Name), property.Name);
                metadata._parameterTypes[property.Name] = PrimaryType(property.Value);
                metadata.AddApplicability(property.Name, property.Value, schemaRequired.Contains(property.Name));
            }
            metadata.Parameters = parameters;
            return metadata;
        }

        internal string? ResolveAction(string requested) =>
            Actions.Contains(requested, StringComparer.Ordinal)
                ? requested
                : _actionsByKey.GetValueOrDefault(Fold(requested));

        internal string? ResolveActionIgnoringCase(string requested) =>
            Actions.FirstOrDefault(action => string.Equals(action, requested, StringComparison.OrdinalIgnoreCase));

        internal bool HasParameter(string name) => _parameterTypes.ContainsKey(name);

        internal string? ResolveParameter(string requested)
        {
            if (_parametersByKey.GetValueOrDefault(Fold(requested)) is { } canonical)
                return canonical;

            // Every tool names its session parameter presentation_session_id; models often send the
            // shorter forms. Map them only when the tool actually has that parameter.
            return SessionAliases.Contains(Fold(requested)) && HasParameter(SessionParameter) ? SessionParameter : null;
        }

        internal IReadOnlyList<string> RequiredParameters(string action) =>
            _required.TryGetValue(action, out var required) ? required : [];

        internal IReadOnlyList<string> ValidParameters(string action)
        {
            if (ToolName == "presentation")
            {
                var allowed = PresentationTools.GetAllowedParameterNames(action);
                return Parameters.Where(allowed.Contains).ToArray();
            }
            return _valid.TryGetValue(action, out var valid) ? valid : [];
        }

        internal string DescribeActions() => $"Valid actions: {string.Join(", ", Actions)}.";

        internal string DescribeUnknownAction(string requested)
        {
            List<string> parts = [$"Unknown action \"{requested}\" for tool \"{ToolName}\"."];
            IReadOnlyList<string> suggestions = ResolveAction(requested) is { } alias ? [alias] : Suggest(requested, Actions);
            if (suggestions.Count > 0)
            {
                parts.Add($"Did you mean {string.Join(" or ", suggestions.Select(s => $"\"{s}\""))}?");
                parts.Add($"Required parameters for \"{suggestions[0]}\": {string.Join(", ", RequiredParameters(suggestions[0]))}.");
            }
            parts.Add(DescribeActions());
            return string.Join(' ', parts);
        }

        internal string DescribeUnknownParameter(string requested, string action)
        {
            List<string> parts = [$"Unknown parameter '{requested}' for {ToolName} action \"{action}\"."];
            var valid = ValidParameters(action);
            IReadOnlyList<string> suggestions = ResolveParameter(requested) is { } alias ? [alias] : Suggest(requested, valid.Count > 0 ? valid : Parameters);
            if (suggestions.Count > 0)
                parts.Add($"Did you mean {string.Join(" or ", suggestions.Select(s => $"'{s}'"))}?");
            parts.Add($"Valid parameters for \"{action}\": {string.Join(", ", valid)}.");
            return string.Join(' ', parts);
        }

        internal string ExampleCall(string action)
        {
            var example = new JsonObject { ["action"] = action };
            foreach (var name in RequiredParameters(action))
                example[name] = ExampleValue(name, _parameterTypes.GetValueOrDefault(name, "string"));
            return example.ToJsonString(ExampleJsonOptions);
        }

        private void AddApplicability(string name, JsonElement schema, bool requiredForAll)
        {
            var description = schema.TryGetProperty("description", out var value) ? value.GetString() ?? "" : "";
            var requiredFor = new List<string>();
            var validFor = new List<string>();
            var hasMarker = false;

            if (description.Contains("(required)", StringComparison.OrdinalIgnoreCase))
            {
                requiredForAll = true;
            }
            foreach (Match match in GeneratedMarker().Matches(description))
            {
                hasMarker = true;
                var target = match.Groups[1].Value.StartsWith("required", StringComparison.OrdinalIgnoreCase) ? requiredFor : validFor;
                target.AddRange(SplitActions(match.Groups[2].Value, prose: false));
            }
            if (!hasMarker)
            {
                // Hand-written (presentation) descriptions: "Required for: create (...), open or test (...)."
                foreach (Match match in ProseMarker().Matches(Parenthetical().Replace(description, "")))
                {
                    hasMarker = true;
                    var target = match.Groups[1].Value.StartsWith("required", StringComparison.OrdinalIgnoreCase) ? requiredFor : validFor;
                    target.AddRange(SplitActions(match.Groups[2].Value, prose: true));
                }
            }

            foreach (var action in Actions)
            {
                var required = requiredForAll || requiredFor.Contains(action, StringComparer.Ordinal);
                if (required)
                    _required[action].Add(name);
                if (required || !hasMarker || validFor.Contains(action, StringComparer.Ordinal))
                    _valid[action].Add(name);
            }
        }

        private IEnumerable<string> SplitActions(string list, bool prose) =>
            (prose ? ListSeparator().Split(list) : list.Split(','))
                .Select(item => item.Trim())
                .Where(item => Actions.Contains(item, StringComparer.Ordinal));

        private static void AddUnique(Dictionary<string, string?> map, string key, string value)
        {
            // A key shared by two canonical names is ambiguous and never resolves.
            map[key] = map.ContainsKey(key) ? null : value;
        }

        private static string PrimaryType(JsonElement schema)
        {
            if (!schema.TryGetProperty("type", out var type))
                return "string";
            return type.ValueKind == JsonValueKind.Array
                ? type.EnumerateArray().Select(item => item.GetString()).FirstOrDefault(item => item != "null") ?? "string"
                : type.GetString() ?? "string";
        }

        private static JsonNode ExampleValue(string name, string type)
        {
            if (name.Contains("session", StringComparison.OrdinalIgnoreCase))
                return "<presentation_session_id>";
            return type switch
            {
                "integer" => 1,
                "number" => 100,
                "boolean" => true,
                "array" => new JsonArray(1, 2),
                _ when name.Contains("path", StringComparison.OrdinalIgnoreCase) => @"C:\path\file.pptx",
                _ when name == "text" => "Hello",
                _ => $"<{name}>",
            };
        }
    }
}
