using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// A parsed semantic selector. All conditions must match (logical AND). Pure: evaluated against
/// inspection snapshots, never against COM.
/// </summary>
/// <remarks>
/// <para>Syntax: space-separated terms. Values with spaces go in double quotes. <c>*</c> and
/// <c>?</c> are wildcards; without wildcards <c>text:</c> and <c>title:</c> match a substring and
/// every other string key matches the whole value. Matching ignores case.</para>
/// <para>Slide keys: <c>slide:3</c> (position), <c>slideid:257</c>, <c>title:"Revenue*"</c>,
/// <c>slideappid:intro</c>, <c>slidetag:NAME=VALUE</c>, <c>section:"Appendix"</c>,
/// <c>layout:"Title Only"</c>.</para>
/// <para>Object keys: <c>id:5</c> (Shape.Id), <c>index:2</c> (shape position), <c>appid:kpi-1</c>,
/// <c>name:"Title 1"</c>, <c>role:title</c>, <c>kind:table</c>, <c>placeholder:body</c>,
/// <c>text:"grew"</c>, <c>tag:NAME=VALUE</c>, <c>component:card</c>, <c>group:"Cards"</c> (objects
/// inside the named group, or <c>group:12</c> by Shape.Id), and the flags <c>hidden</c>,
/// <c>visible</c>, <c>has-text</c>.</para>
/// <para>Numeric comparisons (points): <c>left</c>, <c>top</c>, <c>width</c>, <c>height</c>,
/// <c>right</c>, <c>bottom</c>, <c>rotation</c>, <c>z</c>, <c>font</c> (smallest font size),
/// with <c>&lt;</c>, <c>&lt;=</c>, <c>&gt;</c>, <c>&gt;=</c>, or <c>=</c>, e.g. <c>font&lt;12</c>.</para>
/// </remarks>
public sealed partial class ObjectSelector
{
    private static readonly HashSet<string> SlideKeys = new(StringComparer.Ordinal)
    {
        "slide", "slideid", "title", "slideappid", "slidetag", "section", "layout",
    };

    private static readonly HashSet<string> ObjectStringKeys = new(StringComparer.Ordinal)
    {
        "id", "index", "appid", "name", "role", "kind", "placeholder", "text", "tag", "component", "group",
    };

    private static readonly HashSet<string> NumericKeys = new(StringComparer.Ordinal)
    {
        "left", "top", "width", "height", "right", "bottom", "rotation", "z", "font",
    };

    private static readonly HashSet<string> Flags = new(StringComparer.Ordinal) { "hidden", "visible", "has-text" };

    private readonly List<Condition> _conditions;

    private ObjectSelector(string source, List<Condition> conditions)
    {
        Source = source;
        _conditions = conditions;
    }

    /// <summary>The selector text as supplied.</summary>
    public string Source { get; }

    /// <summary>Whether any condition constrains objects (not only slides).</summary>
    public bool HasObjectConditions => _conditions.Any(condition => !condition.IsSlideCondition);

    /// <summary>Parses selector text; throws <see cref="ArgumentException"/> describing the first problem.</summary>
    public static ObjectSelector Parse(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var conditions = new List<Condition>();
        foreach (var token in Tokenize(text))
        {
            conditions.Add(ParseTerm(token));
        }
        if (conditions.Count == 0)
            throw new ArgumentException("The selector is empty. Example: kind:table slide:3, or text:\"Revenue\" role:body.");
        return new ObjectSelector(text, conditions);
    }

    /// <summary>Whether the slide satisfies every slide-level condition.</summary>
    public bool MatchesSlide(DeckSlideInfo slide) =>
        _conditions.Where(condition => condition.IsSlideCondition).All(condition => condition.Matches(slide, null, []));

    /// <summary>Whether the object (on its slide, with its slide's objects for group lookups) matches.</summary>
    public bool Matches(DeckSlideInfo slide, DeckObjectInfo item, IReadOnlyList<DeckObjectInfo> slideObjects) =>
        _conditions.All(condition => condition.Matches(slide, item, slideObjects));

    /// <summary>Evaluates the selector over snapshots, in slide order then shape order.</summary>
    public IReadOnlyList<DeckObjectInfo> Find(IEnumerable<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> slides)
    {
        ArgumentNullException.ThrowIfNull(slides);
        var matches = new List<DeckObjectInfo>();
        foreach (var (slide, objects) in slides)
        {
            if (!MatchesSlide(slide))
                continue;
            matches.AddRange(objects.Where(item => Matches(slide, item, objects)));
        }
        return matches;
    }

    /// <summary>
    /// Requires exactly one match. Throws <see cref="ArgumentException"/> listing up to ten
    /// candidates when several objects match, so a caller can never edit an arbitrary one.
    /// </summary>
    public static DeckObjectInfo RequireSingle(string selector, IReadOnlyList<DeckObjectInfo> matches, int scannedObjects)
    {
        ArgumentNullException.ThrowIfNull(matches);
        if (matches.Count == 1)
            return matches[0];
        if (matches.Count == 0)
        {
            throw new ArgumentException(
                $"Selector '{selector}' matched none of the {scannedObjects} object(s) scanned. Inspect the slide with deck inspect-objects to see names, roles, and ids.");
        }

        var candidates = string.Join("; ", matches.Take(10).Select(Describe));
        var more = matches.Count > 10 ? $" (and {matches.Count - 10} more)" : "";
        throw new ArgumentException(
            $"Selector '{selector}' is ambiguous: it matched {matches.Count} objects: {candidates}{more}. " +
            "Nothing was changed. Narrow it with id:, appid:, name:, or slide:, or pass slide_id and shape_id.");
    }

    /// <summary>Compact one-line description of an object for candidate lists.</summary>
    public static string Describe(DeckObjectInfo item)
    {
        ArgumentNullException.ThrowIfNull(item);
        var text = item.Text is { Length: > 0 } value ? $" \"{Shorten(value, 40)}\"" : "";
        var appId = item.AppId is { Length: > 0 } id ? $" appid:{id}" : "";
        return $"slide {item.SlideIndex} (slideid:{item.SlideId}) id:{item.ShapeId}{appId} {item.Kind} '{item.Name}'{text}";
    }

    private static string Shorten(string value, int length)
    {
        var flat = value.Replace('\r', ' ').Replace('\n', ' ').Replace('\v', ' ');
        return flat.Length <= length ? flat : flat[..length] + "…";
    }

    private static List<string> Tokenize(string text)
    {
        var tokens = new List<string>();
        var current = new StringBuilder();
        var quoted = false;
        foreach (var character in text)
        {
            if (character == '"')
            {
                quoted = !quoted;
                current.Append(character);
            }
            else if (char.IsWhiteSpace(character) && !quoted)
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
            }
            else
            {
                current.Append(character);
            }
        }
        if (quoted)
            throw new ArgumentException($"Selector has an unterminated quote: {text}");
        if (current.Length > 0)
            tokens.Add(current.ToString());
        return tokens;
    }

    private static Condition ParseTerm(string token)
    {
        if (Flags.Contains(token.ToLowerInvariant()))
            return new Condition(token.ToLowerInvariant(), "=", "", null);

        var numeric = NumericTerm().Match(token);
        if (numeric.Success && NumericKeys.Contains(numeric.Groups["key"].Value.ToLowerInvariant()))
        {
            if (!double.TryParse(numeric.Groups["value"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                throw new ArgumentException($"Selector term '{token}' needs a number after {numeric.Groups["op"].Value}.");
            return new Condition(numeric.Groups["key"].Value.ToLowerInvariant(), numeric.Groups["op"].Value, "", number);
        }

        var colon = token.IndexOf(':', StringComparison.Ordinal);
        if (colon <= 0)
            throw new ArgumentException(
                $"Selector term '{token}' is not key:value, a comparison such as font<12, or a flag (hidden, visible, has-text).");

        var key = token[..colon].ToLowerInvariant();
        var value = token[(colon + 1)..];
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            value = value[1..^1];
        if (value.Length == 0)
            throw new ArgumentException($"Selector term '{token}' has no value.");
        if (!SlideKeys.Contains(key) && !ObjectStringKeys.Contains(key))
        {
            throw new ArgumentException(
                $"Unknown selector key '{key}'. Slide keys: {string.Join(", ", SlideKeys)}. Object keys: {string.Join(", ", ObjectStringKeys)}. Comparisons: {string.Join(", ", NumericKeys)}.");
        }
        if (key is "slide" or "slideid" or "id" or "index" && !int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out _))
            throw new ArgumentException($"Selector term '{token}' needs a whole number.");
        if (key is "tag" or "slidetag" && !value.Contains('=', StringComparison.Ordinal))
            throw new ArgumentException($"Selector term '{token}' must be {key}:NAME=VALUE.");
        return new Condition(key, ":", value, null);
    }

    [GeneratedRegex(@"^(?<key>[A-Za-z]+)(?<op><=|>=|<|>|=)(?<value>-?[0-9]+(\.[0-9]+)?)$")]
    private static partial Regex NumericTerm();

    private sealed class Condition(string key, string op, string value, double? number)
    {
        public bool IsSlideCondition => SlideKeys.Contains(key);

        public bool Matches(DeckSlideInfo slide, DeckObjectInfo? item, IReadOnlyList<DeckObjectInfo> slideObjects)
        {
            switch (key)
            {
                case "slide": return slide.SlideIndex == Int(value);
                case "slideid": return slide.SlideId == Int(value);
                case "title": return TextMatch(slide.Title, value);
                case "slideappid": return Exact(slide.AppId, value);
                case "section": return Exact(slide.SectionName, value);
                case "layout": return Exact(slide.LayoutName, value);
                case "slidetag": return TagMatch(slide.Tags, slide.AppId, null, value);
            }

            if (item is null)
                return true;

            switch (key)
            {
                case "id": return item.ShapeId == Int(value);
                case "index": return item.ShapeIndex == Int(value);
                case "appid": return Exact(item.AppId, value);
                case "name": return Exact(item.Name, value);
                case "role": return Exact(item.Role, value);
                case "kind": return Exact(item.Kind, value);
                case "placeholder": return Exact(PlaceholderShortName(item.PlaceholderType), value) || Exact(item.PlaceholderType, value);
                case "text": return TextMatch(item.Text, value);
                case "tag": return TagMatch(item.Tags, item.AppId, item.Component, value);
                case "component": return Exact(item.Component?.Split('@')[0], value) || Exact(item.Component, value);
                case "group": return GroupMatch(item, slideObjects, value);
                case "hidden": return !item.Visible;
                case "visible": return item.Visible;
                case "has-text": return item.HasText;
            }

            double? actual = key switch
            {
                "left" => item.Left,
                "top" => item.Top,
                "width" => item.Width,
                "height" => item.Height,
                "right" => item.Left + item.Width,
                "bottom" => item.Top + item.Height,
                "rotation" => item.Rotation,
                "z" => item.ZOrder,
                "font" => item.MinFontSize,
                _ => null,
            };
            if (actual is not { } a || number is not { } n)
                return false;
            return op switch
            {
                "<" => a < n,
                "<=" => a <= n + 0.001,
                ">" => a > n,
                ">=" => a >= n - 0.001,
                _ => Math.Abs(a - n) < 0.5,
            };
        }

        private static int Int(string text) => int.Parse(text, CultureInfo.InvariantCulture);

        private static bool Exact(string? actual, string pattern) =>
            actual is not null && (HasWildcard(pattern)
                ? Wildcard(pattern).IsMatch(actual)
                : string.Equals(actual, pattern, StringComparison.OrdinalIgnoreCase));

        private static bool TextMatch(string? actual, string pattern)
        {
            if (actual is null)
                return false;
            var flat = actual.Replace('\r', ' ').Replace('\n', ' ').Replace('\v', ' ');
            return HasWildcard(pattern)
                ? Wildcard(pattern).IsMatch(flat)
                : flat.Contains(pattern, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TagMatch(IReadOnlyDictionary<string, string>? tags, string? appId, string? component, string pattern)
        {
            var equals = pattern.IndexOf('=', StringComparison.Ordinal);
            var name = pattern[..equals].ToUpperInvariant();
            var expected = pattern[(equals + 1)..];
            string? actual = name switch
            {
                "PPTMCP_ID" => appId,
                "PPTMCP_COMPONENT" => component,
                _ => tags is not null && tags.TryGetValue(name, out var value) ? value : null,
            };
            return Exact(actual, expected);
        }

        private static bool GroupMatch(DeckObjectInfo item, IReadOnlyList<DeckObjectInfo> slideObjects, string group)
        {
            if (item.GroupShapeId is not { } parentId)
                return false;
            if (int.TryParse(group, NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                return parentId == id;
            var parent = slideObjects.FirstOrDefault(candidate => candidate.ShapeId == parentId);
            return parent is not null && (Exact(parent.Name, group) || Exact(parent.AppId, group));
        }

        private static bool HasWildcard(string pattern) => pattern.Contains('*', StringComparison.Ordinal) || pattern.Contains('?', StringComparison.Ordinal);

        private static Regex Wildcard(string pattern) =>
            new("^" + Regex.Escape(pattern).Replace("\\*", ".*", StringComparison.Ordinal).Replace("\\?", ".", StringComparison.Ordinal) + "$",
                RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    }

    /// <summary>Short placeholder name without the ppPlaceholder prefix, lower-cased (e.g. body, centertitle).</summary>
    public static string? PlaceholderShortName(string? placeholderType) =>
        placeholderType is { } type && type.StartsWith("ppPlaceholder", StringComparison.Ordinal)
            ? type["ppPlaceholder".Length..].ToLowerInvariant()
            : placeholderType;
}
