using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Template;

/// <summary>One paragraph to write into a placeholder.</summary>
public sealed record ContentParagraph(string Text, int Level);

/// <summary>Content for one placeholder role: text paragraphs or a local picture.</summary>
public sealed record ContentItem(string Key, string Role, int Ordinal, IReadOnlyList<ContentParagraph>? Paragraphs, string? PicturePath);

/// <summary>A placeholder on a slide, in reading order within its role.</summary>
public sealed record PlaceholderSlot(int ShapeId, string Name, string Role, int Ordinal, float Left, float Top, float Width, float Height, bool HasText);

/// <summary>Which placeholder each content item goes to, and what did not match.</summary>
public sealed record PlaceholderAssignment(IReadOnlyList<(ContentItem Item, PlaceholderSlot Slot)> Matches, IReadOnlyList<ContentItem> Unmatched, IReadOnlyList<PlaceholderSlot> Unused);

/// <summary>
/// Placeholder content by role, as JSON: {"title": "...", "subtitle": "...", "body": "text" or
/// ["point", {"text": "detail", "level": 2}], "body2": [...], "picture": "C:\\path\\photo.jpg",
/// "notes": "..."}. A number after a role selects the Nth placeholder of that role in reading
/// order (body = body1). "content" is an alias of body. Pure.
/// </summary>
public static partial class PlaceholderContent
{
    /// <summary>Roles a key may name.</summary>
    public static readonly IReadOnlyList<string> Roles = ["title", "subtitle", "body", "picture"];

    /// <summary>Parses the content JSON; throws <see cref="ArgumentException"/> with what is wrong.</summary>
    public static (IReadOnlyList<ContentItem> Items, string? Notes) Parse(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("content is required, e.g. {\"title\": \"Q3 results\", \"body\": [\"Revenue up 12%\", \"Margin stable\"]}.");
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip });
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"content is not valid JSON: {ex.Message}");
        }
        using (document)
        {
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                throw new ArgumentException("content must be a JSON object keyed by role (title, subtitle, body, body2, picture, notes).");
            var items = new List<ContentItem>();
            string? notes = null;
            foreach (var property in document.RootElement.EnumerateObject())
            {
                var key = property.Name.Trim().ToLowerInvariant();
                if (key == "notes")
                {
                    notes = property.Value.ValueKind == JsonValueKind.String ? property.Value.GetString() : throw new ArgumentException("notes must be a string.");
                    continue;
                }
                var match = KeyPattern().Match(key);
                if (!match.Success)
                    throw new ArgumentException($"Unknown content key '{property.Name}'. Use title, subtitle, body, body2, ..., picture, picture2, ..., or notes.");
                var role = match.Groups["role"].Value == "content" ? "body" : match.Groups["role"].Value;
                int ordinal = match.Groups["n"].Success ? int.Parse(match.Groups["n"].Value, CultureInfo.InvariantCulture) : 1;
                if (ordinal < 1)
                    throw new ArgumentException($"'{property.Name}': numbering starts at 1.");
                if (items.Any(item => item.Role == role && item.Ordinal == ordinal))
                    throw new ArgumentException($"'{property.Name}' is given twice.");
                if (role == "picture")
                {
                    if (property.Value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(property.Value.GetString()))
                        throw new ArgumentException($"'{property.Name}' must be the full path of a local image file.");
                    items.Add(new ContentItem(property.Name, role, ordinal, null, property.Value.GetString()));
                    continue;
                }
                items.Add(new ContentItem(property.Name, role, ordinal, Paragraphs(property.Name, property.Value), null));
            }
            return (items, notes);
        }
    }

    private static List<ContentParagraph> Paragraphs(string key, JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.String:
                return [.. Normalize(value.GetString()!).Split('\r').Select(line => new ContentParagraph(line, 1))];
            case JsonValueKind.Array:
                var result = new List<ContentParagraph>();
                foreach (var item in value.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.String)
                    {
                        result.Add(new ContentParagraph(Normalize(item.GetString()!).Replace('\r', '\v'), 1));
                    }
                    else if (item.ValueKind == JsonValueKind.Object && item.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
                    {
                        int level = item.TryGetProperty("level", out var levelValue) && levelValue.TryGetInt32(out var parsed) ? parsed : 1;
                        if (level is < 1 or > 9)
                            throw new ArgumentException($"'{key}': level must be between 1 and 9.");
                        result.Add(new ContentParagraph(Normalize(text.GetString()!).Replace('\r', '\v'), level));
                    }
                    else
                    {
                        throw new ArgumentException($"'{key}': list items must be strings or {{\"text\": \"...\", \"level\": 2}}.");
                    }
                }
                return result;
            default:
                throw new ArgumentException($"'{key}' must be a string or a list of paragraphs.");
        }
    }

    private static string Normalize(string text) => text.Replace("\r\n", "\r", StringComparison.Ordinal).Replace('\n', '\r');

    /// <summary>
    /// Assigns content to placeholders: each item first takes the placeholder of its role with the
    /// same ordinal (body2 = the second body placeholder in reading order). Then, for what is
    /// left, a picture falls back to the first free body (content) placeholder and a subtitle to
    /// the first free body placeholder.
    /// </summary>
    public static PlaceholderAssignment Assign(IReadOnlyList<ContentItem> items, IReadOnlyList<PlaceholderSlot> slots)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(slots);
        var free = slots.ToList();
        var matches = new List<(ContentItem, PlaceholderSlot)>();
        var pending = new List<ContentItem>();
        foreach (var item in items)
        {
            var slot = free.FirstOrDefault(candidate => candidate.Role == item.Role && candidate.Ordinal == item.Ordinal);
            if (slot is null)
            {
                pending.Add(item);
                continue;
            }
            free.Remove(slot);
            matches.Add((item, slot));
        }
        var unmatched = new List<ContentItem>();
        foreach (var item in pending.OrderBy(item => item.Role == "picture" ? 0 : 1))
        {
            var slot = item.Role is "picture" or "subtitle"
                ? free.Where(candidate => candidate.Role == "body").OrderBy(candidate => candidate.Ordinal).FirstOrDefault()
                : null;
            if (slot is null)
            {
                unmatched.Add(item);
                continue;
            }
            free.Remove(slot);
            matches.Add((item, slot));
        }
        return new PlaceholderAssignment(matches, unmatched, free);
    }

    [GeneratedRegex("^(?<role>title|subtitle|body|content|picture)(?<n>[0-9]{1,2})?$", RegexOptions.CultureInvariant)]
    private static partial Regex KeyPattern();
}

/// <summary>Parses slide lists like "1-3,5" and resolves layout names. Pure.</summary>
public static class TemplateText
{
    /// <summary>Parses "1-3,5,8-9" into ordered, distinct 1-based slide numbers within 1..count.</summary>
    public static IReadOnlyList<int> ParseSlides(string? text, int count)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Trim() is "all" or "*")
            return Enumerable.Range(1, count).ToList();
        var result = new List<int>();
        foreach (var part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var bounds = part.Split('-', StringSplitOptions.TrimEntries);
            if (bounds.Length is < 1 or > 2
                || !int.TryParse(bounds[0], NumberStyles.None, CultureInfo.InvariantCulture, out var first)
                || (bounds.Length == 2 && !int.TryParse(bounds[1], NumberStyles.None, CultureInfo.InvariantCulture, out _)))
                throw new ArgumentException($"'{part}' is not a slide number or range like 2-4.");
            int last = bounds.Length == 2 ? int.Parse(bounds[1], CultureInfo.InvariantCulture) : first;
            if (first < 1 || last < first || last > count)
                throw new ArgumentException($"Slides {part} are outside the source presentation (1-{count}).");
            for (int slide = first; slide <= last; slide++)
            {
                if (!result.Contains(slide))
                    result.Add(slide);
            }
        }
        return result;
    }

    /// <summary>Groups slide numbers into consecutive runs (for InsertFromFile).</summary>
    public static IReadOnlyList<(int First, int Last)> Runs(IReadOnlyList<int> slides)
    {
        var runs = new List<(int, int)>();
        foreach (var slide in slides)
        {
            if (runs.Count > 0 && runs[^1].Item2 == slide - 1)
                runs[^1] = (runs[^1].Item1, slide);
            else
                runs.Add((slide, slide));
        }
        return runs;
    }

    /// <summary>
    /// Finds a layout by name: exact (ignoring case and extra spaces), then a unique partial
    /// match. Returns the index into <paramref name="names"/> or an error listing the choices.
    /// </summary>
    public static (int Index, string? Error) FindLayout(IReadOnlyList<string> names, string wanted)
    {
        ArgumentNullException.ThrowIfNull(names);
        static string Key(string value) => string.Join(' ', value.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant();
        var key = Key(wanted ?? "");
        if (key.Length == 0)
            return (-1, "layout is required. Use template list-layouts to see the names.");
        for (int index = 0; index < names.Count; index++)
        {
            if (Key(names[index]) == key)
                return (index, null);
        }
        var partial = Enumerable.Range(0, names.Count).Where(index => Key(names[index]).Contains(key, StringComparison.Ordinal)).ToList();
        if (partial.Count == 1)
            return (partial[0], null);
        var choices = partial.Count > 1 ? partial.Select(index => names[index]) : names;
        return (-1, $"{(partial.Count > 1 ? "Several layouts match" : "No layout is named")} '{wanted}'. Layouts: {string.Join(", ", choices.Distinct())}.");
    }
}
