using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>One match in a text: 0-based start, length, matched text, and the replacement it would get.</summary>
public sealed record TextHit(int Start, int Length, string Matched, string? Replacement);

/// <summary>
/// Finds literal or regular-expression matches in PowerPoint text. PowerPoint separates paragraphs
/// with CR and line breaks with VT; a newline in the query matches either. Matching is ordinal
/// (Unicode, so Cyrillic and Latin case-fold alike) and never culture-sensitive. Pure.
/// </summary>
public static class TextSearch
{
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(1);

    /// <summary>Builds the matcher; throws <see cref="ArgumentException"/> for an empty query or an invalid regular expression.</summary>
    public static Regex Compile(string findWhat, bool matchCase, bool wholeWords, bool useRegex)
    {
        if (string.IsNullOrEmpty(findWhat))
            throw new ArgumentException("find_what must not be empty.");
        var pattern = useRegex ? findWhat : Regex.Escape(findWhat).Replace("\\n", "[\\r\\v\\n]", StringComparison.Ordinal);
        if (wholeWords)
            pattern = $@"(?<![\w]){(useRegex ? "(?:" + pattern + ")" : pattern)}(?![\w])";
        var options = RegexOptions.CultureInvariant | (matchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
        try
        {
            return new Regex(pattern, options, Timeout);
        }
        catch (ArgumentException ex)
        {
            throw new ArgumentException($"find_what is not a valid regular expression: {ex.Message}");
        }
    }

    /// <summary>
    /// All non-overlapping matches. With <paramref name="replaceWhat"/>, each hit carries its
    /// replacement ($1, ${name} expand for regular expressions; literal otherwise).
    /// </summary>
    public static IReadOnlyList<TextHit> Find(string text, Regex matcher, string? replaceWhat, bool useRegex)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(matcher);
        var hits = new List<TextHit>();
        foreach (Match match in matcher.Matches(text))
        {
            if (match.Length == 0)
                continue;
            string? replacement = replaceWhat is null ? null : useRegex ? match.Result(replaceWhat) : replaceWhat;
            hits.Add(new TextHit(match.Index, match.Length, match.Value, replacement));
        }
        return hits;
    }

    /// <summary>The text around a hit, with line breaks shown as spaces and the hit marked with «».</summary>
    public static string Context(string text, TextHit hit, int radius = 30, string? replacement = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(hit);
        int from = Math.Max(0, hit.Start - radius);
        int to = Math.Min(text.Length, hit.Start + hit.Length + radius);
        static string Flat(string value) => value.Replace('\r', ' ').Replace('\v', ' ').Replace('\n', ' ');
        return (from > 0 ? "…" : "") + Flat(text[from..hit.Start]) + "«" + Flat(replacement ?? hit.Matched) + "»" + Flat(text[(hit.Start + hit.Length)..to]) + (to < text.Length ? "…" : "");
    }
}
