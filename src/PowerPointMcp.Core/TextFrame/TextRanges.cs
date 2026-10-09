using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>
/// Resolves which characters of a text frame an edit targets: an explicit 1-based start and
/// length, the Nth match of a literal string, or whole paragraphs. Positions are PowerPoint
/// character positions (1-based, paragraphs separated by CR). Pure.
/// </summary>
public static class TextRanges
{
    /// <summary>
    /// Returns the 1-based start and length the arguments select. With no selection the whole text
    /// is selected. Throws <see cref="ArgumentException"/> with an actionable message otherwise.
    /// </summary>
    /// <param name="text">The frame's text as PowerPoint returns it.</param>
    /// <param name="start">1-based first character (text length + 1 selects the end, for insertion).</param>
    /// <param name="length">Number of characters; null means to the end of the text.</param>
    /// <param name="match">Literal text whose occurrence is selected.</param>
    /// <param name="occurrence">Which occurrence of <paramref name="match"/> (1-based; -1 for the last).</param>
    /// <param name="matchCase">Match case exactly when searching for <paramref name="match"/>.</param>
    /// <param name="paragraph">1-based first paragraph.</param>
    /// <param name="paragraphCount">Paragraphs from <paramref name="paragraph"/> (default 1).</param>
    public static (int Start, int Length) Resolve(
        string text,
        int? start,
        int? length,
        string? match,
        int occurrence,
        bool matchCase,
        int? paragraph,
        int? paragraphCount)
    {
        ArgumentNullException.ThrowIfNull(text);
        int ways = (start is not null ? 1 : 0) + (match is not null ? 1 : 0) + (paragraph is not null ? 1 : 0);
        if (ways > 1)
            throw new ArgumentException("Choose one of start (with length), match (with occurrence), or paragraph (with paragraph_count).");
        if (length is < 0)
            throw new ArgumentException("length must be 0 or greater.");
        if (paragraphCount is not null && paragraph is null)
            throw new ArgumentException("paragraph_count needs paragraph.");
        if (length is not null && start is null)
            throw new ArgumentException("length needs start.");

        if (start is { } first)
        {
            if (first < 1 || first > text.Length + 1)
                throw new ArgumentException($"start {first} is outside the text (1-{text.Length + 1}).");
            int count = length ?? text.Length - first + 1;
            if (first - 1 + count > text.Length)
                throw new ArgumentException($"start {first} with length {count} runs past the end of the text ({text.Length} characters).");
            return (first, count);
        }

        if (match is not null)
        {
            if (occurrence == 0 || occurrence < -1)
                throw new ArgumentException("occurrence must be 1 or greater, or -1 for the last match.");
            var hits = TextSearch.Find(text, TextSearch.Compile(match, matchCase, wholeWords: false, useRegex: false), null, useRegex: false);
            if (hits.Count == 0)
                throw new ArgumentException($"'{match}' does not occur in the text.");
            if (occurrence > hits.Count)
                throw new ArgumentException($"'{match}' occurs {hits.Count} time(s); occurrence {occurrence} does not exist.");
            var hit = occurrence == -1 ? hits[^1] : hits[occurrence - 1];
            return (hit.Start + 1, hit.Length);
        }

        if (paragraph is { } index)
        {
            var bounds = Paragraphs(text);
            int count = paragraphCount ?? 1;
            if (count < 1)
                throw new ArgumentException("paragraph_count must be 1 or greater.");
            if (index < 1 || index + count - 1 > bounds.Count)
                throw new ArgumentException($"The text has {bounds.Count} paragraph(s); paragraphs {index}-{index + count - 1} do not all exist.");
            var firstParagraph = bounds[index - 1];
            var lastParagraph = bounds[index + count - 2];
            return (firstParagraph.Start, lastParagraph.Start + lastParagraph.Length - firstParagraph.Start);
        }

        return (1, text.Length);
    }

    /// <summary>1-based start and length (without the trailing CR) of each paragraph.</summary>
    public static IReadOnlyList<(int Start, int Length)> Paragraphs(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var result = new List<(int Start, int Length)>();
        int from = 0;
        for (int index = 0; index <= text.Length; index++)
        {
            if (index == text.Length || text[index] == '\r')
            {
                result.Add((from + 1, index - from));
                from = index + 1;
            }
        }
        return result;
    }

    /// <summary>Converts any newline style to PowerPoint's paragraph separator (CR).</summary>
    public static string ToPowerPointParagraphs(string text) =>
        (text ?? "").Replace("\r\n", "\r", StringComparison.Ordinal).Replace('\n', '\r');
}
