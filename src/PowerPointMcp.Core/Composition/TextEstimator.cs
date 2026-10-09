using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>
/// Estimates rendered text size from average glyph widths for proportional sans-serif fonts
/// (Latin and Cyrillic). Used for planning and splitting only: rendering always measures with
/// PowerPoint, and results say which numbers are estimated and which are measured.
/// </summary>
public static class TextEstimator
{
    /// <summary>Default line height as a multiple of the font size.</summary>
    public const float DefaultLineHeight = 1.2f;

    /// <summary>Approximate advance width of a character in em.</summary>
    public static float CharWidth(char character)
    {
        if (char.IsWhiteSpace(character))
            return 0.28f;
        if (char.IsDigit(character))
            return 0.56f;
        if (character is >= '一' and <= '鿿' or >= '぀' and <= 'ヿ' or >= '가' and <= '힯')
            return 1f;
        if (character is 'i' or 'l' or 'j' or 'I' or '.' or ',' or ':' or ';' or '!' or '|' or '\'' or 'і' or 'ї')
            return 0.26f;
        if (character is 'm' or 'w' or 'M' or 'W' or 'ш' or 'щ' or 'ж' or 'ю' or 'Ш' or 'Щ' or 'Ж' or 'Ю' or 'ф' or 'Ф' or '%' or '@')
            return 0.86f;
        var category = CharUnicodeInfo.GetUnicodeCategory(character);
        bool cyrillic = character is >= 'Ѐ' and <= 'ӿ';
        return category switch
        {
            UnicodeCategory.UppercaseLetter => cyrillic ? 0.66f : 0.63f,
            UnicodeCategory.LowercaseLetter => cyrillic ? 0.54f : 0.51f,
            UnicodeCategory.DashPunctuation or UnicodeCategory.OpenPunctuation or UnicodeCategory.ClosePunctuation => 0.34f,
            _ => 0.55f,
        };
    }

    /// <summary>Estimated width of a single line in points.</summary>
    public static float LineWidth(string text, float fontSize, bool bold = false)
    {
        ArgumentNullException.ThrowIfNull(text);
        var em = 0f;
        foreach (var character in text)
            em += CharWidth(character);
        return em * fontSize * (bold ? 1.06f : 1f);
    }

    /// <summary>Number of wrapped lines for one paragraph in the given width (word wrap, long words broken).</summary>
    public static int Lines(string paragraph, float fontSize, float width, bool bold = false)
    {
        ArgumentNullException.ThrowIfNull(paragraph);
        if (width <= fontSize)
            return Math.Max(1, paragraph.Length);
        if (paragraph.Length == 0)
            return 1;

        int lines = 1;
        float current = 0;
        var space = LineWidth(" ", fontSize, bold);
        foreach (var word in paragraph.Split(' '))
        {
            var wordWidth = LineWidth(word, fontSize, bold);
            var needed = current == 0 ? wordWidth : current + space + wordWidth;
            if (needed <= width)
            {
                current = needed;
                continue;
            }
            if (current > 0)
            {
                lines++;
                current = 0;
            }
            // A word longer than the line breaks across lines.
            while (wordWidth > width)
            {
                lines++;
                wordWidth -= width;
            }
            current = wordWidth;
        }
        return lines;
    }

    /// <summary>
    /// Estimated height in points of text (paragraphs separated by newlines) in a box of the given
    /// inner width, including paragraph spacing but not frame margins.
    /// </summary>
    public static float Height(string text, float fontSize, float width, float lineHeight = DefaultLineHeight, float spaceAfter = 0f, bool bold = false, float indent = 0f)
    {
        ArgumentNullException.ThrowIfNull(text);
        var paragraphs = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n', '\r', '\v');
        var lines = paragraphs.Sum(paragraph => Lines(paragraph, fontSize, width - indent, bold));
        return (lines * fontSize * lineHeight) + (spaceAfter * Math.Max(0, paragraphs.Length - 1));
    }

    /// <summary>
    /// The largest size from <paramref name="start"/> down to <paramref name="minimum"/> (0.5 pt steps)
    /// whose estimated height fits <paramref name="maxHeight"/>; null when even the minimum does not fit.
    /// </summary>
    public static float? FitSize(string text, float start, float minimum, float width, float maxHeight, float lineHeight = DefaultLineHeight, float spaceAfter = 0f, bool bold = false)
    {
        for (var size = start; size >= minimum - 0.001f; size -= 0.5f)
        {
            if (Height(text, size, width, lineHeight, spaceAfter * size / start, bold) <= maxHeight)
                return size;
        }
        return null;
    }
}
