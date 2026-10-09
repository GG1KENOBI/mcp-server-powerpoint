// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Text;

/// <summary>Literal/regex matching, replacement expansion, and range resolution for text edits (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "TextFrame")]
public sealed class TextSearchTests
{
    [Fact]
    public void Literal_IsCaseInsensitiveForCyrillicAndEscapesRegexCharacters()
    {
        var matcher = TextSearch.Compile("выручка (млн)", matchCase: false, wholeWords: false, useRegex: false);

        var hits = TextSearch.Find("Выручка (млн): 12; ВЫРУЧКА (МЛН): 14", matcher, "доход", useRegex: false);

        Assert.Equal([0, 19], hits.Select(hit => hit.Start));
        Assert.All(hits, hit => Assert.Equal("доход", hit.Replacement));
        Assert.Equal("ВЫРУЧКА (МЛН)", hits[1].Matched);
    }

    [Fact]
    public void MatchCase_DistinguishesCase()
    {
        var matcher = TextSearch.Compile("Q3", matchCase: true, wholeWords: false, useRegex: false);

        Assert.Single(TextSearch.Find("Q3 and q3", matcher, null, useRegex: false));
    }

    [Fact]
    public void WholeWords_SkipsMatchesInsideWords()
    {
        var english = TextSearch.Compile("cat", matchCase: false, wholeWords: true, useRegex: false);
        var russian = TextSearch.Compile("план", matchCase: false, wholeWords: true, useRegex: false);

        Assert.Equal([5], TextSearch.Find("cats cat concat", english, null, false).Select(hit => hit.Start));
        Assert.Equal([0], TextSearch.Find("план планирование", russian, null, false).Select(hit => hit.Start));
    }

    [Fact]
    public void Regex_ExpandsGroupsInReplacement()
    {
        var matcher = TextSearch.Compile(@"(\d{4})-(\d{2})", matchCase: false, wholeWords: false, useRegex: true);

        var hit = Assert.Single(TextSearch.Find("Period 2024-03 closed", matcher, "$2/$1", useRegex: true));

        Assert.Equal("03/2024", hit.Replacement);
    }

    [Fact]
    public void LiteralReplacement_KeepsDollarSigns()
    {
        var matcher = TextSearch.Compile("price", matchCase: false, wholeWords: false, useRegex: false);

        var hit = Assert.Single(TextSearch.Find("price", matcher, "$1 cost", useRegex: false));

        Assert.Equal("$1 cost", hit.Replacement);
    }

    [Fact]
    public void NewlineInLiteralQuery_MatchesParagraphAndLineBreaks()
    {
        var matcher = TextSearch.Compile("end\nstart", matchCase: false, wholeWords: false, useRegex: false);

        var hits = TextSearch.Find("end\rstart, end\vstart", matcher, null, false);

        Assert.Equal(2, hits.Count);
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("([unclosed", true)]
    public void InvalidQueries_Throw(string query, bool useRegex)
    {
        Assert.Throws<ArgumentException>(() => TextSearch.Compile(query, false, false, useRegex));
    }

    [Fact]
    public void Context_MarksTheReplacementAndFlattensBreaks()
    {
        var text = "First line\rSecond line with target word";
        var matcher = TextSearch.Compile("target", false, false, false);
        var hit = Assert.Single(TextSearch.Find(text, matcher, "goal", false));

        var context = TextSearch.Context(text, hit, radius: 12, replacement: hit.Replacement);

        Assert.Equal("…d line with «goal» word", context);
    }

    [Fact]
    public void Resolve_DefaultsToWholeText()
    {
        Assert.Equal((1, 11), TextRanges.Resolve("Hello world", null, null, null, 1, false, null, null));
    }

    [Fact]
    public void Resolve_StartAndLengthAndAppendPosition()
    {
        Assert.Equal((7, 5), TextRanges.Resolve("Hello world", 7, 5, null, 1, false, null, null));
        Assert.Equal((7, 5), TextRanges.Resolve("Hello world", 7, null, null, 1, false, null, null));
        Assert.Equal((12, 0), TextRanges.Resolve("Hello world", 12, 0, null, 1, false, null, null));
        Assert.Throws<ArgumentException>(() => TextRanges.Resolve("Hello world", 8, 5, null, 1, false, null, null));
        Assert.Throws<ArgumentException>(() => TextRanges.Resolve("Hello world", 0, 1, null, 1, false, null, null));
    }

    [Fact]
    public void Resolve_MatchOccurrenceIncludingLast()
    {
        const string text = "Итог: 5. итог: 7. ИТОГ: 9.";
        Assert.Equal((10, 4), TextRanges.Resolve(text, null, null, "итог", 2, false, null, null));
        Assert.Equal((19, 4), TextRanges.Resolve(text, null, null, "итог", -1, false, null, null));
        Assert.Equal((10, 4), TextRanges.Resolve(text, null, null, "итог", 1, true, null, null));
        var missing = Assert.Throws<ArgumentException>(() => TextRanges.Resolve(text, null, null, "итог", 4, false, null, null));
        Assert.Contains("3 time(s)", missing.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Resolve_ParagraphsUseCrSeparators()
    {
        const string text = "One\rTwo two\r\rFour";
        Assert.Equal([(1, 3), (5, 7), (13, 0), (14, 4)], TextRanges.Paragraphs(text));
        Assert.Equal((5, 7), TextRanges.Resolve(text, null, null, null, 1, false, 2, null));
        Assert.Equal((5, 13), TextRanges.Resolve(text, null, null, null, 1, false, 2, 3));
        Assert.Throws<ArgumentException>(() => TextRanges.Resolve(text, null, null, null, 1, false, 4, 2));
    }

    [Fact]
    public void Resolve_RejectsAmbiguousSelections()
    {
        Assert.Throws<ArgumentException>(() => TextRanges.Resolve("abc", 1, 1, "a", 1, false, null, null));
        Assert.Throws<ArgumentException>(() => TextRanges.Resolve("abc", null, 1, null, 1, false, null, null));
        Assert.Throws<ArgumentException>(() => TextRanges.Resolve("abc", null, null, null, 1, false, null, 2));
    }

    [Fact]
    public void Newlines_BecomeParagraphSeparators()
    {
        Assert.Equal("a\rb\rc", TextRanges.ToPowerPointParagraphs("a\r\nb\nc"));
    }
}
