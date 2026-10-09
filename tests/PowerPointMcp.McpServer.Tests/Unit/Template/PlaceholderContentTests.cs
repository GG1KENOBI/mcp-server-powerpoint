// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Template;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Template;

/// <summary>Placeholder content parsing, role assignment, slide ranges, and layout lookup (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Template")]
public sealed class PlaceholderContentTests
{
    private static PlaceholderSlot Slot(int id, string role, int ordinal) => new(id, $"{role} {ordinal}", role, ordinal, 0, 0, 100, 100, false);

    [Fact]
    public void Parse_ReadsRolesListsLevelsAndNotes()
    {
        var (items, notes) = PlaceholderContent.Parse("""
            {"title": "Итоги", "content": ["Рост", {"text": "детали", "level": 2}], "body2": "a\nb", "picture": "C:\\img\\p.png", "notes": "Говорить медленно"}
            """);

        Assert.Equal("Говорить медленно", notes);
        var body = Assert.Single(items, item => item.Role == "body" && item.Ordinal == 1);
        Assert.Equal([("Рост", 1), ("детали", 2)], body.Paragraphs!.Select(p => (p.Text, p.Level)));
        Assert.Equal(2, items.Single(item => item.Key == "body2").Paragraphs!.Count);
        Assert.Equal("C:\\img\\p.png", items.Single(item => item.Role == "picture").PicturePath);
    }

    [Theory]
    [InlineData("""{"heading": "x"}""", "Unknown content key")]
    [InlineData("""{"body": 5}""", "string or a list")]
    [InlineData("""{"body": [{"text": "x", "level": 12}]}""", "level")]
    [InlineData("""[1]""", "JSON object")]
    [InlineData("""{"picture": ""}""", "local image")]
    public void Parse_RejectsBadContent(string json, string expected)
    {
        var error = Assert.Throws<ArgumentException>(() => PlaceholderContent.Parse(json));
        Assert.Contains(expected, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Assign_MatchesByRoleAndOrdinalWithSubtitleFallback()
    {
        var (items, _) = PlaceholderContent.Parse("""{"title": "T", "subtitle": "S", "body2": "B2", "body": "B1", "picture": "C:\\x.png"}""");
        var slots = new[] { Slot(2, "title", 1), Slot(3, "body", 1), Slot(4, "body", 2), Slot(5, "body", 3) };

        var assignment = PlaceholderContent.Assign(items, slots);

        var byKey = assignment.Matches.ToDictionary(match => match.Item.Key, match => match.Slot.ShapeId);
        Assert.Equal(2, byKey["title"]);
        Assert.Equal(3, byKey["body"]);
        Assert.Equal(4, byKey["body2"]);
        Assert.Equal(5, byKey["picture"]);
        Assert.Equal("subtitle", Assert.Single(assignment.Unmatched).Key);
        Assert.Empty(assignment.Unused);
    }

    [Fact]
    public void SlideRanges_ParseAndGroupIntoRuns()
    {
        var slides = TemplateText.ParseSlides("1-3, 5, 3, 8-9", 10);

        Assert.Equal([1, 2, 3, 5, 8, 9], slides);
        Assert.Equal([(1, 3), (5, 5), (8, 9)], TemplateText.Runs(slides));
        Assert.Equal(4, TemplateText.ParseSlides(null, 4).Count);
        Assert.Throws<ArgumentException>(() => TemplateText.ParseSlides("2-12", 10));
        Assert.Throws<ArgumentException>(() => TemplateText.ParseSlides("a", 10));
    }

    [Fact]
    public void FindLayout_ExactThenUniquePartialElseListsChoices()
    {
        string[] names = ["Title Slide", "Title and Content", "Two Content", "Comparison", "Title Only"];

        Assert.Equal(1, TemplateText.FindLayout(names, "title  AND content").Index);
        Assert.Equal(3, TemplateText.FindLayout(names, "compar").Index);
        var ambiguous = TemplateText.FindLayout(names, "title");
        Assert.Equal(-1, ambiguous.Index);
        Assert.Contains("Title Only", ambiguous.Error, StringComparison.Ordinal);
        Assert.Contains("No layout", TemplateText.FindLayout(names, "Agenda").Error, StringComparison.Ordinal);
    }
}
