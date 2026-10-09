using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.Presentation;
using Sbroenne.PowerPointMcp.Core.Template;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for template workflows: layouts, slides from layouts filled by role,
/// slide import with a report, and template preview against live PowerPoint COM.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Template")]
public class TemplateCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly TemplateCommands _template = new();

    public TemplateCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    /// <summary>A layout with exactly one title and at least one body placeholder (names are localized, so it is found by structure).</summary>
    private TemplateLayoutInfo ContentLayout()
    {
        var layouts = _template.ListLayouts(_fixture.Batch, masterIndex: 1);
        Assert.True(layouts.Success, layouts.ErrorMessage);
        return layouts.Layouts!.First(layout => layout.Placeholders!.Count(p => p.Role == "title") == 1 && layout.Placeholders!.Count(p => p.Role == "body") == 1);
    }

    [Fact]
    public void ListLayouts_ShowsPlaceholdersAndUsage()
    {
        _fixture.CreateFreshPresentation();

        var result = _template.ListLayouts(_fixture.Batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Layouts!.Count >= 5);
        Assert.Equal(1, result.Layouts.Sum(layout => layout.SlidesUsing));
        Assert.All(result.Layouts.SelectMany(layout => layout.Placeholders!), p => Assert.Equal(4, p.Bounds.Count));
    }

    [Fact]
    public void AddSlide_FillsTitleBodyLevelsAndNotes()
    {
        _fixture.CreateFreshPresentation();
        var layout = ContentLayout();

        var result = _template.AddSlide(_fixture.Batch, layout.Name, content: """
            {"title": "Итоги квартала", "body": ["Выручка +12%", {"text": "Россия +20%", "level": 2}], "notes": "Начать с выручки"}
            """);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.SlideIndex);
        Assert.Null(result.Unmatched);
        Assert.Null(result.EmptyPlaceholders);
        var notes = new NotesCommands().GetNotesText(_fixture.Batch, 2);
        Assert.Contains("Начать с выручки", notes.NotesText, StringComparison.Ordinal);
        Assert.Contains(result.Filled!, item => item.StartsWith("body", StringComparison.Ordinal));
    }

    [Fact]
    public void AddSlide_UnknownLayoutListsChoices()
    {
        _fixture.CreateFreshPresentation();

        var result = _template.AddSlide(_fixture.Batch, "No Such Layout 123");

        Assert.False(result.Success);
        Assert.Contains("Layouts:", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void FillPlaceholders_PlacesPictureIntoContentPlaceholderAndRemovesEmpty()
    {
        _fixture.CreateFreshPresentation();
        var layout = ContentLayout();
        var added = _template.AddSlide(_fixture.Batch, layout.Name);
        Assert.True(added.Success, added.ErrorMessage);
        var image = CoreTestHelper.CreateUniqueTestImageFile();

        var result = _template.FillPlaceholders(_fixture.Batch, added.SlideIndex!.Value, $$"""{"picture": {{System.Text.Json.JsonSerializer.Serialize(image)}}}""", removeEmpty: true);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains(result.Filled!, item => item.StartsWith("picture", StringComparison.Ordinal));
        Assert.Null(result.EmptyPlaceholders);
    }

    [Fact]
    public void ImportSlides_CopiesSelectedSlidesWithReport()
    {
        var sourcePath = _fixture.CreateFreshPresentation("pptmcp-import-source");
        var layout = ContentLayout();
        Assert.True(_template.AddSlide(_fixture.Batch, layout.Name, content: """{"title": "First", "body": "One"}""").Success);
        Assert.True(_template.AddSlide(_fixture.Batch, layout.Name, content: """{"title": "Second", "body": "Two"}""").Success);
        Assert.True(new PresentationCommands().Save(_fixture.Batch).Success);
        _fixture.CreateFreshPresentation("pptmcp-import-target");

        var result = _template.ImportSlides(_fixture.Batch, sourcePath, slides: "3", position: 1);

        Assert.True(result.Success, result.ErrorMessage);
        var imported = Assert.Single(result.ImportedSlides!);
        Assert.Equal((1, 3, "Second"), (imported.SlideIndex, imported.SourceSlide, imported.Title));
        Assert.Equal(layout.Name, imported.SourceLayout);
    }

    [Fact]
    public void PreviewTemplate_MapsLayoutsWithoutChangingTheDeck()
    {
        _fixture.CreateFreshPresentation();
        var themeBefore = new PresentationCommands().GetThemeName(_fixture.Batch).ThemeName;

        var result = _template.PreviewTemplate(_fixture.Batch, SharedTemplateAsset.TemplatePath);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Layouts!);
        Assert.NotEmpty(result.LayoutMapping!);
        Assert.All(result.LayoutMapping!, mapping => Assert.True(mapping.Status is "matched" or "similar" or "missing", mapping.Status));
        Assert.Equal(themeBefore, new PresentationCommands().GetThemeName(_fixture.Batch).ThemeName);
    }

    [Fact]
    public void ImportSlides_RefusesTheOpenPresentation()
    {
        var path = _fixture.CreateFreshPresentation();
        Assert.True(new PresentationCommands().Save(_fixture.Batch).Success);

        var result = _template.ImportSlides(_fixture.Batch, path);

        Assert.False(result.Success);
        Assert.Contains("open in this session", result.ErrorMessage, StringComparison.Ordinal);
    }
}
