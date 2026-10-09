using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Master;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for design profiles: apply with migration preview, typography
/// normalization, extraction, saved profiles, and component updates against live PowerPoint COM.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Design")]
public class DesignCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private const string CardsSpec = """{"kind":"cards","id":"prio","title":"Priorities","cards":[{"heading":"Growth","body":"Two markets"},{"heading":"People","body":"Hire 40"}]}""";

    private readonly SharedPresentationFixture _fixture;
    private readonly DesignCommands _design = new();
    private readonly ComposeCommands _compose = new();
    private readonly DeckCommands _deck = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TextFrameCommands _text = new();

    public DesignCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ApplyProfile_PreviewsThenRecolorsComposedSlideAndTheme()
    {
        _fixture.CreateFreshPresentation();
        var composed = _compose.Create(_fixture.Batch, spec: CardsSpec, profile: "default");
        Assert.True(composed.Success, composed.ErrorMessage);
        var before = _deck.Fingerprint(_fixture.Batch).Revision;

        var preview = _design.ApplyProfile(_fixture.Batch, "corporate-blue");

        Assert.True(preview.Success, preview.ErrorMessage);
        Assert.False(preview.Applied);
        Assert.True(preview.ChangesByProperty!.ContainsKey("fill"));
        Assert.True(preview.ChangesByProperty.ContainsKey("theme-color"));
        Assert.Equal(before, _deck.Fingerprint(_fixture.Batch).Revision);

        var applied = _design.ApplyProfile(_fixture.Batch, "corporate-blue", dryRun: false);

        Assert.True(applied.Success, applied.ErrorMessage);
        Assert.True(applied.Applied);
        var fills = _deck.InspectObjects(_fixture.Batch, slideIndex: composed.Slides![0].SlideIndex, detail: "full").Objects!.Select(item => item.FillColor).ToList();
        Assert.Contains("#EEF3FA", fills);
        Assert.DoesNotContain("#F2F4F7", fills);
        var theme = new MasterCommands().GetThemeColors(_fixture.Batch, 1);
        Assert.Equal("#1F3864", theme.ThemeColors!["Accent1"]);
    }

    [Fact]
    public void NormalizeTypography_UnifiesFontsAndSnapsSizes()
    {
        _fixture.CreateFreshPresentation();
        var box = _shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 400, 80, "Отчёт за квартал");
        Assert.True(box.Success, box.ErrorMessage);
        Assert.True(_text.FormatRange(_fixture.Batch, 1, box.ShapeIndex!.Value, fontName: "Arial", fontSize: 15).Success);

        var result = _design.NormalizeTypography(_fixture.Batch, profile: "default", dryRun: false);

        Assert.True(result.Success, result.ErrorMessage);
        var run = _text.GetParagraphs(_fixture.Batch, 1, box.ShapeIndex.Value).Paragraphs![0].Runs![0];
        Assert.Equal(("Segoe UI", 16f), (run.FontName, run.FontSize!.Value));
    }

    [Fact]
    public void ExtractProfile_ReturnsAValidDraft()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_compose.Create(_fixture.Batch, spec: CardsSpec, profile: "default").Success);

        var result = _design.ExtractProfile(_fixture.Batch, "from-deck");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.Errors);
        var parsed = DesignProfiles.Parse(result.ProfileJson!);
        Assert.Equal("from-deck", parsed.Name);
        Assert.True(DesignProfiles.Resolve(parsed).IsValid);
    }

    [Fact]
    public void SavedProfile_UpdatesComponentInstancesToItsVersion()
    {
        var folder = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"profiles-{Guid.NewGuid():N}");
        var previous = Environment.GetEnvironmentVariable("PPTMCP_PROFILES_DIR");
        Environment.SetEnvironmentVariable("PPTMCP_PROFILES_DIR", folder);
        try
        {
            _fixture.CreateFreshPresentation();
            Assert.True(_compose.Create(_fixture.Batch, spec: CardsSpec, profile: "default").Success);
            var saved = _design.SaveProfile(_fixture.Batch, """
                {"schema":"pptmcp.design-profile/1","name":"cards-v2","version":"2.0.0","base":"default",
                 "colors":{"surface":"#FFF8E1"},
                 "components":{"card":{"version":"2","fill":"surface"}}}
                """);
            Assert.True(saved.Success, saved.ErrorMessage);

            var listed = _design.ListComponents(_fixture.Batch, "cards-v2");
            var card = Assert.Single(listed.Components!, item => item.Name == "card");
            Assert.True(card.Outdated > 0);

            var updated = _design.UpdateComponents(_fixture.Batch, "cards-v2", fromProfile: "default", dryRun: false);

            Assert.True(updated.Success, updated.ErrorMessage);
            Assert.Equal(0, _design.ListComponents(_fixture.Batch, "cards-v2").Components!.Single(item => item.Name == "card").Outdated);
            var fills = _deck.InspectObjects(_fixture.Batch, selector: "component:card", detail: "full").Objects!.Select(item => item.FillColor);
            Assert.Contains("#FFF8E1", fills);
            Assert.True(_design.DeleteProfile(_fixture.Batch, "cards-v2").Success);
        }
        finally
        {
            Environment.SetEnvironmentVariable("PPTMCP_PROFILES_DIR", previous);
            if (Directory.Exists(folder))
                Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void ValidateProfile_ReportsErrorsWithoutSaving()
    {
        _fixture.CreateFreshPresentation();

        var result = _design.ValidateProfile(_fixture.Batch, """{"name":"bad","colors":{"primary":"blue-ish"}}""");

        Assert.False(result.Success);
        Assert.NotEmpty(result.Errors!);
    }
}
