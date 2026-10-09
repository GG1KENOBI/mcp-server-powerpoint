using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for deck inspection, selectors, and persistent ids against live PowerPoint COM.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Deck")]
public class DeckCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly DeckCommands _deck = new();
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();

    public DeckCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Summary_ReportsSizeSlidesAndTheme()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new PageSetupCommands().SetSize(_fixture.Batch, 960, 540).Success);
        Assert.True(_slides.AddBlank(_fixture.Batch).Success);

        var result = _deck.Summary(_fixture.Batch);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal((960f, 540f, "16:9"), (result.SlideWidth!.Value, result.SlideHeight!.Value, result.AspectRatio));
        Assert.Equal(result.SlideCount, result.Slides!.Count);
        Assert.All(result.Slides, slide => Assert.True(slide.SlideId > 0));
        Assert.All(result.Slides, slide => Assert.Matches("^[0-9a-f]{16}$", slide.Fingerprint));
        Assert.NotNull(result.ThemeColors);
        Assert.NotEmpty(result.Masters!);
        Assert.Matches("^[0-9a-f]{16}$", result.Revision);
    }

    [Fact]
    public void Summary_PaginatesSlides()
    {
        _fixture.CreateFreshPresentation();
        for (int i = 0; i < 4; i++) Assert.True(_slides.AddBlank(_fixture.Batch).Success);

        var result = _deck.Summary(_fixture.Batch, startSlide: 2, maxSlides: 2, includeTheme: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal([2, 3], result.Slides!.Select(slide => slide.SlideIndex));
        Assert.True(result.HasMore);
        Assert.Null(result.ThemeColors);
    }

    [Fact]
    public void InspectObjects_ReturnsIdsGeometryTextAndFonts()
    {
        _fixture.CreateFreshPresentation();
        var box = _shapes.AddTextBox(_fixture.Batch, 1, 50, 60, 300, 80, "Выручка выросла на 12%");
        Assert.True(box.Success, box.ErrorMessage);

        var result = _deck.InspectObjects(_fixture.Batch, slideIndex: 1, selector: "text:выручка");

        Assert.True(result.Success, result.ErrorMessage);
        var item = Assert.Single(result.Objects!);
        Assert.Equal("text-box", item.Kind);
        Assert.Equal((50f, 60f, 300f), (item.Left, item.Top, item.Width), new PointTolerance());
        Assert.Contains("12%", item.Text, StringComparison.Ordinal);
        Assert.NotNull(item.MinFontSize);
        Assert.True(item.ShapeId > 0);
    }

    [Fact]
    public void Find_SurvivesSlideReorderingBySlideIdAndAppId()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_slides.AddBlank(_fixture.Batch).Success);
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 2, 40, 40, 200, 40, "Target").Success);
        var assign = _deck.AssignIds(_fixture.Batch, selector: "text:Target", appId: "target-box");
        Assert.True(assign.Success, assign.ErrorMessage);
        var before = Assert.Single(_deck.Find(_fixture.Batch, "appid:target-box").Objects!);

        Assert.True(_slides.MoveTo(_fixture.Batch, 2, 1).Success);
        var after = Assert.Single(_deck.Find(_fixture.Batch, "appid:target-box").Objects!);

        Assert.Equal((before.SlideId, before.ShapeId), (after.SlideId, after.ShapeId));
        Assert.Equal((2, 1), (before.SlideIndex, after.SlideIndex));
        Assert.Single(_deck.Find(_fixture.Batch, $"slideid:{before.SlideId} id:{before.ShapeId}").Objects!);
    }

    [Fact]
    public void Find_RequireSingle_AmbiguousSelectorFailsWithCandidates()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 200, 40, "KPI one").Success);
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 100, 200, 40, "KPI two").Success);

        var result = _deck.Find(_fixture.Batch, "text:KPI", requireSingle: true);

        Assert.False(result.Success);
        Assert.Contains("ambiguous", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("KPI one", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("KPI two", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void AssignIds_RepairsDuplicatesLeftBySlideDuplication()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 200, 40, "Card").Success);
        Assert.True(_deck.AssignIds(_fixture.Batch, selector: "text:Card", appId: "card").Success);
        Assert.True(_slides.Duplicate(_fixture.Batch, 1).Success);

        var duplicates = _deck.Fingerprint(_fixture.Batch);
        Assert.Contains("object:card", duplicates.DuplicateAppIds!);

        var plan = _deck.AssignIds(_fixture.Batch, dryRun: true);
        Assert.True(plan.Success, plan.ErrorMessage);
        Assert.Contains(plan.Assignments!, item => item.Reason == "duplicate" && item.NewAppId == "card~2");
        Assert.Equal(2, _deck.Find(_fixture.Batch, "appid:card").Objects!.Count);

        var applied = _deck.AssignIds(_fixture.Batch);
        Assert.True(applied.Success, applied.ErrorMessage);
        Assert.Single(_deck.Find(_fixture.Batch, "appid:card").Objects!);
        Assert.Single(_deck.Find(_fixture.Batch, "appid:card~2").Objects!);
        Assert.Null(_deck.Fingerprint(_fixture.Batch).DuplicateAppIds);
    }

    [Fact]
    public void AssignIds_RefusesAnIdAlreadyUsedByAnotherObject()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 200, 40, "A").Success);
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 100, 200, 40, "B").Success);
        Assert.True(_deck.AssignIds(_fixture.Batch, selector: "text:A", appId: "taken").Success);

        var result = _deck.AssignIds(_fixture.Batch, selector: "text:B", appId: "taken");

        Assert.False(result.Success);
        Assert.Contains("taken", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void InspectObjects_ReportsGroupMembership()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddRectangle(_fixture.Batch, 1, 10, 10, 50, 50).Success);
        Assert.True(_shapes.AddRectangle(_fixture.Batch, 1, 80, 10, 50, 50).Success);
        int count = _shapes.GetCount(_fixture.Batch, 1).ShapeCount!.Value;
        var grouped = _shapes.Group(_fixture.Batch, 1, [count - 1, count]);
        Assert.True(grouped.Success, grouped.ErrorMessage);
        Assert.True(_shapes.SetName(_fixture.Batch, 1, grouped.ShapeIndex!.Value, "Pair").Success);

        var members = _deck.InspectObjects(_fixture.Batch, slideIndex: 1, selector: "group:Pair");

        Assert.True(members.Success, members.ErrorMessage);
        Assert.Equal(2, members.Objects!.Count);
        Assert.All(members.Objects, item => Assert.Null(item.ShapeIndex));
    }

    [Fact]
    public void Fingerprint_ChangesOnlyWhenContentChanges()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 200, 40, "Stable").Success);

        var first = _deck.Fingerprint(_fixture.Batch).Revision;
        var second = _deck.Fingerprint(_fixture.Batch).Revision;
        Assert.True(_shapes.SetPosition(_fixture.Batch, 1, 1, 60, 40).Success);
        var third = _deck.Fingerprint(_fixture.Batch).Revision;

        Assert.Equal(first, second);
        Assert.NotEqual(first, third);
    }

    private sealed class PointTolerance : IEqualityComparer<(float, float, float)>
    {
        public bool Equals((float, float, float) x, (float, float, float) y) =>
            Math.Abs(x.Item1 - y.Item1) < 0.5f && Math.Abs(x.Item2 - y.Item2) < 0.5f && Math.Abs(x.Item3 - y.Item3) < 0.5f;

        public int GetHashCode((float, float, float) obj) => 0;
    }
}
