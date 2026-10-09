using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Image;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Review;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for validation and repair against live PowerPoint COM.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Review")]
public class ReviewCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private const string LongText =
        "Revenue grew in every region this quarter. EMEA led with new enterprise contracts, APAC doubled its partner channel, " +
        "and the Americas recovered after the second-quarter slowdown. Operating margin improved by three points.";

    private readonly SharedPresentationFixture _fixture;
    private readonly ReviewCommands _review = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TextFrameCommands _text = new();

    public ReviewCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    private void FreshWideDeck()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new PageSetupCommands().SetSize(_fixture.Batch, 960, 540).Success);
    }

    private int AddFixedTextBox(float left, float top, float width, float height, string text, float fontSize)
    {
        var box = _shapes.AddTextBox(_fixture.Batch, 1, left, top, width, height, text);
        Assert.True(box.Success, box.ErrorMessage);
        int index = box.ShapeIndex!.Value;
        Assert.True(_text.SetAutoSize(_fixture.Batch, 1, index, "ppAutoSizeNone").Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, index, fontSize).Success);
        Assert.True(_shapes.SetSize(_fixture.Batch, 1, index, width, height).Success);
        return index;
    }

    [Fact]
    public void Validate_ReportsMeasuredOverflowWithEvidence()
    {
        FreshWideDeck();
        AddFixedTextBox(48, 120, 300, 40, LongText, 20);

        var result = _review.Validate(_fixture.Batch, slideIndex: 1);

        Assert.True(result.Success, result.ErrorMessage);
        var overflow = Assert.Single(result.Findings!, finding => finding.Code == "text-overflow");
        Assert.Equal("deterministic", overflow.Certainty);
        Assert.True(float.Parse(overflow.Evidence!["textHeight"], System.Globalization.CultureInfo.InvariantCulture) > 40f);
        Assert.Matches("^[0-9a-f]{16}$", result.Revision);
    }

    [Fact]
    public void Repair_ExpandsOverflowingBoxIntoFreeSpace()
    {
        FreshWideDeck();
        AddFixedTextBox(48, 120, 300, 40, LongText, 20);

        var result = _review.Repair(_fixture.Batch, slideIndex: 1, codes: ["text-overflow"]);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains(result.Outcomes!, outcome => outcome.Status == "applied");
        Assert.DoesNotContain(result.Findings!, finding => finding.Code == "text-overflow");
        Assert.Contains(result.Changes!, change => change.Change == "object-modified" && change.Fields!.Contains("height"));
    }

    [Fact]
    public void Repair_ShrinkOnly_ScalesRunsUntilMeasuredFit_KeepingText()
    {
        FreshWideDeck();
        int index = AddFixedTextBox(48, 120, 420, 160, LongText, 24);

        var result = _review.Repair(_fixture.Batch, slideIndex: 1, codes: ["text-overflow"], fitPolicy: ["shrink"], minFontSize: 10);

        Assert.True(result.Success, result.ErrorMessage);
        var outcome = Assert.Single(result.Outcomes!);
        Assert.Equal("applied", outcome.Status);
        Assert.DoesNotContain(result.Findings!, finding => finding.Code == "text-overflow");
        var size = _text.GetFontSize(_fixture.Batch, 1, index);
        Assert.InRange(size.FontSize!.Value, 10f, 23.5f);
        Assert.Equal(LongText, _text.GetText(_fixture.Batch, 1, index).Text);
    }

    [Fact]
    public void Repair_MovesPartlyOffSlideShapeInside_AndRespectsExpectedRevision()
    {
        FreshWideDeck();
        var rect = _shapes.AddRectangle(_fixture.Batch, 1, 900, 100, 120, 60);
        Assert.True(rect.Success, rect.ErrorMessage);

        var stale = _review.Repair(_fixture.Batch, expectedRevision: "0000000000000000");
        Assert.False(stale.Success);
        Assert.Contains("changed since revision", stale.ErrorMessage, StringComparison.Ordinal);

        var revision = new DeckCommands().Fingerprint(_fixture.Batch).Revision;
        var result = _review.Repair(_fixture.Batch, expectedRevision: revision);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains(result.Outcomes!, outcome => outcome.Operation == "move" && outcome.Status == "applied");
        Assert.DoesNotContain(result.Findings!, finding => finding.Code == "off-slide");
        Assert.NotEqual(result.Revision, result.RevisionAfter);
    }

    [Fact]
    public void Repair_RestoresAspectOfStretchedLinkedPicture()
    {
        FreshWideDeck();
        var image = CoreTestHelper.CreateUniqueTestImageFile(); // 1x1 px
        var picture = new ImageCommands().AddPicture(_fixture.Batch, 1, image, 100, 100, 200, 100, linkToFile: true, saveWithDocument: true);
        Assert.True(picture.Success, picture.ErrorMessage);

        var plan = _review.PlanRepair(_fixture.Batch, codes: ["image-distortion"]);
        var action = Assert.Single(plan.Actions!);
        Assert.Equal([150f, 100f, 100f, 100f], action.After);

        var result = _review.Repair(_fixture.Batch, codes: ["image-distortion"]);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.DoesNotContain(result.Findings!, finding => finding.Code == "image-distortion");
    }

    [Fact]
    public void PlanRepair_ChangesNothing()
    {
        FreshWideDeck();
        Assert.True(_shapes.AddRectangle(_fixture.Batch, 1, 900, 100, 120, 60).Success);
        var before = new DeckCommands().Fingerprint(_fixture.Batch).Revision;

        var plan = _review.PlanRepair(_fixture.Batch);

        Assert.True(plan.Success, plan.ErrorMessage);
        Assert.NotEmpty(plan.Actions!);
        Assert.Equal(before, new DeckCommands().Fingerprint(_fixture.Batch).Revision);
    }

    [Fact]
    public void Validate_RejectsUnknownRuleCodes()
    {
        FreshWideDeck();
        var result = _review.Validate(_fixture.Batch, rules: ["overflow"]);
        Assert.False(result.Success);
        Assert.Contains("text-overflow", result.ErrorMessage, StringComparison.Ordinal);
    }
}
