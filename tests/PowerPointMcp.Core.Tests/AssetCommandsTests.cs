using Sbroenne.PowerPointMcp.Core.Assets;
using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for placing, replacing, describing, and inspecting pictures against live PowerPoint COM.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Assets")]
public class AssetCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly AssetCommands _assets = new();
    private readonly DeckCommands _deck = new();

    public AssetCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Place_ContainKeepsAspectAndRecordsSourceAndAltText()
    {
        _fixture.CreateFreshPresentation();
        var image = CoreTestHelper.CreateUniqueTestImageFile();

        var placed = _assets.Place(_fixture.Batch, 1, image, left: 100, top: 100, width: 400, height: 200, altText: "Квадратный значок", name: "Icon");

        Assert.True(placed.Success, placed.ErrorMessage);
        var item = Assert.Single(_deck.InspectObjects(_fixture.Batch, slideIndex: 1, selector: "name:Icon", detail: "full").Objects!);
        Assert.Equal(200f, item.Width, 1f);
        Assert.Equal(200f, item.Height, 1f);
        Assert.Equal("Квадратный значок", item.AltText);
        Assert.Equal([1, 1], item.SourcePixelSize);
    }

    [Fact]
    public void Replace_KeepsFrameNameAndTags()
    {
        _fixture.CreateFreshPresentation();
        var first = CoreTestHelper.CreateUniqueTestImageFile();
        var second = CoreTestHelper.CreateUniqueTestImageFile();
        Assert.True(_assets.Place(_fixture.Batch, 1, first, left: 50, top: 60, width: 300, height: 150, fit: "cover", name: "Hero", altText: "Old").Success);
        Assert.True(_deck.AssignIds(_fixture.Batch, selector: "name:Hero", appId: "hero").Success);

        var replaced = _assets.Replace(_fixture.Batch, "appid:hero", second, altText: "New picture");

        Assert.True(replaced.Success, replaced.ErrorMessage);
        var item = Assert.Single(_deck.Find(_fixture.Batch, "appid:hero").Objects!);
        Assert.Equal(("Hero", 50f, 60f), (item.Name, item.Left, item.Top));
        Assert.Equal(300f, item.Width, 1f);
        Assert.Equal(150f, item.Height, 1f);
    }

    [Fact]
    public void SetAltTextAndInspect_ReportIssues()
    {
        _fixture.CreateFreshPresentation();
        var image = CoreTestHelper.CreateUniqueTestImageFile();
        Assert.True(_assets.Place(_fixture.Batch, 1, image, name: "Logo").Success);

        var before = _assets.Inspect(_fixture.Batch);
        var described = _assets.SetAltText(_fixture.Batch, "name:Logo", decorative: true, attribution: "© Example");
        var after = _assets.Inspect(_fixture.Batch, "name:Logo");

        Assert.Contains("missing-alt-text", Assert.Single(before.Pictures!).Issues!);
        Assert.True(described.Success, described.ErrorMessage);
        var picture = Assert.Single(after.Pictures!);
        Assert.True(picture.Decorative);
        Assert.Equal("© Example", picture.Attribution);
        Assert.Contains("low-resolution", picture.Issues!);
    }

    [Fact]
    public void Place_RefusesUrlsAndMissingFiles()
    {
        _fixture.CreateFreshPresentation();

        Assert.Contains("nothing is downloaded", _assets.Place(_fixture.Batch, 1, "https://example.org/a.png").ErrorMessage, StringComparison.Ordinal);
        Assert.False(_assets.Place(_fixture.Batch, 1, @"C:\no\such\image.png").Success);
    }
}
