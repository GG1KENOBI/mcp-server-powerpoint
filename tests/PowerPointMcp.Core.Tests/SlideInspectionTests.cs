using System.Globalization;
using System.Text.RegularExpressions;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Slide;
using Sbroenne.PowerPointMcp.Core.Table;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for slide inspect and check-layout against live PowerPoint COM: the
/// snapshot must match what the shape commands created, and the layout check must find (and
/// stop finding, once repaired) problems PowerPoint actually renders.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Slide")]
public partial class SlideInspectionTests : IClassFixture<SharedPresentationFixture>
{
    private const string LongText =
        "This sentence is deliberately long so that it needs several lines of text at this font size and cannot fit in a short box.";

    private readonly SharedPresentationFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TextFrameCommands _text = new();
    private readonly TableCommands _tables = new();

    public SlideInspectionTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Inspect_ReportsEveryShapeWithKindGeometryTextAndTableSize()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 30, 600, 60, "Quarterly review").Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, 1, 32).Success);
        Assert.True(_shapes.AddRectangle(_fixture.Batch, 1, 40, 150, 200, 100).Success);
        Assert.True(_tables.AddTable(_fixture.Batch, 1, 3, 2, 300, 150, 400, 150).Success);

        var result = _slides.Inspect(_fixture.Batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.SlideWidth > 0 && result.SlideHeight > 0);
        var shapes = Assert.IsAssignableFrom<IReadOnlyList<SlideShapeInfo>>(result.Shapes);
        Assert.Equal(3, shapes.Count);
        Assert.Equal([1, 2, 3], shapes.Select(shape => shape.ShapeIndex));

        var title = shapes[0];
        Assert.Equal("text-box", title.Kind);
        Assert.True(title.HasText);
        Assert.Equal("Quarterly review", title.Text?.TrimEnd('\r'));
        Assert.Equal(32, title.MinFontSize);
        Assert.Equal(32, title.MaxFontSize);
        Assert.Equal(40, title.Left, 1);
        Assert.Equal(30, title.Top, 1);
        Assert.Equal(600, title.Width, 1);
        Assert.NotNull(title.TextHeight);
        Assert.True(title.Visible);

        var rectangle = shapes[1];
        Assert.Equal("auto-shape", rectangle.Kind);
        Assert.False(rectangle.HasText);
        Assert.Equal(200, rectangle.Width, 1);
        Assert.Equal(100, rectangle.Height, 1);

        var table = shapes[2];
        Assert.Equal("table", table.Kind);
        Assert.Equal(3, table.TableRows);
        Assert.Equal(2, table.TableColumns);
        Assert.True(table.ZOrder > rectangle.ZOrder);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void Inspect_WithInvalidSlideIndex_Fails(int slideIndex)
    {
        _fixture.CreateFreshPresentation();

        var result = _slides.Inspect(_fixture.Batch, slideIndex);

        Assert.False(result.Success);
        Assert.Contains("out of range", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void CheckLayout_WellPlacedSlide_IsOk()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 50, 40, 600, 60, "Title").Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, 1, 32).Success);
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 50, 140, 600, 200, "First point\nSecond point").Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, 2, 20).Success);

        var result = _slides.CheckLayout(_fixture.Batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.LayoutOk, Describe(result));
        Assert.DoesNotContain(result.LayoutIssues!, issue => issue.Severity is "error" or "warning");
    }

    [Fact]
    public void CheckLayout_FindsOverflowOverlapOffSlideAndSmallText()
    {
        _fixture.CreateFreshPresentation();
        float slideWidth = _slides.Inspect(_fixture.Batch, 1).SlideWidth!.Value;

        // 1: overflowing text — fixed-size box (no autofit) far too short for its text.
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 300, 200, LongText).Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, 1, 28).Success);
        Assert.True(_text.SetAutoSize(_fixture.Batch, 1, 1, "ppAutoSizeNone").Success);
        Assert.True(_shapes.SetSize(_fixture.Batch, 1, 1, 300, 40).Success);
        // 2 and 3: two text boxes on top of each other.
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 400, 300, 300, 50, "Left label").Success);
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 420, 310, 300, 50, "Right label").Success);
        // 4: a rectangle entirely to the right of the slide.
        Assert.True(_shapes.AddRectangle(_fixture.Batch, 1, slideWidth + 50, 100, 100, 100).Success);
        // 5: 8 pt text.
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 420, 300, 30, "Fine print").Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, 5, 8).Success);

        var result = _slides.CheckLayout(_fixture.Batch, 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.LayoutOk);
        var issues = result.LayoutIssues!;
        AssertIssue(issues, "text-overflow", "error", [1]);
        AssertIssue(issues, "text-overlap", "error", [2, 3]);
        AssertIssue(issues, "off-slide", "error", [4]);
        AssertIssue(issues, "small-text", "warning", [5]);
        Assert.All(issues, issue => Assert.Equal(1, issue.SlideIndex));
    }

    [Fact]
    public void CheckLayout_OverflowSuggestion_RepairsTheShape()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_shapes.AddTextBox(_fixture.Batch, 1, 40, 40, 300, 200, LongText).Success);
        Assert.True(_text.SetFontSize(_fixture.Batch, 1, 1, 24).Success);
        Assert.True(_text.SetAutoSize(_fixture.Batch, 1, 1, "ppAutoSizeNone").Success);
        Assert.True(_shapes.SetSize(_fixture.Batch, 1, 1, 300, 30).Success);

        var overflow = Assert.Single(
            _slides.CheckLayout(_fixture.Batch, 1).LayoutIssues!,
            issue => issue.Code == "text-overflow");
        var match = SuggestedHeight().Match(overflow.Suggestion);
        Assert.True(match.Success, overflow.Suggestion);
        float height = float.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture);

        Assert.True(_shapes.SetSize(_fixture.Batch, 1, 1, 300, height).Success);
        var repaired = _slides.CheckLayout(_fixture.Batch, 1);

        Assert.True(repaired.Success, repaired.ErrorMessage);
        Assert.DoesNotContain(repaired.LayoutIssues!, issue => issue.Code == "text-overflow");
    }

    [Fact]
    public void CheckLayout_WithoutSlideIndex_ChecksEverySlide()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_slides.AddBlank(_fixture.Batch).Success);
        float slideHeight = _slides.Inspect(_fixture.Batch, 2).SlideHeight!.Value;
        Assert.True(_shapes.AddRectangle(_fixture.Batch, 2, 100, slideHeight + 20, 100, 50).Success);

        var all = _slides.CheckLayout(_fixture.Batch);
        var firstOnly = _slides.CheckLayout(_fixture.Batch, 1);

        Assert.True(all.Success, all.ErrorMessage);
        Assert.Equal(2, all.SlideCount);
        var issue = Assert.Single(all.LayoutIssues!);
        Assert.Equal(("off-slide", 2), (issue.Code, issue.SlideIndex));
        Assert.True(firstOnly.LayoutOk);
    }

    private static void AssertIssue(IReadOnlyList<SlideLayoutIssue> issues, string code, string severity, int[] shapes)
    {
        Assert.Contains(issues, issue =>
            issue.Code == code && issue.Severity == severity && issue.ShapeIndexes.SequenceEqual(shapes));
    }

    private static string Describe(SlideOperationResult result) =>
        string.Join(Environment.NewLine, result.LayoutIssues!.Select(issue => $"{issue.Code} {issue.Severity}: {issue.Message}"));

    [GeneratedRegex(@"height to at least (?<height>[0-9.]+) pt")]
    private static partial Regex SuggestedHeight();
}
