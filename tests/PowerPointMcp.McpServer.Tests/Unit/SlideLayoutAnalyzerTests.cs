// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit;

/// <summary>
/// The layout rules behind slide check-layout are pure geometry over a snapshot, so every rule is
/// covered here without PowerPoint. Real-COM coverage of the snapshot lives in the Core tests.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "LayoutCheck")]
public sealed class SlideLayoutAnalyzerTests
{
    private const float Width = 960;
    private const float Height = 540;

    private static int s_nextIndex;

    private static SlideShapeInfo Box(
        float left, float top, float width, float height,
        string kind = "auto-shape", string? text = null, float? fontSize = null,
        float? textHeight = null, float? textWidth = null, float rotation = 0,
        string? placeholderType = null, bool visible = true, bool? hasPicture = null,
        int? tableRows = null, string? autoSize = null, int? index = null) =>
        new()
        {
            ShapeIndex = index ?? Interlocked.Increment(ref s_nextIndex),
            Name = kind,
            Kind = kind,
            PlaceholderType = placeholderType,
            Left = left,
            Top = top,
            Width = width,
            Height = height,
            Rotation = rotation,
            ZOrder = 1,
            Visible = visible,
            HasText = text is not null,
            Text = text,
            MinFontSize = fontSize,
            MaxFontSize = fontSize,
            AutoSize = autoSize,
            TextLeft = text is null ? null : left + 3.6f,
            TextTop = text is null ? null : top + 3.6f,
            TextWidth = text is null ? null : textWidth ?? width - 7.2f,
            TextHeight = text is null ? null : textHeight ?? height - 7.2f,
            HasPicture = hasPicture,
            TableRows = tableRows,
            TableColumns = tableRows is null ? null : 2,
        };

    private static IReadOnlyList<SlideLayoutIssue> Analyze(params SlideShapeInfo[] shapes) =>
        SlideLayoutAnalyzer.Analyze(3, Width, Height, shapes);

    [Fact]
    public void WellPlacedSlide_HasNoIssues()
    {
        var issues = Analyze(
            Box(50, 40, 860, 70, "text-box", "Title", 32, index: 1),
            Box(50, 140, 420, 300, "text-box", "Left column", 20, index: 2),
            Box(490, 140, 420, 300, "picture", index: 3));

        Assert.Empty(issues);
    }

    [Fact]
    public void TallText_IsVerticalOverflowWithTheHeightThatFits()
    {
        var issue = Assert.Single(Analyze(Box(50, 40, 400, 60, "text-box", "Long text", 24, textHeight: 130.4f, index: 7)));

        Assert.Equal(("text-overflow", "error", 3), (issue.Code, issue.Severity, issue.SlideIndex));
        Assert.Equal([7], issue.ShapeIndexes);
        Assert.Contains("130.4 pt tall but the shape is 60 pt tall", issue.Message, StringComparison.Ordinal);
        Assert.Contains("height to at least 138 pt", issue.Suggestion, StringComparison.Ordinal);
        Assert.Contains("largest now 24 pt", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void WideText_IsHorizontalOverflow()
    {
        var issue = Assert.Single(Analyze(Box(50, 40, 200, 60, "text-box", "Unbroken", 24, textWidth: 290)));

        Assert.Equal("text-overflow", issue.Code);
        Assert.Contains("horizontally", issue.Message, StringComparison.Ordinal);
        Assert.Contains("width to at least 298 pt", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void ShrinkOnOverflow_IsStillReportedAndExplained()
    {
        var issue = Assert.Single(Analyze(Box(50, 40, 400, 60, "placeholder", "Body", 18, textHeight: 100,
            placeholderType: "ppPlaceholderBody", autoSize: "shrink-text-on-overflow")));

        Assert.Contains("already shrinking", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RotatedText_IsNotCheckedForOverflow()
    {
        Assert.Empty(Analyze(Box(300, 200, 200, 40, "text-box", "Tilted", 18, textHeight: 90, rotation: 30)));
    }

    [Fact]
    public void ShapeEntirelyOffTheSlide_IsAnErrorWithAPositionThatFits()
    {
        var issue = Assert.Single(Analyze(Box(1000, 100, 100, 50, index: 4)));

        Assert.Equal(("off-slide", "error"), (issue.Code, issue.Severity));
        Assert.Contains("left=860 top=100", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void ShapePartlyOffTheSlide_IsAWarning()
    {
        var issue = Assert.Single(Analyze(Box(900, 500, 100, 100)));

        Assert.Equal(("off-slide", "warning"), (issue.Code, issue.Severity));
        Assert.Contains("left=860 top=440", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void ShapeLargerThanTheSlide_IsToldToResize()
    {
        var issue = Assert.Single(Analyze(Box(-10, 100, 1000, 100, "text-box", "Banner", 24)));

        Assert.Equal("off-slide", issue.Code);
        Assert.Contains("Resize it to fit within 960 x 540 pt", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void FullBleedBackground_IsNotReported()
    {
        Assert.Empty(Analyze(Box(-2, -2, 964, 544, "picture"), Box(100, 100, 300, 60, "text-box", "On top", 24)));
    }

    [Fact]
    public void RotatedShape_UsesItsRotatedBoundsForSlideEdges()
    {
        // 400 x 40 rotated 90 degrees around its center occupies 40 x 400 centered at (300, 520).
        var issue = Assert.Single(Analyze(Box(100, 500, 400, 40, rotation: 90)));

        Assert.Equal("off-slide", issue.Code);
        Assert.Contains("320..720", issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OverlappingText_IsAnErrorThatSuggestsMovingTheLowerShape()
    {
        var issue = Assert.Single(Analyze(
            Box(100, 100, 300, 80, "text-box", "Upper", 20, index: 1),
            Box(120, 150, 300, 80, "text-box", "Lower", 20, index: 2)));

        Assert.Equal(("text-overlap", "error"), (issue.Code, issue.Severity));
        Assert.Equal([1, 2], issue.ShapeIndexes);
        Assert.Contains("Move shape 2 down to top=192", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void OverlappingText_WithoutRoomBelow_SuggestsRestructuring()
    {
        var issue = Assert.Single(Analyze(
            Box(100, 300, 300, 200, "text-box", "Upper", 20),
            Box(120, 330, 300, 200, "text-box", "Lower", 20)));

        Assert.Contains("no room below", issue.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void TextOverATable_IsTextOverlap()
    {
        var issue = Assert.Single(Analyze(
            Box(100, 100, 400, 200, "table", tableRows: 4),
            Box(150, 150, 200, 40, "text-box", "Note", 18)));

        Assert.Equal("text-overlap", issue.Code);
    }

    [Fact]
    public void TextInsideACard_IsIntentionalLayering()
    {
        Assert.Empty(Analyze(
            Box(100, 100, 400, 200, "auto-shape"),
            Box(120, 120, 360, 60, "text-box", "Card title", 24)));
    }

    [Fact]
    public void PartialOverlapOfTextAndPicture_IsAWarning()
    {
        var issue = Assert.Single(Analyze(
            Box(100, 100, 300, 200, "picture", index: 1),
            Box(350, 150, 300, 60, "text-box", "Caption", 18, index: 2)));

        Assert.Equal(("partial-overlap", "warning"), (issue.Code, issue.Severity));
        Assert.Equal([1, 2], issue.ShapeIndexes);
    }

    [Fact]
    public void TouchingEdges_AreNotOverlap()
    {
        Assert.Empty(Analyze(
            Box(100, 100, 300, 60, "text-box", "Above", 20),
            Box(100, 161, 300, 60, "text-box", "Below", 20)));
    }

    [Fact]
    public void LinesAndConnectors_AreIgnoredForOverlap()
    {
        Assert.Empty(Analyze(
            Box(100, 100, 300, 60, "text-box", "Label", 20),
            Box(90, 90, 400, 100, "connector"),
            Box(90, 120, 400, 0, "line")));
    }

    [Theory]
    [InlineData(8f, true)]
    [InlineData(9.5f, true)]
    [InlineData(10f, false)]
    [InlineData(28f, false)]
    public void TextBelowTenPoints_IsSmallText(float size, bool reported)
    {
        var issues = Analyze(Box(100, 100, 300, 60, "text-box", "Note", size));

        Assert.Equal(reported, issues.Any(issue => issue.Code == "small-text" && issue.Severity == "warning"));
    }

    [Fact]
    public void EmptyContentPlaceholder_IsReported()
    {
        var issue = Assert.Single(Analyze(Box(50, 140, 860, 300, "placeholder", placeholderType: "ppPlaceholderBody", index: 2)));

        Assert.Equal(("empty-placeholder", "warning"), (issue.Code, issue.Severity));
        Assert.Contains("set-placeholder-text (shape_index 2)", issue.Suggestion, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("ppPlaceholderFooter", null)]
    [InlineData("ppPlaceholderSlideNumber", null)]
    [InlineData("ppPlaceholderDate", null)]
    [InlineData("ppPlaceholderPicture", true)]
    public void FooterOrFilledPlaceholders_AreNotEmpty(string placeholderType, bool? hasPicture)
    {
        Assert.Empty(Analyze(Box(50, 140, 400, 300, "placeholder", placeholderType: placeholderType, hasPicture: hasPicture)));
    }

    [Fact]
    public void EmptyPlaceholderUnderContent_IsReportedOnceNotAsOverlap()
    {
        var issues = Analyze(
            Box(50, 140, 860, 300, "placeholder", placeholderType: "ppPlaceholderBody"),
            Box(60, 150, 300, 200, "text-box", "Typed over it", 20));

        Assert.Equal("empty-placeholder", Assert.Single(issues).Code);
    }

    [Theory]
    [InlineData(102f, true)]
    [InlineData(104f, true)]
    [InlineData(100.3f, false)]
    [InlineData(110f, false)]
    public void NearlyAlignedLeftEdges_AreInfo(float secondLeft, bool reported)
    {
        var issues = Analyze(
            Box(100, 100, 300, 60, "text-box", "First", 20, index: 1),
            Box(secondLeft, 200, 300, 60, "text-box", "Second", 20, index: 2));

        var nearMiss = issues.Where(issue => issue.Code == "near-misaligned").ToArray();
        Assert.Equal(reported, nearMiss.Length == 1);
        if (reported)
        {
            Assert.Equal("info", nearMiss[0].Severity);
            Assert.Contains("Set left=100 on shape 2", nearMiss[0].Suggestion, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void HiddenShapes_AreIgnored()
    {
        Assert.Empty(Analyze(Box(2000, 2000, 10, 10, visible: false)));
    }

    [Fact]
    public void Issues_AreOrderedBySeverity()
    {
        var issues = Analyze(
            Box(100, 100, 300, 60, "text-box", "Note", 8, index: 1),
            Box(102, 300, 300, 60, "text-box", "Other", 20, index: 2),
            Box(1000, 100, 50, 50, index: 3));

        Assert.Equal(["error", "warning", "info"], issues.Select(issue => issue.Severity));
    }
}
