// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Composition;

/// <summary>Layout box arithmetic, text estimation, and picture fitting (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Composition")]
public sealed class LayoutEngineTests
{
    private static readonly ResolvedProfile Profile = DesignProfiles.ResolveBuiltIn("default");

    [Theory]
    [InlineData(960f, 540f)]
    [InlineData(720f, 540f)]
    [InlineData(1280f, 720f)]
    public void Frame_RegionsAreInsideTheSafeAreaAndOrdered(float width, float height)
    {
        var frame = LayoutEngine.Frame(Profile, width, height, hasTitle: true, hasSource: true);
        Assert.Equal(new Box(48, 32, width - 48, height - 28), frame.Safe);
        Assert.True(frame.Safe.Contains(frame.Title, 0) && frame.Safe.Contains(frame.Body, 0) && frame.Safe.Contains(frame.Footer, 0));
        Assert.True(frame.Title.Bottom < frame.Body.Top);
        Assert.True(frame.Body.Bottom < frame.Source.Top);
        Assert.True(frame.Source.Bottom <= frame.Footer.Top);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(7)]
    public void Columns_TileTheAreaWithGutters(int count)
    {
        var area = new Box(48, 100, 912, 400);
        var columns = LayoutEngine.Columns(area, count, 16);
        Assert.Equal(count, columns.Count);
        Assert.Equal(area.Left, columns[0].Left);
        Assert.Equal(area.Right, columns[^1].Right);
        for (int i = 1; i < count; i++)
            Assert.Equal(16f, columns[i].Left - columns[i - 1].Right, 1);
        Assert.True(columns.Max(c => c.Width) - columns.Min(c => c.Width) <= 0.2f);
    }

    [Fact]
    public void Columns_HonourWeights()
    {
        var columns = LayoutEngine.Columns(new Box(0, 0, 1016, 100), 2, 16, [3f, 1f]);
        Assert.Equal(750f, columns[0].Width, 1);
        Assert.Equal(250f, columns[1].Width, 1);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(2, 2)]
    [InlineData(3, 3)]
    [InlineData(4, 4)]
    [InlineData(5, 3)]
    [InlineData(7, 4)]
    [InlineData(8, 4)]
    public void AutoColumns_BalancesRows(int items, int expected)
    {
        Assert.Equal(expected, LayoutEngine.AutoColumns(items, 864, 16, 180));
    }

    [Fact]
    public void Grid_LastRowKeepsCellWidth()
    {
        var cells = LayoutEngine.Grid(new Box(0, 0, 632, 416), 5, 3, 16, 16);
        Assert.Equal(5, cells.Count);
        Assert.Equal(cells[0].Width, cells[4].Width, 1);
        Assert.Equal(cells[0].Left, cells[3].Left);
        Assert.True(cells[3].Top > cells[0].Bottom);
    }

    [Fact]
    public void Stack_IsDeterministicAndRounded()
    {
        var first = LayoutEngine.Stack(new Box(10.04f, 20.06f, 300, 600), [33.333f, 12.25f, 40f], 8.05f);
        var second = LayoutEngine.Stack(new Box(10.04f, 20.06f, 300, 600), [33.333f, 12.25f, 40f], 8.05f);
        Assert.Equal(first, second);
        Assert.All(first, box => Assert.Equal(box.Top, MathF.Round(box.Top * 10) / 10));
        Assert.Equal(LayoutEngine.StackHeight([33.333f, 12.25f, 40f], 8.05f), first[^1].Bottom - first[0].Top, 0);
    }

    [Fact]
    public void Inset_NeverInvertsBoxes()
    {
        var box = LayoutEngine.Inset(new Box(0, 0, 10, 10), 20);
        Assert.True(box.Width >= 0 && box.Height >= 0);
    }

    [Fact]
    public void TextEstimator_GrowsWithTextAndShrinksWithWidth()
    {
        var shortText = TextEstimator.Height("Revenue grew", 16, 300);
        var longText = TextEstimator.Height(string.Join(' ', Enumerable.Repeat("Revenue grew in every region", 10)), 16, 300);
        var wide = TextEstimator.Height(string.Join(' ', Enumerable.Repeat("Revenue grew in every region", 10)), 16, 800);
        Assert.Equal(16 * TextEstimator.DefaultLineHeight, shortText, 1);
        Assert.True(longText > shortText * 3);
        Assert.True(wide < longText);
    }

    [Fact]
    public void TextEstimator_CyrillicIsWiderThanLatin()
    {
        Assert.True(TextEstimator.LineWidth("выручка выросла", 16) > TextEstimator.LineWidth("revenue climbed", 16));
    }

    [Fact]
    public void TextEstimator_LongWordsBreak()
    {
        Assert.True(TextEstimator.Lines(new string('W', 200), 20, 200) > 5);
    }

    [Fact]
    public void TextEstimator_FitSizeRespectsMinimum()
    {
        var text = string.Join(' ', Enumerable.Repeat("word", 400));
        Assert.Null(TextEstimator.FitSize(text, 20, 12, 300, 100));
        Assert.Equal(20f, TextEstimator.FitSize("short", 20, 12, 300, 100));
    }

    [Theory]
    [InlineData(16.0 / 9.0)]
    [InlineData(1.0)]
    [InlineData(0.5)]
    public void PictureFit_ContainKeepsAspectInsideFrame(double aspect)
    {
        var frame = new Box(100, 100, 500, 400);
        var placement = PictureFit.Compute(frame, aspect, "contain");
        Assert.True(frame.Contains(placement.Frame, 0.2f));
        Assert.Equal(aspect, placement.Frame.Width / placement.Frame.Height, 2);
        Assert.False(placement.Cropped);
    }

    [Theory]
    [InlineData(16.0 / 9.0, 0.5f, 0.5f)]
    [InlineData(0.5, 0.5f, 0.1f)]
    [InlineData(3.0, 0f, 0.5f)]
    public void PictureFit_CoverFillsFrameWithoutDistortion(double aspect, float focalX, float focalY)
    {
        var frame = new Box(100, 100, 500, 400);
        var placement = PictureFit.Compute(frame, aspect, "cover", focalX, focalY);
        Assert.Equal(frame, placement.Frame);
        Assert.Equal(aspect, placement.PictureWidth / placement.PictureHeight, 2);
        Assert.True(placement.PictureWidth >= frame.Width - 0.1f && placement.PictureHeight >= frame.Height - 0.1f);
        Assert.True(Math.Abs(placement.OffsetX) <= ((placement.PictureWidth - frame.Width) / 2) + 0.1f);
        Assert.True(Math.Abs(placement.OffsetY) <= ((placement.PictureHeight - frame.Height) / 2) + 0.1f);
    }

    [Fact]
    public void PictureFit_FocalPointShiftsTowardTheSubject()
    {
        var frame = new Box(0, 0, 400, 400);
        var top = PictureFit.Compute(frame, 0.5, "cover", 0.5f, 0.1f);
        var bottom = PictureFit.Compute(frame, 0.5, "cover", 0.5f, 0.9f);
        Assert.True(top.OffsetY > 0);
        Assert.True(bottom.OffsetY < 0);
    }

    [Fact]
    public void TreeLayout_CentersParentsAndFitsArea()
    {
        var tree = new SpecNode("CEO", null,
        [
            new SpecNode("CFO", null, [new SpecNode("Controller", null, []), new SpecNode("Treasury", null, [])]),
            new SpecNode("CTO", null, [new SpecNode("Platform", null, [])]),
        ]);
        var area = new Box(48, 120, 912, 480);
        var nodes = TreeLayout.Layout(tree, area, 44, 64, 190, 8)!;
        Assert.Equal(6, nodes.Count);
        Assert.All(nodes, node => Assert.True(area.Contains(node.Box, 0.5f)));
        var ceo = nodes.Single(node => node.Label == "CEO");
        var cfo = nodes.Single(node => node.Label == "CFO");
        var children = nodes.Where(node => node.ParentId == cfo.Id).ToList();
        Assert.Equal((children[0].Box.Left + children[^1].Box.Right) / 2, (cfo.Box.Left + cfo.Box.Right) / 2, 0);
        Assert.Equal(0, ceo.Depth);
        Assert.Null(TreeLayout.Layout(tree, new Box(0, 0, 100, 100), 44, 64, 190, 8));
    }
}
