// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Catalog;
using Sbroenne.PowerPointMcp.Core.Preview;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Review.DeckValidatorTests;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Preview;

/// <summary>Wireframe sketch, image comparison, preview cache keys, and the command catalog (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Preview")]
public sealed class PreviewTests
{
    [Fact]
    public void Sketch_OutlinesObjectsWithLegend()
    {
        var sketch = LayoutSketch.Render(960, 540,
        [
            Obj(2, 48, 30, 864, 60, text: "Квартальные итоги", role: "title"),
            Obj(3, 48, 120, 400, 300, kind: "chart"),
            Obj(4, 500, 120, 400, 300, text: "Takeaway"),
            Obj(5, 0, 0, 960, 540, kind: "picture"),
        ], columns: 48);
        var lines = sketch.Split('\n');
        Assert.StartsWith("+" + new string('=', 48) + "+", lines[0], StringComparison.Ordinal);
        Assert.Contains("A = text-box role=title id=2", sketch, StringComparison.Ordinal);
        Assert.Contains("\"Квартальные итоги\"", sketch, StringComparison.Ordinal);
        Assert.Contains("D = picture id=5 at (0,0) 960x540 backdrop", sketch, StringComparison.Ordinal);
        Assert.Contains(lines, line => line.Contains("|B", StringComparison.Ordinal));
        Assert.Contains("one column = 20 pt", sketch, StringComparison.Ordinal);
    }

    [Fact]
    public void Sketch_IsDeterministicAndSkipsHidden()
    {
        var objects = new[] { Obj(2, 48, 30, 864, 60, text: "T"), Obj(3, 100, 200, 50, 50, kind: "auto-shape", visible: false) };
        Assert.Equal(LayoutSketch.Render(960, 540, objects), LayoutSketch.Render(960, 540, objects));
        Assert.Contains("B = auto-shape id=3 at (100,200) 50x50 hidden", LayoutSketch.Render(960, 540, objects), StringComparison.Ordinal);
    }

    [Fact]
    public void ImageDiff_FindsChangedBox()
    {
        const int width = 10, height = 8;
        var a = new byte[width * height * 4];
        var b = (byte[])a.Clone();
        foreach (var (x, y) in new[] { (2, 3), (5, 6) })
            b[(y * width * 4) + (x * 4) + 2] = 200;
        b[(1 * width * 4) + (1 * 4)] = 10; // below threshold
        var result = ImageDiff.Compare(width, height, width * 4, a, b);
        Assert.Equal(2, result.ChangedPixels);
        Assert.Equal((2, 3, 5, 6), (result.Left!.Value, result.Top!.Value, result.Right!.Value, result.Bottom!.Value));
        Assert.Equal(2.0 / 80, result.ChangedRatio, 6);
        Assert.Null(ImageDiff.Compare(width, height, width * 4, a, a).Left);
    }

    [Fact]
    public void PreviewCache_KeyChangesWithGenerationAndFingerprint()
    {
        var session = new object();
        var first = PreviewCache.PathFor(session, 256, 1280, "abc");
        Assert.Equal(first, PreviewCache.PathFor(session, 256, 1280, "abc"));
        Assert.NotEqual(first, PreviewCache.PathFor(session, 256, 1280, "abd"));
        Assert.NotEqual(first, PreviewCache.PathFor(session, 256, 640, "abc"));
        PreviewCache.MarkChanged(session);
        Assert.Equal(1, PreviewCache.Generation(session));
        Assert.NotEqual(first, PreviewCache.PathFor(session, 256, 1280, "abc"));
        Assert.NotEqual(PreviewCache.PathFor(new object(), 256, 1280, "abc"), PreviewCache.PathFor(session, 256, 1280, "abc"));
    }

    [Fact]
    public void Catalog_KnowsNewDomainsAndReadOnlyActions()
    {
        foreach (var category in new[] { "deck", "review", "compose", "diagram", "preview", "slide", "shape", "table", "chart" })
            Assert.True(CommandCatalog.All.ContainsKey(category), category);
        Assert.True(CommandCatalog.IsReadOnly("deck", "summary"));
        Assert.False(CommandCatalog.IsReadOnly("deck", "assign-ids"));
        Assert.True(CommandCatalog.IsReadOnly("preview", "snapshot"));
        Assert.False(CommandCatalog.IsReadOnly("shape", "add-text-box"));
        Assert.Contains("add-text-box", CommandCatalog.All["shape"].Actions);
        Assert.Equal("deck", CommandCatalog.All["deck"].McpTool);
    }

    [Fact]
    public void Catalog_ValidatesArgumentsThroughTheGeneratedRegistry()
    {
        Assert.Null(CommandCatalog.ValidateArguments("shape", "add-text-box", """{"slideIndex":1,"left":1,"top":1,"width":10,"height":10,"text":"x"}"""));
        Assert.Contains("Unknown or non-canonical parameter", CommandCatalog.ValidateArguments("shape", "add-text-box", """{"slide_index":1}"""), StringComparison.Ordinal);
        Assert.Contains("Unknown action 'add-textbox'", CommandCatalog.ValidateArguments("shape", "add-textbox", null), StringComparison.Ordinal);
        Assert.Contains("Unknown category 'shapes'", CommandCatalog.ValidateArguments("shapes", "x", null), StringComparison.Ordinal);
    }
}
