using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Preview;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Slide;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for PowerPoint-rendered previews.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Preview")]
public class PreviewCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly PreviewCommands _preview = new();

    public PreviewCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    private void FreshWideDeck()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new PageSetupCommands().SetSize(_fixture.Batch, 960, 540).Success);
    }

    [Fact]
    public void Snapshot_RendersPngWithMetadataAndCaches()
    {
        FreshWideDeck();
        Assert.True(new ComposeCommands().Create(_fixture.Batch, spec: CompositionKinds.Find("kpis")!.Example).Success);

        var first = _preview.Snapshot(_fixture.Batch, slideIndex: 2, width: 960);
        Assert.True(first.Success, first.ErrorMessage);
        Assert.True(File.Exists(first.ImagePath));
        Assert.Equal((960, 540), (first.PixelWidth!.Value, first.PixelHeight!.Value));
        var header = Sbroenne.PowerPointMcp.Core.Assets.ImageHeaderReader.ReadFile(first.ImagePath!)!;
        Assert.Equal(("png", 960, 540), (header.Format, header.Width, header.Height));
        Assert.False(first.Cached);
        Assert.NotEmpty(first.Objects!);
        Assert.Contains("kpi", first.Sketch, StringComparison.Ordinal);

        var second = _preview.Snapshot(_fixture.Batch, slideIndex: 2, width: 960);
        Assert.True(second.Cached);
        Assert.Equal(first.ImagePath, second.ImagePath);
    }

    [Fact]
    public void Snapshot_ExplicitPathRequiresOverwriteToReplace()
    {
        FreshWideDeck();
        var path = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"слайд-{Guid.NewGuid():N}.png");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        Assert.True(_preview.Snapshot(_fixture.Batch, slideIndex: 1, outputPath: path).Success);
        Assert.False(_preview.Snapshot(_fixture.Batch, slideIndex: 1, outputPath: path).Success);
        Assert.True(_preview.Snapshot(_fixture.Batch, slideIndex: 1, outputPath: path, overwrite: true).Success);
    }

    [Fact]
    public void ContactSheet_ComposesAllSlides()
    {
        FreshWideDeck();
        var slides = new SlideCommands();
        for (int i = 0; i < 5; i++) Assert.True(slides.AddBlank(_fixture.Batch).Success);
        Assert.True(slides.SetHidden(_fixture.Batch, 3, true).Success);

        var sheet = _preview.ContactSheet(_fixture.Batch, columns: 3, thumbnailWidth: 200);

        Assert.True(sheet.Success, sheet.ErrorMessage);
        Assert.Equal(6, sheet.Slides!.Count);
        Assert.True(sheet.Slides.Single(slide => slide.SlideIndex == 3).Hidden);
        Assert.Equal((2, 2), (sheet.Slides[4].Row, sheet.Slides[4].Column));
        Assert.True(File.Exists(sheet.ImagePath));
    }

    [Fact]
    public void CompareImages_LocatesTheChange()
    {
        FreshWideDeck();
        var before = _preview.Snapshot(_fixture.Batch, slideIndex: 1, width: 960, refresh: true);
        var copy = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"before-{Guid.NewGuid():N}.png");
        File.Copy(before.ImagePath!, copy);
        var rectangle = new ShapeCommands().AddRectangle(_fixture.Batch, 1, 600, 300, 100, 80);
        Assert.True(rectangle.Success);
        Sbroenne.PowerPointMcp.Core.Preview.PreviewCache.MarkChanged(_fixture.Batch);
        var after = _preview.Snapshot(_fixture.Batch, slideIndex: 1, width: 960);

        var diff = _preview.CompareImages(_fixture.Batch, copy, after.ImagePath!);

        Assert.True(diff.Success, diff.ErrorMessage);
        Assert.True(diff.Difference!.ChangedPixels > 0);
        Assert.InRange(diff.ChangedArea![0], 590f, 610f);
        Assert.InRange(diff.ChangedArea[1], 290f, 310f);
    }
}
