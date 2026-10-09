// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Assets;
using Sbroenne.PowerPointMcp.Core.Catalog;
using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Assets;

/// <summary>Local asset catalog, picture quality checks, and capability discovery (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Assets")]
public sealed class AssetCatalogTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "pptmcp-assets-" + Guid.NewGuid().ToString("N"));

    public AssetCatalogTests()
    {
        Directory.CreateDirectory(Path.Combine(_root, "images", "команда"));
        File.WriteAllBytes(Path.Combine(_root, "images", "team-office.png"), ImageHeaderReaderTests.Png(1600, 900));
        File.WriteAllBytes(Path.Combine(_root, "images", "команда", "портрет.png"), ImageHeaderReaderTests.Png(800, 1200));
        File.WriteAllBytes(Path.Combine(_root, "images", "copy-of-team.png"), ImageHeaderReaderTests.Png(1600, 900));
        File.WriteAllText(Path.Combine(_root, "images", "notes.txt"), "not an image");
        File.WriteAllText(Path.Combine(_root, "images", "broken.png"), "not a png");
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    [Fact]
    public void Scan_AddsImagesSkipsBrokenAndRescansIncrementally()
    {
        var catalog = AssetCatalog.Load(Path.Combine(_root, "store"));

        var first = catalog.Scan(Path.Combine(_root, "images"), recursive: true);
        catalog.Save();
        var reloaded = AssetCatalog.Load(Path.Combine(_root, "store"));
        File.Delete(Path.Combine(_root, "images", "copy-of-team.png"));
        var second = reloaded.Scan(Path.Combine(_root, "images"), recursive: true);

        Assert.Equal((3, 1), (first.Added, first.Skipped));
        Assert.Equal((0, 2, 1), (second.Added, second.Unchanged, second.Removed));
        var portrait = reloaded.Entries.Single(entry => entry.Path.EndsWith("портрет.png", StringComparison.Ordinal));
        Assert.Equal(("portrait", 800, 1200), (portrait.Orientation, portrait.Width, portrait.Height));
    }

    [Fact]
    public void Search_RanksWordsAndFiltersOrientationAndTags()
    {
        var catalog = AssetCatalog.Load(Path.Combine(_root, "store"));
        catalog.Scan(Path.Combine(_root, "images"), recursive: true);
        var portrait = catalog.Entries.Single(entry => entry.Path.EndsWith("портрет.png", StringComparison.Ordinal));
        portrait.Tags.Add("people");
        portrait.Description = "Руководитель отдела у окна";

        Assert.Equal(2, catalog.Search("team", null, 0, null, 10).Count);
        Assert.Single(catalog.Search("команда", null, 0, null, 10));
        Assert.Single(catalog.Search("руководитель", null, 0, null, 10));
        Assert.Single(catalog.Search(null, "portrait", 0, null, 10));
        Assert.Empty(catalog.Search(null, null, 1000, "people", 10));
        Assert.Equal("team-office.png", Path.GetFileName(catalog.Search("team office", null, 0, null, 10)[0].Asset.Path));
    }

    [Fact]
    public void Duplicates_GroupIdenticalFiles()
    {
        var catalog = AssetCatalog.Load(Path.Combine(_root, "store"));
        catalog.Scan(Path.Combine(_root, "images"), recursive: true);

        var group = Assert.Single(catalog.Duplicates());

        Assert.Equal(2, group.Count);
    }

    [Fact]
    public void PictureReport_FlagsResolutionDistortionAltTextAndLinks()
    {
        var item = new DeckObjectInfo
        {
            SlideIndex = 2,
            SlideId = 257,
            ShapeId = 4,
            Name = "Hero",
            Kind = "picture",
            Left = 0,
            Top = 0,
            Width = 720,
            Height = 300,
            Rotation = 0,
            ZOrder = 1,
            Visible = true,
            HasText = false,
            SourcePixelSize = [800, 600],
            LinkSource = @"C:\missing\hero.png",
            LinkSourceExists = false,
            Crop = [0, 0, 0, 0],
        };

        var report = PictureQuality.Describe(item, pictureWidth: 720, pictureHeight: 300);

        Assert.Equal(80, report.EffectivePpi);
        Assert.True(report.DistortionPercent > 2);
        Assert.Equal(["missing-alt-text", "low-resolution", "distorted", "broken-link"], report.Issues);
    }

    [Fact]
    public void Capabilities_ListsToolsWorkflowsAndSettingsWithoutASession()
    {
        var report = Capabilities.Build("overview", "1.2.3", openSessions: 0, localUi: null);
        var tools = Capabilities.Build("tools", "1.2.3", 0, null);

        Assert.Contains(report.Tools!, tool => tool.Tool == "asset" && tool.ReadOnlyActions.Contains("inspect"));
        Assert.Contains(report.Tools!, tool => tool.Tool == "presentation" && tool.ReadOnlyActions.Contains("list"));
        Assert.Contains(report.Workflows!, workflow => workflow.Name == "build-slides");
        Assert.Contains("PPTMCP_LENIENT_ARGUMENTS", report.Settings.Keys);
        Assert.Contains("cards", report.CompositionKinds!);
        Assert.NotEmpty(tools.Tools!.Single(tool => tool.Tool == "design").AllActions);
        Assert.Null(Capabilities.Build("workflows", "1", 0, null).Tools);
    }
}
