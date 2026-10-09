// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Design;

/// <summary>Typography normalization, profile color/font/size migration, and profile extraction (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Design")]
public sealed class StylePlannerTests
{
    private static readonly ResolvedProfile Default = DesignProfiles.ResolveBuiltIn("default");
    private static readonly ResolvedProfile Corporate = DesignProfiles.ResolveBuiltIn("corporate-blue");
    private static readonly ResolvedProfile HighContrast = DesignProfiles.ResolveBuiltIn("high-contrast");

    private static StyledRun Run(string text, string font, float size, string? role = "body", string? color = "#1A1A1A", bool theme = false, int slideId = 256, int shapeId = 2, int? row = null) =>
        new(1, slideId, shapeId, "Shape", role, row, row is null ? null : 1, 1, text.Length, font, size, color, theme, text);

    [Fact]
    public void Typography_UnifiesFontsByRoleAndKeepsSymbolFonts()
    {
        var (changes, _) = StylePlanner.PlanTypography(
        [
            Run("Revenue", "Arial", 30, role: "title"),
            Run("Body text", "Calibri", 16),
            Run("", "Wingdings", 16),
            Run("✓", "Wingdings", 16),
            Run("code()", "Courier New", 16),
        ], Default, new TypographyOptions());

        var fonts = changes.Where(change => change.Property == "font-name").ToList();
        Assert.Equal(["Segoe UI Semibold", "Segoe UI", "Consolas"], fonts.Select(change => change.To));
        Assert.DoesNotContain(changes, change => change.From == "Wingdings");
    }

    [Fact]
    public void Typography_SnapsNearSizesSetsTitleSizeAndRaisesSmallText()
    {
        var (changes, notes) = StylePlanner.PlanTypography(
        [
            Run("Title", "Segoe UI Semibold", 44, role: "title"),
            Run("Nearly body", "Segoe UI", 15),
            Run("Tiny", "Segoe UI", 7),
            Run("Huge callout", "Segoe UI", 60),
        ], Default, new TypographyOptions());

        var sizes = changes.Where(change => change.Property == "font-size").Select(change => (change.Sample, change.To)).ToList();
        Assert.Contains(("Title", "30"), sizes);
        Assert.Contains(("Nearly body", "16"), sizes);
        Assert.Contains(("Tiny", "12"), sizes);
        Assert.DoesNotContain(sizes, size => size.Sample == "Huge callout");
        Assert.Contains(notes, note => note.Contains("60", StringComparison.Ordinal));
        Assert.Contains(notes, note => note.Contains("min_font_size", StringComparison.Ordinal));
    }

    [Fact]
    public void ColorMap_MapsTokensAndClaimsUnchangedColors()
    {
        var map = StylePlanner.ColorMap(Default, Corporate);

        Assert.Equal(("#1F3864", "colors.primary"), map["#1F4E79"]);
        Assert.Equal("#1A1A1A", map["#1a1a1a"].To);
    }

    [Fact]
    public void PlanColors_SkipsThemeColorsWhenThemeIsUpdatedAndReportsUnmapped()
    {
        var map = StylePlanner.ColorMap(Default, HighContrast);
        var runs = new[]
        {
            Run("explicit primary", "Segoe UI", 16, color: "#1F4E79"),
            Run("theme primary", "Segoe UI", 16, color: "#1F4E79", theme: true),
            Run("odd color", "Segoe UI", 16, color: "#123456"),
        };
        var fills = new[] { new StyledFill(1, 256, 3, "Card", "fill", "#F2F4F7", false, "card@1") };
        var slides = new[] { new StyledSlide(1, 256, "default@1.0.0", "#FFFFFF") };

        var (changes, notes) = StylePlanner.PlanColors(runs, fills, slides, _ => map, themeUpdated: true);

        var text = Assert.Single(changes, change => change.Property == "text-color");
        Assert.Equal(("explicit primary", "#000000"), (text.Sample, text.To));
        Assert.Equal("#F2F2F2", Assert.Single(changes, change => change.Property == "fill").To);
        Assert.DoesNotContain(changes, change => change.Property == "background");
        Assert.Contains(notes, note => note.Contains("#123456", StringComparison.Ordinal));
        Assert.Contains(notes, note => note.Contains("follow the updated theme", StringComparison.Ordinal));
    }

    [Fact]
    public void FontAndSizeMigration_FollowsRolesAndLeavesThemeFontsToTheTheme()
    {
        var runs = new[]
        {
            Run("Title", "Segoe UI Semibold", 30, role: "title"),
            Run("Body", "Segoe UI", 16),
            Run("Theme body", "Calibri", 16),
            Run("Other", "Georgia", 16),
        };

        var (changes, notes) = StylePlanner.PlanFontAndSizeMigration(runs, _ => Default, HighContrast, fonts: true, sizes: true, ["Calibri"], themeUpdated: true);

        Assert.Equal(3, changes.Count(change => change.Property == "font-size"));
        Assert.All(changes.Where(change => change.Property == "font-size"), change => Assert.Equal("20", change.To));
        Assert.DoesNotContain(changes, change => change.Property == "font-name");
        Assert.Contains(notes, note => note.Contains("Georgia", StringComparison.Ordinal));
        Assert.Contains(notes, note => note.Contains("theme fonts", StringComparison.Ordinal));
    }

    [Fact]
    public void Extract_InfersSizesColorsAndMargins()
    {
        var runs = new List<StyledRun>
        {
            Run("Quarterly results", "Aptos Display", 36, role: "title"),
            Run(new string('x', 400), "Aptos", 18),
            Run(new string('y', 60), "Aptos", 12),
            Run("Cell", "Aptos", 11, role: "table", row: 1),
        };
        var sample = new DeckStyleSample(
            new Dictionary<string, string> { ["Accent1"] = "#0F6CBD", ["Dark1"] = "#000000", ["Light1"] = "#FFFFFF" },
            "Aptos Display",
            "Aptos",
            960,
            540,
            runs,
            [new StyledFill(1, 256, 4, "Panel", "fill", "#EEF2F7", false, null), new StyledFill(2, 257, 4, "Panel", "fill", "#EEF2F7", false, null)],
            [(60, 30, 840, 70), (60, 30, 840, 70)],
            [(60, 120, 840, 380)]);

        var (profile, notes) = ProfileExtractor.Extract(sample, "acme");
        var resolved = DesignProfiles.Resolve(profile);

        Assert.True(resolved.IsValid, string.Join("; ", resolved.Errors));
        Assert.Equal(36f, profile.TypeScale!["title"]);
        Assert.Equal(18f, profile.TypeScale["body"]);
        Assert.Equal(12f, profile.TypeScale["caption"]);
        Assert.Equal(11f, profile.Table!.FontSize);
        Assert.Equal("#0F6CBD", profile.Colors!["primary"]);
        Assert.Equal("#EEF2F7", profile.Colors["surface"]);
        Assert.Equal(60f, profile.Margins!["left"]);
        Assert.Equal("Aptos Display", resolved.Resolved!.HeadingFont);
        Assert.NotEmpty(notes);
    }
}
