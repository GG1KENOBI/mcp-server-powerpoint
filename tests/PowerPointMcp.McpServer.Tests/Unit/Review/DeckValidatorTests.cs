// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Review;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Review;

/// <summary>Validation rules over snapshots (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Review")]
public sealed class DeckValidatorTests
{
    private const float W = 960;
    private const float H = 540;

    internal static DeckSlideInfo Slide(int index = 1, int id = 256, string? title = "Title", string? layout = "Title and Content", string? background = "#FFFFFF") => new()
    {
        SlideIndex = index,
        SlideId = id,
        Title = title,
        LayoutName = layout,
        Hidden = false,
        ShapeCount = 0,
        BackgroundColor = background,
    };

    internal static DeckObjectInfo Obj(int id, float left, float top, float width, float height, string kind = "text-box",
        string? text = null, float? font = null, float[]? textBounds = null, string? role = null, int slideId = 256, int slideIndex = 1,
        float rotation = 0, bool topLevel = true, int? group = null, string? placeholder = null, string? fill = null, string? textColor = null,
        int[]? pixels = null, float[]? crop = null, string? alt = null, bool? linkExists = null, string[]? fonts = null,
        bool? wrap = true, string? appId = null, Dictionary<string, string>? tags = null, bool visible = true) => new()
        {
            SlideIndex = slideIndex,
            SlideId = slideId,
            ShapeIndex = topLevel ? id : null,
            ShapeId = id,
            Name = $"Shape {id}",
            Kind = kind,
            Role = role,
            Left = left,
            Top = top,
            Width = width,
            Height = height,
            Rotation = rotation,
            ZOrder = id,
            Visible = visible,
            HasText = text is not null,
            Text = text,
            MinFontSize = font,
            MaxFontSize = font,
            TextBounds = textBounds,
            WordWrap = text is null ? null : wrap,
            GroupShapeId = group,
            PlaceholderType = placeholder,
            FillColor = fill,
            TextColor = textColor,
            SourcePixelSize = pixels,
            Crop = crop,
            AltText = alt,
            LinkSourceExists = linkExists,
            LinkSource = linkExists is null ? null : @"C:\missing\photo.png",
            FontNames = fonts,
            AppId = appId,
            Tags = tags,
        };

    private static List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> One(params DeckObjectInfo[] objects) =>
        [(Slide(), objects)];

    private static IReadOnlyList<DeckFinding> Run(List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck, ValidationOptions? options = null) =>
        DeckValidator.Validate(W, H, deck, options ?? new ValidationOptions()).Findings;

    private static string[] Codes(IReadOnlyList<DeckFinding> findings) => findings.Select(finding => finding.Code).ToArray();

    [Fact]
    public void CleanSlide_HasNoFindings()
    {
        var findings = Run(One(
            Obj(2, 48, 30, 864, 60, text: "Title", font: 32, textBounds: [52, 34, 300, 40], role: "title"),
            Obj(3, 48, 120, 864, 300, text: "Body", font: 18, textBounds: [52, 124, 200, 30])));
        Assert.Empty(findings);
    }

    [Fact]
    public void TextOverflow_IsDeterministicMeasuredErrorWithEvidence()
    {
        var finding = Assert.Single(Run(One(Obj(3, 48, 120, 400, 50, text: "Long", font: 18, textBounds: [52, 124, 390, 120]))));
        Assert.Equal(("text-overflow", "error", "deterministic"), (finding.Code, finding.Severity, finding.Certainty));
        Assert.StartsWith("measured:", finding.Method, StringComparison.Ordinal);
        Assert.Equal("120", finding.Evidence!["textHeight"]);
        Assert.Equal("50", finding.Evidence["boxHeight"]);
        Assert.Equal([3], finding.ShapeIds);
        Assert.True(finding.AutoRepairable);
        Assert.Equal("text-overflow:256:3", finding.Id);
    }

    [Fact]
    public void TextOverflow_RotatedTextIsNotJudged()
    {
        var findings = Run(One(Obj(3, 48, 120, 400, 50, text: "Long", font: 18, textBounds: [52, 124, 390, 120], rotation: 90)));
        Assert.DoesNotContain(findings, finding => finding.Code == "text-overflow");
    }

    [Fact]
    public void OffSlide_PartialIsWarningAndRepairable_FullIsErrorAndNot()
    {
        var findings = Run(One(Obj(3, 900, 100, 100, 50, kind: "auto-shape"), Obj(4, 1000, 100, 100, 50, kind: "auto-shape")));
        var partial = Assert.Single(findings, finding => finding.ShapeIds![0] == 3);
        var full = Assert.Single(findings, finding => finding.ShapeIds![0] == 4);
        Assert.Equal(("warning", true), (partial.Severity, partial.AutoRepairable));
        Assert.Equal(("error", false), (full.Severity, full.AutoRepairable));
    }

    [Fact]
    public void OffSlide_UsesRotatedBounds()
    {
        // 400x40 bar rotated 90 degrees around its center at y=40 sticks out above the slide.
        var finding = Assert.Single(Run(One(Obj(3, 300, 20, 400, 40, kind: "auto-shape", rotation: 90))));
        Assert.Equal("off-slide", finding.Code);
    }

    [Fact]
    public void Backdrops_AreNotOffSlideOrOverlapping()
    {
        Assert.Empty(Run(One(
            Obj(2, -5, -5, 970, 550, kind: "picture", pixels: [970, 550], alt: "background photo"),
            Obj(3, 48, 120, 400, 60, text: "On the photo", font: 20, textBounds: [52, 124, 200, 30]))));
    }

    [Fact]
    public void SmallText_UsesFootnoteThresholdForFooters()
    {
        var findings = Run(One(
            Obj(3, 48, 120, 400, 60, text: "tiny", font: 9, textBounds: [52, 124, 100, 20]),
            Obj(4, 48, 500, 400, 20, text: "Confidential", font: 9, textBounds: [52, 502, 100, 12], role: "footer")));
        var finding = Assert.Single(findings);
        Assert.Equal(("small-text", 3), (finding.Code, finding.ShapeIds![0]));
    }

    [Fact]
    public void TextOverlap_IsErrorAndContainmentIsLayering()
    {
        var findings = Run(One(
            Obj(3, 48, 120, 400, 100, text: "A", font: 18, textBounds: [52, 124, 100, 20]),
            Obj(4, 48, 180, 400, 100, text: "B", font: 18, textBounds: [52, 184, 100, 20]),
            Obj(5, 500, 120, 300, 200, kind: "auto-shape", fill: "#EEEEEE"),
            Obj(6, 520, 140, 200, 60, text: "on card", font: 18, textBounds: [524, 144, 100, 20])));
        var finding = Assert.Single(findings);
        Assert.Equal("text-overlap", finding.Code);
        Assert.Equal([3, 4], finding.ShapeIds);
        Assert.Equal("deterministic", finding.Certainty);
    }

    [Fact]
    public void PartialOverlap_IsHeuristicAndCanBeAllowedByTag()
    {
        var findings = Run(One(Obj(3, 100, 100, 200, 200, kind: "auto-shape"), Obj(4, 250, 250, 200, 200, kind: "picture", pixels: [200, 200], alt: "x")));
        Assert.Equal(("partial-overlap", "heuristic"), (Assert.Single(findings).Code, findings[0].Certainty));

        var allowed = Run(One(Obj(3, 100, 100, 200, 200, kind: "auto-shape", tags: new() { ["PPTMCP_ALLOW_OVERLAP"] = "1" }),
            Obj(4, 250, 250, 200, 200, kind: "picture", pixels: [200, 200], alt: "x")));
        Assert.Empty(allowed);
    }

    [Fact]
    public void IncludeHeuristicFalse_DropsHeuristicFindings()
    {
        var deck = One(Obj(3, 100, 100, 200, 200, kind: "auto-shape"), Obj(4, 250, 250, 200, 200, kind: "auto-shape"));
        Assert.Empty(Run(deck, new ValidationOptions { IncludeHeuristic = false }));
    }

    [Fact]
    public void NearMisaligned_ReportsEachEdgeOnce()
    {
        var findings = Run(One(
            Obj(3, 48, 120, 400, 60, text: "A", font: 18, textBounds: [52, 124, 100, 20]),
            Obj(4, 51, 300, 400, 60, text: "B", font: 18, textBounds: [55, 304, 100, 20])));
        var codes = findings.Select(finding => finding.Id).ToArray();
        Assert.Contains("near-misaligned:256:3-4:left", codes);
        Assert.Contains("near-misaligned:256:3-4:right", codes);
        Assert.All(findings, finding => Assert.Equal("info", finding.Severity));
    }

    [Theory]
    [InlineData(1600, 900, 400f, 225f, false)]
    [InlineData(1600, 900, 400f, 300f, true)]
    [InlineData(1000, 1000, 300f, 306f, false)]
    public void ImageDistortion_ComparesSourceAndDisplayedAspect(int pw, int ph, float width, float height, bool distorted)
    {
        var findings = Run(One(Obj(3, 100, 100, width, height, kind: "picture", pixels: [pw, ph], alt: "photo")));
        Assert.Equal(distorted, findings.Any(finding => finding.Code == "image-distortion"));
    }

    [Fact]
    public void ImageDistortion_WithCropIsHeuristic()
    {
        // 1600x900 px = 1200x675 pt at 96 dpi; cropping 337.5 pt from each side leaves 525x675 (aspect 0.778).
        var exact = Run(One(Obj(3, 100, 100, 311.1f, 400, kind: "picture", pixels: [1600, 900], crop: [337.5f, 0, 337.5f, 0], alt: "a")));
        Assert.Empty(exact);
        var stretched = Assert.Single(Run(One(Obj(3, 100, 100, 400, 400, kind: "picture", pixels: [1600, 900], crop: [337.5f, 0, 337.5f, 0], alt: "a"))));
        Assert.Equal(("image-distortion", "heuristic"), (stretched.Code, stretched.Certainty));
    }

    [Fact]
    public void MissingAssetAltTextAndFonts_AreReported()
    {
        var findings = Run(One(
            Obj(3, 100, 100, 200, 100, kind: "picture", linkExists: false, alt: "linked"),
            Obj(4, 400, 100, 200, 100, kind: "chart"),
            Obj(5, 100, 300, 400, 60, text: "Привет", font: 18, textBounds: [104, 304, 100, 20], fonts: ["Calibri", "Corporate Sans", "+mn-lt"])),
            new ValidationOptions { InstalledFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "calibri" } });
        Assert.Contains("missing-asset", Codes(findings));
        Assert.Contains("missing-alt-text", Codes(findings));
        var font = Assert.Single(findings, finding => finding.Code == "missing-font");
        Assert.Equal("Corporate Sans", font.Evidence!["font"]);
    }

    [Fact]
    public void MissingFont_IsSkippedWithoutFontList()
    {
        var report = DeckValidator.Validate(W, H, One(Obj(5, 100, 300, 400, 60, text: "x", font: 18, fonts: ["Nope"])), new ValidationOptions());
        Assert.DoesNotContain(report.Findings, finding => finding.Code == "missing-font");
        Assert.Contains(report.NotChecked, note => note.StartsWith("missing-font", StringComparison.Ordinal));
    }

    [Fact]
    public void EmptyPlaceholder_IgnoresFooterPlaceholders()
    {
        var findings = Run(One(
            Obj(3, 48, 120, 400, 200, kind: "placeholder", placeholder: "ppPlaceholderBody", role: "body"),
            Obj(4, 48, 500, 200, 20, kind: "placeholder", placeholder: "ppPlaceholderFooter", role: "footer")));
        var finding = Assert.Single(findings);
        Assert.Equal(("empty-placeholder", 3), (finding.Code, finding.ShapeIds![0]));
    }

    [Fact]
    public void LowContrast_UsesFillThenContainerThenBackground()
    {
        var findings = Run(One(
            Obj(3, 48, 120, 300, 60, text: "light on white", font: 14, textBounds: [52, 124, 100, 20], textColor: "#CCCCCC"),
            Obj(4, 400, 100, 400, 300, kind: "auto-shape", fill: "#1F3864"),
            Obj(5, 420, 120, 300, 60, text: "navy on navy", font: 14, textBounds: [424, 124, 100, 20], textColor: "#203864"),
            Obj(6, 420, 250, 300, 60, text: "white on navy", font: 14, textBounds: [424, 254, 100, 20], textColor: "#FFFFFF")));
        var low = findings.Where(finding => finding.Code == "low-contrast").Select(finding => finding.ShapeIds![0]).ToArray();
        Assert.Equal([3, 5], low);
        Assert.Equal("shape 4", findings.First(finding => finding.ShapeIds![0] == 5).Evidence!["backgroundSource"]);
    }

    [Fact]
    public void LargeText_UsesLowerContrastThreshold()
    {
        // #909090 on white is about 3.2:1 — fails for body text, passes for 24 pt text.
        Assert.Single(Run(One(Obj(3, 48, 120, 300, 60, text: "x", font: 14, textBounds: [52, 124, 10, 20], textColor: "#909090"))));
        Assert.Empty(Run(One(Obj(3, 48, 120, 300, 60, text: "x", font: 24, textBounds: [52, 124, 10, 20], textColor: "#909090"))));
    }

    [Fact]
    public void MissingTitle_AndDuplicateTitles()
    {
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, title: null), [Obj(3, 48, 120, 300, 60, kind: "auto-shape", slideId: 256)]),
            (Slide(2, 257, title: "Results"), []),
            (Slide(3, 258, title: "Results"), []),
            (Slide(4, 259, title: "Results (cont.)"), []),
        };
        var findings = Run(deck);
        Assert.Equal("missing-title", Assert.Single(findings, finding => finding.Code == "missing-title").Code);
        var duplicate = Assert.Single(findings, finding => finding.Code == "duplicate-title");
        Assert.Equal([2, 3], duplicate.SlideIndexes);
    }

    [Fact]
    public void InconsistentTitlePosition_FindsTheOutlier()
    {
        DeckObjectInfo Title(int slideId, int slideIndex, float left, float top) =>
            Obj(2, left, top, 800, 60, kind: "placeholder", text: "T", font: 32, textBounds: [left + 4, top + 4, 100, 40], role: "title", slideId: slideId, slideIndex: slideIndex);
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, title: "A"), [Title(256, 1, 48, 30)]),
            (Slide(2, 257, title: "B"), [Title(257, 2, 48, 30)]),
            (Slide(3, 258, title: "C"), [Title(258, 3, 60, 42)]),
        };
        var finding = Assert.Single(Run(deck), finding => finding.Code == "inconsistent-title-position");
        Assert.Equal((258, "48", "30"), (finding.SlideId, finding.Evidence!["expectedLeft"], finding.Evidence["expectedTop"]));
    }

    [Fact]
    public void FooterInconsistency_ReportsMissingAndDifferentFooters()
    {
        DeckObjectInfo Footer(int slideId, int slideIndex, string text) =>
            Obj(9, 300, 510, 360, 20, kind: "placeholder", text: text, font: 10, textBounds: [304, 512, 100, 12], role: "footer", slideId: slideId, slideIndex: slideIndex);
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, title: "Cover", layout: "Title Slide"), []),
            (Slide(2, 257, title: "B"), [Footer(257, 2, "ACME — Confidential")]),
            (Slide(3, 258, title: "C"), [Footer(258, 3, "ACME — Confidential")]),
            (Slide(4, 259, title: "D"), [Footer(259, 4, "ACME - Draft")]),
            (Slide(5, 260, title: "E"), []),
        };
        var findings = Run(deck).Where(finding => finding.Code == "footer-inconsistent").ToList();
        Assert.Equal(2, findings.Count);
        Assert.Contains(findings, finding => finding.SlideIndexes!.SequenceEqual([5]));
        Assert.Contains(findings, finding => finding.SlideIndexes!.SequenceEqual([4]) && finding.Severity == "warning");
    }

    [Fact]
    public void DuplicateFooter_OnOneSlide()
    {
        var finding = Assert.Single(Run(One(
            Obj(8, 300, 510, 360, 20, kind: "placeholder", text: "ACME", font: 10, textBounds: [304, 512, 50, 12], role: "footer"),
            Obj(9, 600, 500, 300, 30, text: "ACME", font: 10, textBounds: [604, 504, 50, 12]))), finding => finding.Code == "duplicate-footer");
        Assert.Equal([8, 9], finding.ShapeIds);
    }

    [Fact]
    public void DuplicateAppIds_AreRepairable()
    {
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256), [Obj(3, 48, 120, 300, 60, kind: "auto-shape", appId: "kpi")]),
            (Slide(2, 257), [Obj(3, 48, 120, 300, 60, kind: "auto-shape", appId: "kpi", slideId: 257, slideIndex: 2)]),
        };
        var finding = Assert.Single(Run(deck), finding => finding.Code == "duplicate-app-id");
        Assert.True(finding.AutoRepairable);
        Assert.Equal([1, 2], finding.SlideIndexes);
    }

    [Fact]
    public void HiddenObjectsAndMembersOfHiddenGroups_AreIgnored()
    {
        Assert.Empty(Run(One(
            Obj(3, 48, 120, 400, 50, text: "Long", font: 6, textBounds: [52, 124, 390, 120], visible: false),
            Obj(10, 100, 300, 300, 100, kind: "group", visible: false),
            Obj(11, 100, 300, 300, 50, text: "tiny", font: 6, textBounds: [104, 304, 100, 10], topLevel: false, group: 10))));
    }

    [Fact]
    public void RuleFilter_RunsOnlyNamedRules()
    {
        var report = DeckValidator.Validate(W, H, One(
            Obj(3, 900, 100, 100, 50, text: "x", font: 6, textBounds: [904, 104, 20, 10])),
            new ValidationOptions { Rules = new HashSet<string> { "small-text" } });
        Assert.Equal(["small-text"], report.RulesRun);
        Assert.Equal(["small-text"], Codes(report.Findings));
    }

    [Fact]
    public void SafeArea_IsOptInAndInfo()
    {
        var deck = One(Obj(3, 10, 100, 300, 60, kind: "auto-shape"));
        Assert.Empty(Run(deck));
        var finding = Assert.Single(Run(deck, new ValidationOptions { SafeMargin = 24 }));
        Assert.Equal(("outside-safe-area", "info"), (finding.Code, finding.Severity));
    }

    [Fact]
    public void Findings_AreOrderedBySeverity()
    {
        var findings = Run(One(
            Obj(3, 48, 120, 400, 60, text: "A", font: 18, textBounds: [52, 124, 100, 20]),
            Obj(4, 51, 300, 400, 60, text: "B", font: 8, textBounds: [55, 304, 100, 20]),
            Obj(5, 48, 400, 400, 20, text: "C", font: 18, textBounds: [52, 404, 300, 60])));
        var ranks = findings.Select(finding => finding.Severity switch { "error" => 0, "warning" => 1, _ => 2 }).ToArray();
        Assert.Equal(ranks.Order(), ranks);
        Assert.Equal(["error", "warning", "info"], findings.Select(finding => finding.Severity).Distinct());
    }

    [Fact]
    public void RuleCatalog_IsCompleteAndConsistent()
    {
        Assert.Equal(DeckValidator.Rules.Count, DeckValidator.Rules.Select(rule => rule.Code).Distinct().Count());
        Assert.All(DeckValidator.Rules, rule => Assert.True(rule.Certainty is "deterministic" or "heuristic"));
        Assert.All(DeckValidator.Rules.Where(rule => rule.RepairedByDefault), rule => Assert.True(rule.AutoRepairable));
    }
}
