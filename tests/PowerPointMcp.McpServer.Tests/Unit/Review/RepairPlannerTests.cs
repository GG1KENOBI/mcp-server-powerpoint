// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Review;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Review.DeckValidatorTests;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Review;

/// <summary>Repair planning over snapshots (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Review")]
public sealed class RepairPlannerTests
{
    private const float W = 960;
    private const float H = 540;

    private static RepairPlan Plan(RepairOptions? options, params DeckObjectInfo[] objects)
    {
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> { (Slide(), objects) };
        var findings = DeckValidator.Validate(W, H, deck, new ValidationOptions()).Findings;
        return RepairPlanner.Plan(W, H, deck, findings, options ?? new RepairOptions());
    }

    [Fact]
    public void Overflow_ExpandsIntoFreeSpaceUsingMeasuredHeight()
    {
        var plan = Plan(null, Obj(3, 48, 120, 400, 50, text: "Long", font: 18, textBounds: [52, 124, 390, 120]));
        var action = Assert.Single(plan.Actions);
        Assert.Equal(("resize", "measured"), (action.Operation, action.Basis));
        Assert.Equal([48f, 120f, 400f, 128f], action.After); // 120 + 7.2 margins, rounded up
        Assert.Equal([48f, 120f, 400f, 50f], action.Before);
    }

    [Fact]
    public void Overflow_ShrinksWhenBlockedBelow_NeverBelowMinimum()
    {
        var plan = Plan(null,
            Obj(3, 48, 120, 400, 100, text: "Long", font: 20, textBounds: [52, 124, 390, 130]),
            Obj(4, 48, 240, 400, 200, kind: "chart", alt: "chart"));
        Assert.Collection(plan.Actions,
            grow => Assert.Equal(("resize", 108f), (grow.Operation, grow.After![3])),
            shrink =>
            {
                Assert.Equal(("scale-font", "estimated"), (shrink.Operation, shrink.Basis));
                Assert.InRange(shrink.FontSizeAfter!.Value, 10f, 19.5f);
                Assert.Equal(10f, shrink.MinFontSize);
            });
    }

    [Fact]
    public void Overflow_ThatNeedsTextBelowMinimum_IsUnresolvedWithReason()
    {
        var plan = Plan(new RepairOptions { MinFontSize = 12 },
            Obj(3, 48, 120, 400, 40, text: "Very long", font: 12, textBounds: [52, 124, 390, 300]),
            Obj(4, 48, 170, 400, 200, kind: "chart", alt: "chart"));
        Assert.Empty(plan.Actions);
        var unresolved = Assert.Single(plan.Unresolved, item => item.Finding.Code == "text-overflow");
        Assert.Contains("below the 12 pt minimum", unresolved.Reason, StringComparison.Ordinal);
        Assert.Contains("Split", unresolved.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void Overflow_ExpandOnlyPolicy_DoesNotShrink()
    {
        var plan = Plan(new RepairOptions { FitPolicy = ["expand"] },
            Obj(3, 48, 120, 400, 100, text: "Long", font: 20, textBounds: [52, 124, 390, 130]),
            Obj(4, 48, 240, 400, 200, kind: "chart", alt: "chart"));
        Assert.All(plan.Actions, action => Assert.NotEqual("scale-font", action.Operation));
    }

    [Fact]
    public void OffSlide_MovesInsideKeepingSize()
    {
        var action = Assert.Single(Plan(null, Obj(3, 900, -10, 100, 50, kind: "auto-shape")).Actions);
        Assert.Equal("move", action.Operation);
        Assert.Equal([860f, 0f, 100f, 50f], action.After);
    }

    [Fact]
    public void OffSlide_TooLargeOrEntirelyOff_IsUnresolved()
    {
        var plan = Plan(null, Obj(3, -10, 100, 1000, 50, kind: "auto-shape"), Obj(4, 2000, 100, 100, 50, kind: "auto-shape"));
        Assert.Empty(plan.Actions);
        Assert.Equal(2, plan.Unresolved.Count);
    }

    [Fact]
    public void TextOverlap_MovesLowerShapeBelowWithGap()
    {
        var plan = Plan(null,
            Obj(3, 48, 120, 400, 100, text: "A", font: 18, textBounds: [52, 124, 100, 20]),
            Obj(4, 48, 180, 400, 100, text: "B", font: 18, textBounds: [52, 184, 100, 20]));
        var action = Assert.Single(plan.Actions);
        Assert.Equal((4, "move"), (action.ShapeId, action.Operation));
        Assert.Equal(232f, action.After![1]);
    }

    [Fact]
    public void TextOverlap_WithoutRoom_IsUnresolved()
    {
        var plan = Plan(null,
            Obj(3, 48, 300, 400, 200, text: "A", font: 18, textBounds: [52, 304, 100, 20]),
            Obj(4, 48, 400, 400, 120, text: "B", font: 18, textBounds: [52, 404, 100, 20]));
        Assert.DoesNotContain(plan.Actions, action => action.Code == "text-overlap");
        Assert.Contains(plan.Unresolved, item => item.Finding.Code == "text-overlap");
    }

    [Fact]
    public void ImageDistortion_FitsInsideFrameCentered()
    {
        var action = Assert.Single(Plan(null, Obj(3, 100, 100, 400, 300, kind: "picture", pixels: [1600, 900], alt: "x")).Actions);
        Assert.Equal("move-resize", action.Operation);
        Assert.Equal([100f, 137.5f, 400f, 225f], action.After);
    }

    [Fact]
    public void NearMisaligned_MovesSecondShapeOnly_OncePerObject()
    {
        var plan = Plan(null,
            Obj(3, 48, 120, 400, 60, text: "A", font: 18, textBounds: [52, 124, 100, 20]),
            Obj(4, 51, 300, 400, 60, text: "B", font: 18, textBounds: [55, 304, 100, 20]));
        var action = Assert.Single(plan.Actions);
        Assert.Equal((4, 48f), (action.ShapeId, action.After![0]));
        Assert.Contains(plan.Unresolved, item => item.Reason.Contains("Another repair", StringComparison.Ordinal));
    }

    [Fact]
    public void Codes_RestrictWhatIsPlanned_AndSmallTextIsOptIn()
    {
        var shape = Obj(3, 48, 120, 400, 60, text: "tiny", font: 8, textBounds: [52, 124, 100, 12]);
        Assert.Empty(Plan(null, shape).Actions);
        var action = Assert.Single(Plan(new RepairOptions { Codes = new HashSet<string> { "small-text" } }, shape).Actions);
        Assert.Equal(("scale-font", 10f), (action.Operation, action.FontSizeAfter!.Value));
    }

    [Fact]
    public void NonRepairableFindings_AreReturnedWithSuggestion()
    {
        var plan = Plan(new RepairOptions { Codes = new HashSet<string> { "empty-placeholder" } },
            Obj(3, 48, 120, 400, 200, kind: "placeholder", placeholder: "ppPlaceholderBody", role: "body"));
        var unresolved = Assert.Single(plan.Unresolved);
        Assert.StartsWith("Not auto-repairable.", unresolved.Reason, StringComparison.Ordinal);
        Assert.Contains("never deletes", unresolved.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void FindingIds_SelectIndividualRepairs()
    {
        var objects = new[]
        {
            Obj(3, 900, 100, 100, 50, kind: "auto-shape"),
            Obj(4, 900, 300, 100, 50, kind: "auto-shape"),
        };
        var plan = Plan(new RepairOptions { FindingIds = new HashSet<string> { "off-slide:256:4" } }, objects);
        Assert.Equal(4, Assert.Single(plan.Actions).ShapeId);
    }

    [Fact]
    public void DuplicateIds_RenameLaterCopies()
    {
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256), [Obj(3, 48, 120, 300, 60, kind: "auto-shape", appId: "kpi")]),
            (Slide(2, 257), [Obj(3, 48, 120, 300, 60, kind: "auto-shape", appId: "kpi", slideId: 257, slideIndex: 2)]),
        };
        var findings = DeckValidator.Validate(W, H, deck, new ValidationOptions()).Findings;
        var action = Assert.Single(RepairPlanner.Plan(W, H, deck, findings, new RepairOptions()).Actions);
        Assert.Equal(("assign-id", 257, 3, "kpi~2"), (action.Operation, action.SlideId, action.ShapeId, action.NewAppId));
    }

    [Fact]
    public void Plan_IsDeterministic()
    {
        DeckObjectInfo[] objects =
        [
            Obj(3, 48, 120, 400, 50, text: "Long", font: 18, textBounds: [52, 124, 390, 120]),
            Obj(4, 900, -10, 100, 50, kind: "auto-shape"),
        ];
        var first = Plan(null, objects).Actions.Select(action => $"{action.ShapeId}:{action.Operation}:{string.Join(',', action.After ?? [])}");
        var second = Plan(null, objects).Actions.Select(action => $"{action.ShapeId}:{action.Operation}:{string.Join(',', action.After ?? [])}");
        Assert.Equal(first, second);
    }
}
