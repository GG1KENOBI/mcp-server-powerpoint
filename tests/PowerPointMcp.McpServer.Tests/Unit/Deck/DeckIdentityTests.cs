// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Deck.DeckTestData;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Deck;

/// <summary>Fingerprints, structural diffs, persistent-id planning, and aspect ratios (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Deck")]
public sealed class DeckIdentityTests
{
    [Fact]
    public void Fingerprint_IsDeterministicAndOrderIndependentWithinSlide()
    {
        var objects = SampleDeck()[0].Objects;
        var forward = DeckFingerprint.ForSlide(256, objects);
        var reversed = DeckFingerprint.ForSlide(256, objects.Reverse());
        Assert.Equal(forward, reversed);
        Assert.Matches("^[0-9a-f]{16}$", forward);
    }

    [Fact]
    public void Fingerprint_IgnoresSubPrecisionNoiseButSeesRealMoves()
    {
        var a = DeckFingerprint.ForSlide(1, [Shape(1, 1, 2, "Box", left: 10.01f)]);
        var b = DeckFingerprint.ForSlide(1, [Shape(1, 1, 2, "Box", left: 10.02f)]);
        var c = DeckFingerprint.ForSlide(1, [Shape(1, 1, 2, "Box", left: 12f)]);
        Assert.Equal(a, b);
        Assert.NotEqual(a, c);
    }

    [Fact]
    public void Fingerprint_ChangesWithText()
    {
        var a = DeckFingerprint.ForSlide(1, [Shape(1, 1, 2, "Box", text: "one")]);
        var b = DeckFingerprint.ForSlide(1, [Shape(1, 1, 2, "Box", text: "two")]);
        Assert.NotEqual(a, b);
    }

    [Fact]
    public void DeckFingerprint_ChangesWhenSlidesReorder()
    {
        var a = DeckFingerprint.ForDeck([(256, "aaaa"), (260, "bbbb")]);
        var b = DeckFingerprint.ForDeck([(260, "bbbb"), (256, "aaaa")]);
        Assert.NotEqual(a, b);
    }

    [Fact]
    public void Diff_ReportsMovesAddsRemovesAndModifiedFields()
    {
        var before = SampleDeck();
        var after = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 260, "Q3 Results"),
            [
                Shape(1, 260, 2, "Title 1", "placeholder", "Q3 Results (final)", role: "title", minFont: 40),
                Shape(1, 260, 4, "Chart 3", "chart", appId: "q3-chart", left: 60, top: 120, width: 400, height: 300),
                Shape(1, 260, 9, "Note", "text-box", "Source: finance"),
            ]),
            (Slide(2, 270, "New"), []),
        };

        var changes = DeckDiff.Compare(before, after);
        string[] summary = changes.Select(change => $"{change.Change}:{change.SlideId}:{change.ShapeId}:{string.Join(",", change.Fields ?? [])}").ToArray();

        Assert.Contains("slide-moved:260::", summary);
        Assert.Contains("object-modified:260:2:text", summary);
        Assert.Contains("object-modified:260:4:left", summary);
        Assert.Contains("object-added:260:9:", summary);
        Assert.Contains("object-removed:260:5:", summary);
        Assert.Contains("slide-added:270::", summary);
        Assert.Contains("slide-removed:256::", summary);
    }

    [Fact]
    public void Diff_IdenticalSnapshotsHaveNoChanges()
    {
        Assert.Empty(DeckDiff.Compare(SampleDeck(), SampleDeck()));
    }

    [Fact]
    public void FindDuplicates_SeparatesSlideAndObjectNamespaces()
    {
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, appId: "intro"), [Shape(1, 256, 2, "A", appId: "kpi")]),
            (Slide(2, 257, appId: "intro"), [Shape(2, 257, 2, "A", appId: "kpi"), Shape(2, 257, 3, "B", appId: "intro")]),
        };
        Assert.Equal(["slide:intro", "object:kpi"], AppIdPlanner.FindDuplicates(deck));
    }

    [Fact]
    public void Plan_AssignsMissingSlideIdsAndRenamesLaterDuplicates()
    {
        // Slide 2 is a copy of slide 1: same slide and object ids.
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, appId: "intro"), [Shape(1, 256, 2, "A", appId: "kpi")]),
            (Slide(2, 257, appId: "intro"), [Shape(2, 257, 2, "A", appId: "kpi")]),
            (Slide(3, 258), [Shape(3, 258, 4, "Chart", "chart")]),
        };

        var plan = AppIdPlanner.Plan(deck, null, "id");

        Assert.Collection(plan,
            item => Assert.Equal((257, (int?)null, "intro", "intro~2", "duplicate"), (item.SlideId, item.ShapeId, item.OldAppId, item.NewAppId, item.Reason)),
            item => Assert.Equal((257, (int?)2, "kpi", "kpi~2", "duplicate"), (item.SlideId, item.ShapeId, item.OldAppId, item.NewAppId, item.Reason)),
            item => Assert.Equal((258, (int?)null, (string?)null, "id-slide-258", "missing"), (item.SlideId, item.ShapeId, item.OldAppId, item.NewAppId, item.Reason)));
    }

    [Fact]
    public void Plan_TagsSelectedObjectsByRoleOrKindWithoutCollisions()
    {
        var chart = Shape(1, 256, 4, "Chart", "chart");
        var title = Shape(1, 256, 2, "Title 1", "placeholder", role: "title");
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, appId: "s1"), [title, chart, Shape(1, 256, 7, "Taken", appId: "deck-chart-256-4")]),
        };

        var plan = AppIdPlanner.Plan(deck, [chart, title], "deck");

        Assert.Equal(["deck-title-256-2", "deck-chart-256-4~2"], plan.Select(item => item.NewAppId));
        Assert.All(plan, item => Assert.Equal("missing", item.Reason));
    }

    [Fact]
    public void Plan_IsIdempotentAfterApplying()
    {
        var deck = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>
        {
            (Slide(1, 256, appId: "id-slide-256"), [Shape(1, 256, 2, "A", appId: "a")]),
        };
        Assert.Empty(AppIdPlanner.Plan(deck, null, "id"));
    }

    [Theory]
    [InlineData(960f, 540f, "16:9")]
    [InlineData(720f, 540f, "4:3")]
    [InlineData(720f, 450f, "16:10")]
    [InlineData(612f, 792f, "0.773:1")]
    [InlineData(100f, 0f, "unknown")]
    public void AspectRatio_LabelsCommonSizes(float width, float height, string expected)
    {
        Assert.Equal(expected, DeckGeometry.AspectRatio(width, height));
    }
}
