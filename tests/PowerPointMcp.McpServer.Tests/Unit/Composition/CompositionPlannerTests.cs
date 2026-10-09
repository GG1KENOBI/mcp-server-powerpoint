// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Review;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Composition;

/// <summary>Composition planning for every kind (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Composition")]
public sealed class CompositionPlannerTests
{
    private static readonly ResolvedProfile Profile = DesignProfiles.ResolveBuiltIn("default");

    public static TheoryData<string, float, float> KindsAndSizes()
    {
        var data = new TheoryData<string, float, float>();
        foreach (var kind in CompositionKinds.All)
        {
            data.Add(kind.Name, 960, 540);
            data.Add(kind.Name, 720, 540);
        }
        return data;
    }

    internal static CompositionPlan PlanJson(string json, float width = 960, float height = 540, IReadOnlyList<int>? partition = null, string profile = "default")
    {
        var parsed = CompositionParser.Parse(json);
        Assert.True(parsed.Spec is not null, string.Join("; ", parsed.Errors));
        var context = new PlanContext
        {
            AppId = "s1",
            ImageSizes = new Dictionary<string, (int, int)> { [@"C:\Assets\plant.jpg"] = (1600, 900), ["a.png"] = (800, 800) },
            Partition = partition,
        };
        return CompositionPlanner.Plan(parsed.Spec!, DesignProfiles.ResolveBuiltIn(profile), width, height, context);
    }

    [Theory]
    [MemberData(nameof(KindsAndSizes))]
    public void EveryKind_PlansInsideTheSlideWithUniqueKeys(string kind, float width, float height)
    {
        var plan = PlanJson(CompositionKinds.Find(kind)!.Example, width, height);
        var slide = Assert.Single(plan.Slides);
        Assert.Equal(slide.Elements.Count, slide.Elements.Select(element => element.Key).Distinct().Count());
        var bounds = new Box(-0.5f, -0.5f, width + 0.5f, height + 0.5f);
        Assert.All(slide.Elements.Where(element => element.Type != "connector"), element =>
            Assert.True(bounds.Contains(element.Box, 0), $"{kind}: {element.Key} at {element.Box} is outside {width}x{height}"));
        Assert.All(slide.Elements, element => Assert.True(element.Box.Width >= 0 && element.Box.Height >= 0, $"{element.Key} has a negative size"));
        Assert.All(slide.Regions.SelectMany(region => region.Keys), key => Assert.Contains(slide.Elements, element => element.Key == key));
    }

    [Theory]
    [MemberData(nameof(KindsAndSizes))]
    public void EveryKind_HasNoPlannedTextOverlapOrOffSlide(string kind, float width, float height)
    {
        var plan = PlanJson(CompositionKinds.Find(kind)!.Example, width, height);
        var snapshot = Snapshot(plan.Slides[0]);
        var report = DeckValidator.Validate(width, height, [(snapshot.Slide, snapshot.Objects)],
            new ValidationOptions { Rules = new HashSet<string> { "off-slide", "text-overlap" } });
        Assert.True(report.Findings.Count == 0, $"{kind}: " + string.Join(" | ", report.Findings.Select(finding => finding.Message)));
    }

    [Fact]
    public void Plans_AreDeterministic()
    {
        foreach (var kind in CompositionKinds.All)
        {
            var first = Describe(PlanJson(kind.Example));
            var second = Describe(PlanJson(kind.Example));
            Assert.Equal(first, second);
        }
    }

    [Fact]
    public void ExecutiveSummary_SplitsLongContentWithContinuationTitles()
    {
        var points = string.Join(",", Enumerable.Range(1, 14).Select(i => $$"""{"text":"Point {{i}} with a fairly long headline that wraps","detail":"Supporting detail sentence number {{i}} that adds context and also wraps onto a second line."}"""));
        var plan = PlanJson($$"""{"kind":"executive-summary","title":"Итоги года","points":[{{points}}],"takeaway":"Plan achieved"}""");
        Assert.True(plan.Slides.Count >= 2);
        Assert.Equal("Итоги года", plan.Slides[0].Title);
        Assert.All(plan.Slides.Skip(1), slide => Assert.Equal("Итоги года (продолжение)", slide.Title));
        Assert.Equal(14, plan.Slides.Sum(slide => slide.ItemCount));
        Assert.Equal(["s1", "s1~2"], plan.Slides.Take(2).Select(slide => slide.AppId));
        Assert.Contains(plan.Slides[^1].Elements, element => element.Key == "takeaway");
        Assert.DoesNotContain(plan.Slides[0].Elements, element => element.Key == "takeaway");
        // Numbering continues across slides.
        Assert.Contains(plan.Slides[1].Elements, element => element.Key == $"point-{plan.Slides[0].ItemCount + 1}");
        // Text is never shortened.
        var allText = plan.Slides.SelectMany(slide => slide.Elements).SelectMany(element => element.Paragraphs ?? []).Select(paragraph => paragraph.Text).ToList();
        Assert.Contains("Supporting detail sentence number 14 that adds context and also wraps onto a second line.", allText);
    }

    [Fact]
    public void ExecutiveSummary_ShrinkOnlyPolicy_DoesNotSplit()
    {
        var points = string.Join(",", Enumerable.Range(1, 14).Select(i => $$"""{"text":"Point {{i}} with a fairly long headline that wraps","detail":"Supporting detail sentence {{i}} that adds context."}"""));
        var plan = PlanJson($$$"""{"kind":"executive-summary","title":"T","points":[{{{points}}}],"fit":{"policy":["shrink"],"min_font_size":12}}""");
        Assert.Single(plan.Slides);
        Assert.All(plan.Slides[0].Elements.SelectMany(element => element.Paragraphs ?? []).Where(paragraph => paragraph.Size > 0 && paragraph.Text.StartsWith("Point", StringComparison.Ordinal)),
            paragraph => Assert.True(paragraph.Size >= 12f));
        Assert.Contains(plan.Warnings, warning => warning.Contains("overflow", StringComparison.Ordinal));
    }

    [Fact]
    public void ForcedPartition_IsHonoured()
    {
        var plan = PlanJson("""{"kind":"executive-summary","title":"T","points":["a","b","c","d","e"]}""", partition: [2, 3]);
        Assert.Equal([2, 3], plan.Slides.Select(slide => slide.ItemCount));
        Assert.Equal([0, 2], plan.Slides.Select(slide => slide.ItemOffset));
    }

    [Fact]
    public void Table_SplitsRowsAndRepeatsHeader_TotalOnlyOnLastSlide()
    {
        var rows = string.Join(",", Enumerable.Range(1, 40).Select(i => $"""["Region {i}","{i * 10}","{i}%"]""")) + ",[\"Total\",\"8200\",\"100%\"]";
        var plan = PlanJson($$$"""{"kind":"table","title":"Revenue","table":{"header":["Region","Revenue","Share"],"rows":[{{{rows}}}],"total_row":true}}""");
        Assert.True(plan.Slides.Count >= 2);
        Assert.All(plan.Slides, slide => Assert.Equal(["Region", "Revenue", "Share"], slide.Elements.Single(element => element.Key == "table").Table!.Header));
        Assert.All(plan.Slides.SkipLast(1), slide => Assert.False(slide.Elements.Single(element => element.Key == "table").Table!.TotalRow));
        Assert.True(plan.Slides[^1].Elements.Single(element => element.Key == "table").Table!.TotalRow);
        Assert.Equal(41, plan.Slides.Sum(slide => slide.ItemCount));
        Assert.Equal("Revenue (cont.)", plan.Slides[1].Title);
    }

    [Fact]
    public void Table_DetectsNumericColumnsForRightAlignment()
    {
        var plan = PlanJson("""{"kind":"table","title":"T","table":{"header":["Region","Revenue","Δ","Note"],"rows":[["EMEA","$3.1M","+16%","ok"],["APAC","2 400","−4 п.п.","—"]]}}""");
        Assert.Equal(["left", "right", "right", "left"], plan.Slides[0].Elements.Single(element => element.Key == "table").Table!.Align);
    }

    [Theory]
    [InlineData("1 234,5", true)]
    [InlineData("$4.2M", true)]
    [InlineData("(12.5)", true)]
    [InlineData("+3 pts", true)]
    [InlineData("12 млн руб.", true)]
    [InlineData("Q3", false)]
    [InlineData("EMEA", false)]
    [InlineData("2025-03-01", false)]
    public void LooksNumeric(string cell, bool expected) => Assert.Equal(expected, CompositionPlanner.LooksNumeric(cell));

    [Fact]
    public void Cards_ShareOneComponentIdentityAndGroupPerCard()
    {
        var plan = PlanJson(CompositionKinds.Find("cards")!.Example);
        var cards = plan.Slides[0].Elements.Where(element => element.Group?.StartsWith("card-", StringComparison.Ordinal) == true).GroupBy(element => element.Group).ToList();
        Assert.Equal(3, cards.Count);
        Assert.All(cards.SelectMany(group => group).Where(element => element.Role != "badge"), element => Assert.Equal("card@1", element.Component));
    }

    [Fact]
    public void TemplateTitle_IsRespected()
    {
        var parsed = CompositionParser.Parse(CompositionKinds.Find("cards")!.Example).Spec!;
        var template = new Box(60, 20, 900, 90);
        var plan = CompositionPlanner.Plan(parsed, Profile, 960, 540, new PlanContext { AppId = "x", TemplateTitle = template });
        Assert.Equal(template, plan.Slides[0].Elements.Single(element => element.Key == "title").Box);
        Assert.True(plan.Slides[0].Elements.Where(element => element.Key != "title").All(element => element.Box.Top >= template.Bottom));
    }

    [Fact]
    public void Hierarchy_TagsNodesAndConnectsEveryChild()
    {
        var plan = PlanJson(CompositionKinds.Find("hierarchy")!.Example);
        var elements = plan.Slides[0].Elements;
        var nodes = elements.Where(element => element.Role.StartsWith("node", StringComparison.Ordinal)).ToList();
        var edges = elements.Where(element => element.Type == "connector").ToList();
        Assert.Equal(5, nodes.Count);
        Assert.Equal(4, edges.Count);
        Assert.All(nodes, node => Assert.True(node.Tags!.ContainsKey(DeckRoles.NodeTag)));
        Assert.All(edges, edge => Assert.Contains(nodes, node => node.Key == edge.FromKey));
    }

    [Fact]
    public void Hierarchy_TooWide_IsAnErrorNotAnOverlap()
    {
        var children = string.Join(",", Enumerable.Range(1, 30).Select(i => $$"""{"label":"Unit {{i}}"}"""));
        Assert.Throws<ArgumentException>(() => PlanJson($$$"""{"kind":"hierarchy","title":"T","tree":{"label":"Root","children":[{{{children}}}]}}""", 720, 540));
    }

    [Fact]
    public void Kpis_ArrowFollowsTrendAndColorFollowsSentiment()
    {
        var plan = PlanJson(CompositionKinds.Find("kpis")!.Example);
        var elements = plan.Slides[0].Elements;
        var churnArrow = elements.Single(element => element.Key == "kpi-3.arrow");
        var churnDelta = elements.Single(element => element.Key == "kpi-3.delta");
        Assert.Equal(0f, churnArrow.Rotation);
        Assert.Equal(Profile.Color("negative"), churnArrow.Fill);
        Assert.Equal(Profile.Color("negative"), churnDelta.Paragraphs![0].Color);
    }

    [Fact]
    public void Source_GetsPrefixOnlyWhenItHasNone()
    {
        var withPrefix = PlanJson("""{"kind":"cards","title":"T","cards":[{"heading":"a"}],"source":"Источник: Росстат"}""");
        var without = PlanJson("""{"kind":"cards","title":"T","cards":[{"heading":"a"}],"source":"Rosstat"}""");
        Assert.Equal("Источник: Росстат", withPrefix.Slides[0].Elements.Single(element => element.Key == "source").Paragraphs![0].Text);
        Assert.Equal("Source: Rosstat", without.Slides[0].Elements.Single(element => element.Key == "source").Paragraphs![0].Text);
    }

    private static string Describe(CompositionPlan plan) => string.Join("\n", plan.Slides.SelectMany(slide => slide.Elements.Select(element =>
        $"{slide.Sequence}|{element.Key}|{element.Box}|{string.Join('/', (element.Paragraphs ?? []).Select(paragraph => $"{paragraph.Size}:{paragraph.Text}"))}")));

    /// <summary>Converts planned elements to snapshot objects with estimated text bounds, for rule checks.</summary>
    internal static (DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects) Snapshot(PlannedSlide slide)
    {
        int id = 1;
        var objects = new List<DeckObjectInfo>();
        foreach (var element in slide.Elements.Where(element => element.Type is not "connector"))
        {
            id++;
            bool hasText = element.Paragraphs is { Count: > 0 };
            objects.Add(new DeckObjectInfo
            {
                SlideIndex = 1,
                SlideId = 256,
                ShapeIndex = id,
                ShapeId = id,
                Name = element.Key,
                Kind = element.Type switch { "text" => "text-box", "shape" => "auto-shape", "line" => "line", "table" => "table", "chart" => "chart", "picture" => "picture", _ => "other" },
                Role = element.Role,
                Left = element.Box.Left,
                Top = element.Box.Top,
                Width = element.Box.Width,
                Height = element.Box.Height,
                Rotation = element.Rotation,
                ZOrder = id,
                Visible = true,
                HasText = hasText,
                Text = hasText ? string.Join("\n", element.Paragraphs!.Select(paragraph => paragraph.Text)) : null,
                TableRows = element.Table is null ? null : element.Table.Rows.Count + 1,
            });
        }
        return (new DeckSlideInfo { SlideIndex = 1, SlideId = 256, Title = slide.Title, Hidden = false, ShapeCount = objects.Count }, objects);
    }
}
