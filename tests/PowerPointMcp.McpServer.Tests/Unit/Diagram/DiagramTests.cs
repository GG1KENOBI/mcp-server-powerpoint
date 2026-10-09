// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Diagram;
using Sbroenne.PowerPointMcp.Core.Review;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Diagram;

/// <summary>Diagram parsing, ranking, and layouts (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Diagram")]
public sealed class DiagramTests
{
    private static readonly ResolvedProfile Profile = DesignProfiles.ResolveBuiltIn("default");
    private static readonly Box Area = new(48, 116, 912, 500);

    public static TheoryData<string> Types() => new(DiagramParser.Types);

    private static DiagramSpec Parse(string json)
    {
        var (spec, errors, _) = DiagramParser.Parse(json);
        Assert.True(spec is not null, string.Join("; ", errors));
        return spec!;
    }

    [Theory]
    [MemberData(nameof(Types))]
    public void Examples_ParseAndLayOutInsideTheAreaWithoutTextOverlap(string type)
    {
        var spec = Parse(DiagramParser.Examples[type]);
        var plan = DiagramLayout.Layout(spec, Area, Profile);

        Assert.Equal(spec.Nodes.Count, plan.NodeKeys.Count);
        Assert.Equal(plan.Elements.Count, plan.Elements.Select(element => element.Key).Distinct().Count());
        foreach (var element in plan.Elements.Where(element => element.Type != "connector" && Math.Abs(element.Rotation) < 1))
            Assert.True(Area.Contains(element.Box, 1f), $"{type}: {element.Key} {element.Box} outside {Area}");
        Assert.All(plan.Elements.Where(element => element.Type == "connector"), connector =>
        {
            Assert.Contains(plan.Elements, element => element.Key == connector.FromKey);
            Assert.Contains(plan.Elements, element => element.Key == connector.ToKey);
        });

        // Nodes never overlap each other.
        var nodes = plan.Elements.Where(element => element.Role == "diagram-node").ToList();
        for (int i = 0; i < nodes.Count; i++)
        {
            for (int j = i + 1; j < nodes.Count; j++)
            {
                var overlap = nodes[i].Box.Intersect(nodes[j].Box);
                Assert.False(overlap.Width > 1 && overlap.Height > 1, $"{type}: {nodes[i].Key} overlaps {nodes[j].Key}");
            }
        }
    }

    [Fact]
    public void JsonRoundTrip_IsStable()
    {
        foreach (var json in DiagramParser.Examples.Values)
        {
            var spec = Parse(json);
            var again = Parse(DiagramParser.ToJson(spec));
            Assert.Equal(DiagramParser.ToJson(spec), DiagramParser.ToJson(again));
            Assert.Equal(spec.Nodes, again.Nodes);
            Assert.Equal(spec.Edges, again.Edges);
        }
    }

    [Theory]
    [InlineData("""{"type":"flowchart","id":"x","nodes":[]}""", "$.nodes: at least one node")]
    [InlineData("""{"type":"flowchart","id":"x","nodes":[{"id":"a","label":"A"},{"id":"a","label":"B"}]}""", "id 'a' is used 2 times")]
    [InlineData("""{"type":"flowchart","id":"x","nodes":[{"id":"a","label":"A"}],"edges":[{"from":"a","to":"b"}]}""", "$.edges[0].to: no node has id 'b'")]
    [InlineData("""{"type":"flowchart","id":"x","nodes":[{"id":"a","label":"A"}],"edges":[{"from":"a","to":"a"}]}""", "to itself")]
    [InlineData("""{"type":"swimlane","id":"x","lanes":["A"],"nodes":[{"id":"a","label":"A","lane":"B"}]}""", "lane: 'B' is not one of the lanes")]
    [InlineData("""{"type":"matrix","id":"x","nodes":[{"id":"a","label":"A"}]}""", "need a quadrant")]
    [InlineData("""{"type":"hub-spoke","id":"x","nodes":[{"id":"a","label":"A"}]}""", "$.hub: required")]
    [InlineData("""{"type":"tree","id":"x","nodes":[{"id":"a","label":"A"}]}""", "$.type")]
    [InlineData("""{"type":"flowchart","nodes":[{"id":"a","label":"A"}]}""", "$.id: required")]
    [InlineData("""{"type":"flowchart","id":"x","nodes":[{"id":"a","label":"A","shape":"star"}]}""", "$.nodes[0].shape")]
    public void Invalid_ReportsPaths(string json, string expected)
    {
        var (spec, errors, _) = DiagramParser.Parse(json);
        Assert.Null(spec);
        Assert.Contains(errors, error => error.Contains(expected, StringComparison.Ordinal));
    }

    [Fact]
    public void Rank_IsLongestPathAndFindsBackEdges()
    {
        var (ranks, back) = DiagramLayout.Rank(["a", "b", "c", "d"], [("a", "b"), ("b", "c"), ("a", "c"), ("c", "d"), ("d", "b")]);
        Assert.Equal(0, ranks["a"]);
        Assert.Equal(1, ranks["b"]);
        Assert.Equal(2, ranks["c"]);
        Assert.Equal(3, ranks["d"]);
        Assert.Contains(("d", "b"), back);
    }

    [Fact]
    public void CountCrossings_DetectsInversions()
    {
        var ranks = new Dictionary<string, int> { ["a"] = 0, ["b"] = 0, ["c"] = 1, ["d"] = 1 };
        var crossing = new Dictionary<string, float> { ["a"] = 0, ["b"] = 1, ["c"] = 1, ["d"] = 0 };
        var straight = new Dictionary<string, float> { ["a"] = 0, ["b"] = 1, ["c"] = 0, ["d"] = 1 };
        Assert.Equal(1, DiagramLayout.CountCrossings([("a", "c"), ("b", "d")], ranks, crossing));
        Assert.Equal(0, DiagramLayout.CountCrossings([("a", "c"), ("b", "d")], ranks, straight));
    }

    [Fact]
    public void Flowchart_DecisionBranchesShareARankAndLoopsGoAroundTheSide()
    {
        var spec = Parse("""{"type":"flowchart","id":"f","nodes":[{"id":"s","label":"Start"},{"id":"d","label":"OK?","shape":"decision"},{"id":"y","label":"Yes path"},{"id":"n","label":"No path"}],"edges":[{"from":"s","to":"d"},{"from":"d","to":"y","label":"yes"},{"from":"d","to":"n","label":"no"},{"from":"n","to":"s"}]}""");
        var plan = DiagramLayout.Layout(spec, Area, Profile);
        var y = plan.Elements.Single(element => element.Key == "node-y").Box;
        var n = plan.Elements.Single(element => element.Key == "node-n").Box;
        Assert.Equal(y.Top, n.Top, 1);
        var loop = plan.Elements.Single(element => element.Key == "edge-n-s");
        Assert.Equal((2, 2), (loop.FromSite, loop.ToSite));
        Assert.Contains(plan.Notes, note => note.Contains("loop-back", StringComparison.Ordinal));
        Assert.Equal(2, plan.Elements.Count(element => element.Role == "diagram-edge-label"));
    }

    [Fact]
    public void Flowchart_DirectionRight_UsesSideSites()
    {
        var spec = Parse("""{"type":"flowchart","id":"f","direction":"right","nodes":[{"id":"a","label":"A"},{"id":"b","label":"B"}],"edges":[{"from":"a","to":"b"}]}""");
        var plan = DiagramLayout.Layout(spec, Area, Profile);
        var edge = plan.Elements.Single(element => element.Type == "connector");
        Assert.Equal((4, 2), (edge.FromSite, edge.ToSite));
        Assert.True(plan.Elements.Single(element => element.Key == "node-b").Box.Left > plan.Elements.Single(element => element.Key == "node-a").Box.Right);
    }

    [Fact]
    public void Nodes_AreTaggedWithDiagramAndNodeIds()
    {
        var plan = DiagramLayout.Layout(Parse(DiagramParser.Examples["architecture"]), Area, Profile);
        var db = plan.Elements.Single(element => element.Key == "node-db");
        Assert.Equal("arch:db", db.Tags![DeckRoles.NodeTag]);
        Assert.Equal("arch", db.Tags["PPTMCP_DIAGRAM"]);
        Assert.Equal("cylinder", db.Geometry);
        Assert.Equal("diagram@1", db.Component);
    }

    [Fact]
    public void Swimlane_NodesSitInTheirLanes()
    {
        var spec = Parse(DiagramParser.Examples["swimlane"]);
        var plan = DiagramLayout.Layout(spec, Area, Profile);
        foreach (var node in spec.Nodes)
        {
            var lane = plan.Elements.Single(element => element.Role == "diagram-lane" && element.Tags!["PPTMCP_LANE"] == node.Lane);
            Assert.True(lane.Box.Contains(plan.Elements.Single(element => element.Key == $"node-{node.Id}").Box, 0.5f));
        }
    }

    [Fact]
    public void Layout_IsDeterministic()
    {
        foreach (var json in DiagramParser.Examples.Values)
        {
            string Describe() => string.Join("|", DiagramLayout.Layout(Parse(json), Area, Profile).Elements.Select(element => $"{element.Key}:{element.Box}"));
            Assert.Equal(Describe(), Describe());
        }
    }

    [Fact]
    public void DarkNodes_GetLightText()
    {
        var plan = DiagramLayout.Layout(Parse("""{"type":"flowchart","id":"f","nodes":[{"id":"a","label":"A","tone":"primary"},{"id":"b","label":"B","tone":"surface"}]}"""), Area, Profile);
        Assert.True(ColorContrast.Ratio(plan.Elements.Single(element => element.Key == "node-a").Paragraphs![0].Color, Profile.Color("primary")) >= 4.5);
        Assert.True(ColorContrast.Ratio(plan.Elements.Single(element => element.Key == "node-b").Paragraphs![0].Color, Profile.Color("surface")) >= 4.5);
    }
}
