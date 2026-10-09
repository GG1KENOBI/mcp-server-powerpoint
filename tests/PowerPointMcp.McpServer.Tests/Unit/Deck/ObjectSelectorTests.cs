// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Deck.DeckTestData;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Deck;

/// <summary>Selector grammar and evaluation over inspection snapshots (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Deck")]
public sealed class ObjectSelectorTests
{
    private static IReadOnlyList<DeckObjectInfo> Find(string selector) => ObjectSelector.Parse(selector).Find(SampleDeck());

    private static int[] Ids(string selector) => Find(selector).Select(item => item.SlideId * 1000 + item.ShapeId).ToArray();

    [Theory]
    [InlineData("kind:chart", new[] { 260004 })]
    [InlineData("KIND:Chart", new[] { 260004 })]
    [InlineData("role:title", new[] { 256002, 260002 })]
    [InlineData("role:title slide:2", new[] { 260002 })]
    [InlineData("slideid:256 role:card", new[] { 256011, 256012 })]
    [InlineData("title:agenda kind:auto-shape", new[] { 256011, 256012 })]
    [InlineData("title:\"Q3*\" kind:table", new[] { 260005 })]
    [InlineData("appid:q3-chart", new[] { 260004 })]
    [InlineData("slideappid:agenda id:2", new[] { 256002 })]
    [InlineData("section:Results kind:table", new[] { 260005 })]
    [InlineData("layout:\"Title and Content\" role:title", new[] { 256002 })]
    [InlineData("name:\"Card ?\"", new[] { 256011, 256012 })]
    [InlineData("text:grew", new[] { 256011 })]
    [InlineData("text:\"Revenue*\"", new[] { 256011 })]
    [InlineData("placeholder:title", new[] { 256002, 260002 })]
    [InlineData("group:Cards", new[] { 256011, 256012 })]
    [InlineData("group:cards", new[] { 256011, 256012 })]
    [InlineData("group:10 font<10", new[] { 256012 })]
    [InlineData("font>=40", new[] { 256002, 260002 })]
    [InlineData("hidden", new[] { 260005 })]
    [InlineData("kind:table visible", new int[0])]
    [InlineData("tag:SOURCE=q3.csv", new[] { 260005 })]
    [InlineData("tag:source=Q3*", new[] { 260005 })]
    [InlineData("tag:PPTMCP_ID=q3-chart", new[] { 260004 })]
    [InlineData("slidetag:OWNER=finance role:title", new[] { 260002 })]
    [InlineData("component:card", new[] { 256011, 256012 })]
    [InlineData("component:card@1 text:fell", new[] { 256012 })]
    [InlineData("right>900", new[] { 260005 })]
    [InlineData("left>=40 top<=120 width=400", new[] { 260004 })]
    [InlineData("index:5", new[] { 260005 })]
    public void Find_ReturnsExpectedObjects(string selector, int[] expected)
    {
        Assert.Equal(expected, Ids(selector));
    }

    [Theory]
    [InlineData("", "empty")]
    [InlineData("   ", "empty")]
    [InlineData("colour:red", "Unknown selector key 'colour'")]
    [InlineData("kind:", "has no value")]
    [InlineData("slide:two", "whole number")]
    [InlineData("id:-1", "whole number")]
    [InlineData("tag:SOURCE", "tag:NAME=VALUE")]
    [InlineData("text:\"open", "unterminated quote")]
    [InlineData("chart", "not key:value")]
    [InlineData("font<big", "not key:value")]
    public void Parse_RejectsInvalidSelectors(string selector, string expected)
    {
        var error = Assert.Throws<ArgumentException>(() => ObjectSelector.Parse(selector));
        Assert.Contains(expected, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void QuotedValues_KeepSpaces()
    {
        Assert.Equal([256002, 260002], Ids("name:\"Title 1\""));
    }

    [Fact]
    public void HasObjectConditions_DistinguishesSlideOnlySelectors()
    {
        Assert.False(ObjectSelector.Parse("slide:2 title:Q3").HasObjectConditions);
        Assert.True(ObjectSelector.Parse("slide:2 kind:chart").HasObjectConditions);
    }

    [Fact]
    public void RequireSingle_ReturnsTheOnlyMatch()
    {
        var match = ObjectSelector.RequireSingle("appid:q3-chart", Find("appid:q3-chart"), 7);
        Assert.Equal(4, match.ShapeId);
    }

    [Fact]
    public void RequireSingle_AmbiguousListsCandidatesAndChangesNothing()
    {
        var error = Assert.Throws<ArgumentException>(() => ObjectSelector.RequireSingle("role:card", Find("role:card"), 7));
        Assert.Contains("ambiguous", error.Message, StringComparison.Ordinal);
        Assert.Contains("matched 2 objects", error.Message, StringComparison.Ordinal);
        Assert.Contains("id:11", error.Message, StringComparison.Ordinal);
        Assert.Contains("id:12", error.Message, StringComparison.Ordinal);
        Assert.Contains("\"Revenue grew 12%\"", error.Message, StringComparison.Ordinal);
        Assert.Contains("Nothing was changed", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RequireSingle_NoMatchExplainsHowToInspect()
    {
        var error = Assert.Throws<ArgumentException>(() => ObjectSelector.RequireSingle("kind:video", Find("kind:video"), 7));
        Assert.Contains("matched none of the 7", error.Message, StringComparison.Ordinal);
        Assert.Contains("inspect-objects", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RequireSingle_CapsCandidateList()
    {
        var many = Enumerable.Range(1, 14).Select(id => Shape(1, 256, id, $"Box {id}")).ToList();
        var error = Assert.Throws<ArgumentException>(() => ObjectSelector.RequireSingle("kind:text-box", many, 14));
        Assert.Contains("(and 4 more)", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("id:11 ", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Matching_IsCultureInvariantForRussianText()
    {
        var deck = new List<(DeckSlideInfo, IReadOnlyList<DeckObjectInfo>)>
        {
            (Slide(1, 300, "Выручка за квартал"), [Shape(1, 300, 3, "Текст", text: "Рост ВЫРУЧКИ на 12%")]),
        };
        Assert.Single(ObjectSelector.Parse("title:выручка text:\"рост выручки\"").Find(deck));
        Assert.Single(ObjectSelector.Parse("name:текст").Find(deck));
    }

    [Theory]
    [InlineData("ppPlaceholderBody", "body")]
    [InlineData("ppPlaceholderCenterTitle", "centertitle")]
    [InlineData("custom", "custom")]
    [InlineData(null, null)]
    public void PlaceholderShortName_StripsPrefix(string? type, string? expected)
    {
        Assert.Equal(expected, ObjectSelector.PlaceholderShortName(type));
    }
}
