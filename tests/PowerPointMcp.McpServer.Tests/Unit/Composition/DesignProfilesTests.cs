// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Composition;

/// <summary>Design profile validation, inheritance, and token resolution (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Design")]
public sealed class DesignProfilesTests
{
    [Fact]
    public void BuiltIns_AllResolveWithoutErrors()
    {
        foreach (var name in DesignProfiles.BuiltInNames)
        {
            var result = DesignProfiles.Resolve(DesignProfiles.GetBuiltIn(name)!);
            Assert.True(result.IsValid, $"{name}: {string.Join("; ", result.Errors)}");
            Assert.All(DesignProfiles.KnownComponents, component => Assert.True(result.Resolved!.Components.ContainsKey(component), $"{name} lacks {component}"));
        }
    }

    [Fact]
    public void Inheritance_OverridesOnlyGivenValues()
    {
        var profile = DesignProfiles.Parse("""
            {
              "schema": "pptmcp.design-profile/1",
              "name": "acme",
              "base": "corporate-blue",
              "colors": { "accent": "#FF6600", "brand": "accent" },
              "fonts": { "heading": "Arial" },
              "type_scale": { "body": 18 },
              "table": { "header_fill": "brand" },
              "components": { "card": { "radius": 0 } }
            }
            """);
        var resolved = DesignProfiles.Resolve(profile).Resolved!;
        Assert.Equal("#FF6600", resolved.Color("brand"));
        Assert.Equal("#1F3864", resolved.Color("primary")); // from corporate-blue
        Assert.Equal("#2E7D32", resolved.Color("positive")); // from default
        Assert.Equal("Arial", resolved.HeadingFont);
        Assert.Equal("Segoe UI", resolved.BodyFont);
        Assert.Equal(18f, resolved.Size("body"));
        Assert.Equal("#FF6600", resolved.Table.HeaderFill);
        Assert.Equal(0f, resolved.Component("card").Radius);
        Assert.Equal(14f, resolved.Component("card").Padding); // kept from default
    }

    [Theory]
    [InlineData("""{"name":"x","colors":{"primary":"blue"}}""", "colors.primary")]
    [InlineData("""{"name":"x","colors":{"a":"b","b":"a"}}""", "colors.a")]
    [InlineData("""{"name":"x","type_scale":{"body":300}}""", "type_scale.body")]
    [InlineData("""{"name":"x","min_font_size":20}""", "type_scale.body")]
    [InlineData("""{"name":"x","base":"missing"}""", "base: profile 'missing'")]
    [InlineData("""{"name":"bad name"}""", "name:")]
    [InlineData("""{"name":"x","schema":"other/1"}""", "schema:")]
    [InlineData("""{"name":"x","chart":{"legend":"middle"}}""", "chart.legend")]
    [InlineData("""{"name":"x","table":{"header_fill":"nope"}}""", "table.header_fill")]
    [InlineData("""{"name":"x","components":{"card":{"heading_size":"huge"}}}""", "components.card.heading_size")]
    public void Invalid_ReportsPath(string json, string expected)
    {
        var result = DesignProfiles.Resolve(DesignProfiles.Parse(json));
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Contains(expected, StringComparison.Ordinal));
    }

    [Fact]
    public void UnknownProperties_AreWarnings()
    {
        var result = DesignProfiles.Resolve(DesignProfiles.Parse("""{"name":"x","shadows":true,"table":{"zebra":1},"components":{"widget":{"fill":"primary"}}}"""));
        Assert.True(result.IsValid, string.Join("; ", result.Errors));
        Assert.Contains(result.Warnings, warning => warning.StartsWith("shadows", StringComparison.Ordinal));
        Assert.Contains(result.Warnings, warning => warning.StartsWith("table.zebra", StringComparison.Ordinal));
        Assert.Contains(result.Warnings, warning => warning.StartsWith("components.widget", StringComparison.Ordinal));
    }

    [Fact]
    public void InheritanceCycle_IsAnError()
    {
        var profiles = new Dictionary<string, DesignProfile>
        {
            ["a"] = DesignProfiles.Parse("""{"name":"a","base":"b"}"""),
            ["b"] = DesignProfiles.Parse("""{"name":"b","base":"a"}"""),
        };
        var result = DesignProfiles.Resolve(profiles["a"], name => profiles.GetValueOrDefault(name));
        Assert.Contains(result.Errors, error => error.Contains("cycle", StringComparison.Ordinal));
    }

    [Fact]
    public void Json_RoundTripsAndAcceptsComments()
    {
        var original = DesignProfiles.GetBuiltIn("corporate-blue")!;
        var json = DesignProfiles.ToJson(original);
        Assert.Contains("\"type_scale\"", DesignProfiles.ToJson(DesignProfiles.Default()), StringComparison.Ordinal);
        var parsed = DesignProfiles.Parse("// comment\n" + json);
        Assert.Equal(original.Colors, parsed.Colors);
        Assert.Contains("\"min_font_size\"", DesignProfiles.ResolveBuiltIn("default").ToJson(), StringComparison.Ordinal);
    }

    [Fact]
    public void Resolved_ColorAndSizeLookups()
    {
        var profile = DesignProfiles.ResolveBuiltIn("default");
        Assert.Equal("#ABCDEF", profile.Color("#abcdef"));
        Assert.Equal(profile.Color("text"), profile.Color("unknown-token"));
        Assert.Equal(13.5f, profile.Size("13.5"));
        Assert.Equal(profile.Size("body"), profile.Size("not-a-role"));
        Assert.Equal("card@1", profile.Component("card").Identity);
        Assert.Equal("widget", profile.Component("widget").Name);
    }
}
