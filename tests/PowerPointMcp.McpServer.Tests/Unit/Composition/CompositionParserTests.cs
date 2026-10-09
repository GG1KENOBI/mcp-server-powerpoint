// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Composition;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Composition;

/// <summary>Composition schema validation with JSONPath errors (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Composition")]
public sealed class CompositionParserTests
{
    public static TheoryData<string> Kinds() => new(CompositionKinds.All.Select(kind => kind.Name));

    [Theory]
    [MemberData(nameof(Kinds))]
    public void EveryKindExample_IsValid(string kind)
    {
        var result = CompositionParser.Parse(CompositionKinds.Find(kind)!.Example);
        Assert.Empty(result.Errors);
        Assert.Equal(kind, result.Spec!.Kind);
        Assert.DoesNotContain(result.Warnings, warning => warning.Contains("unknown property", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("not json", "$: invalid JSON")]
    [InlineData("[]", "must be a JSON object")]
    [InlineData("{}", "$.kind: required")]
    [InlineData("""{"kind":"slide"}""", "unknown kind 'slide'")]
    [InlineData("""{"kind":"cards","title":"T"}""", "$.cards: required for kind cards")]
    [InlineData("""{"kind":"cards","title":"T","cards":[{"body":"x"}]}""", "$.cards[0].heading: required")]
    [InlineData("""{"kind":"cards","title":42,"cards":[{"heading":"x"}]}""", "$.title: must be a string (got number)")]
    [InlineData("""{"kind":"comparison","title":"T","columns":[{"heading":"A","points":["x"]}]}""", "exactly 2 columns")]
    [InlineData("""{"kind":"table","title":"T","table":{"header":["a","b"],"rows":[["1"]]}}""", "$.table.rows[0]: has 1 cells but the header has 2 columns")]
    [InlineData("""{"kind":"table","title":"T","table":{"header":["a"],"rows":[["1"]],"highlight":[3]}}""", "row 3 is outside 1-1")]
    [InlineData("""{"kind":"chart","title":"T","chart":{"categories":["a","b"],"series":[{"name":"s","values":[1]}]}}""", "has 1 values but there are 2 categories")]
    [InlineData("""{"kind":"chart","title":"T","chart":{"categories":["a"],"series":[{"name":"s","values":["12%"]}]}}""", "text such as \"12%\" is not converted")]
    [InlineData("""{"kind":"chart","title":"T","chart":{"type":"pie","categories":["a"],"series":[{"name":"s","values":[1]},{"name":"t","values":[2]}]}}""", "exactly one series")]
    [InlineData("""{"kind":"chart","title":"T","chart":{"type":"radar","categories":["a"],"series":[{"name":"s","values":[1]}]}}""", "$.chart.type: 'radar'")]
    [InlineData("""{"kind":"kpis","title":"T","kpis":[{"value":"1","label":"x","trend":"sideways"}]}""", "$.kpis[0].trend")]
    [InlineData("""{"kind":"timeline","title":"T","milestones":[{"date":"Jan","label":"x"}]}""", "needs 2-12 items")]
    [InlineData("""{"kind":"executive-summary","title":"T","points":["a"],"fit":{"policy":["report","shrink"]}}""", "cannot be combined")]
    [InlineData("""{"kind":"executive-summary","title":"T","points":["a"],"insert_at":0}""", "$.insert_at: 0 is outside")]
    [InlineData("""{"kind":"executive-summary","title":"T","points":["a"],"id":"has space"}""", "$.id:")]
    [InlineData("""{"kind":"executive-summary","title":"T","points":[""]}""", "$.points[0]: empty point")]
    [InlineData("""{"kind":"image-text","title":"T","image":{"path":"a.png","fit":"stretch"}}""", "$.image.fit")]
    [InlineData("""{"kind":"title","title":"   "}""", "$.title: required for kind title")]
    public void Errors_HaveJsonPaths(string json, string expected)
    {
        var result = CompositionParser.Parse(json);
        Assert.Null(result.Spec);
        Assert.Contains(result.Errors, error => error.Contains(expected, StringComparison.Ordinal));
    }

    [Fact]
    public void UnusedAndUnknownFields_AreWarnings()
    {
        var result = CompositionParser.Parse("""{"kind":"quote","quote":{"text":"x","author":"y","year":1999},"cards":[{"heading":"h"}],"colour":"red"}""");
        Assert.Empty(result.Errors);
        Assert.Contains("$.colour: unknown property ignored.", result.Warnings);
        Assert.Contains("$.quote.year: unknown property ignored.", result.Warnings);
        Assert.Contains("$.cards: not used by kind quote; ignored.", result.Warnings);
    }

    [Fact]
    public void Text_IsKeptVerbatim()
    {
        const string text = "  Выручка: +12 %  \t(оценка)  ";
        var result = CompositionParser.Parse($$"""{"kind":"executive-summary","title":"Итоги","points":["{{text.Replace("\t", "\\t", StringComparison.Ordinal)}}"]}""");
        Assert.Equal(text, result.Spec!.Points![0].Text);
    }

    [Fact]
    public void TableNumbers_KeepTheirJsonSpelling()
    {
        var result = CompositionParser.Parse("""{"kind":"table","title":"T","table":{"header":["a","b"],"rows":[[1.50,"1,5"],[null,true]]}}""");
        Assert.Equal(["1.50", "1,5"], result.Spec!.Table!.Rows[0]);
        Assert.Equal(["", "true"], result.Spec.Table.Rows[1]);
    }

    [Fact]
    public void ChartNulls_AreMissingValues()
    {
        var result = CompositionParser.Parse("""{"kind":"chart","title":"T","chart":{"categories":["a","b"],"series":[{"name":"s","values":[1,null]}]}}""");
        Assert.True(double.IsNaN(result.Spec!.Chart!.Series[0].Values[1]));
    }

    [Fact]
    public void Defaults_AreApplied()
    {
        var spec = CompositionParser.Parse("""{"kind":"image-text","title":"T","image":{"path":"a.png"},"points":["x"]}""").Spec!;
        Assert.Equal(("cover", "left", 0.5f, 0.5f), (spec.Image!.Fit, spec.Image.Position, spec.Image.FocalX, spec.Image.FocalY));
        Assert.Equal(["shrink", "split"], spec.Fit.Policy);
        Assert.Contains("\"kind\":\"image-text\"", spec.OriginalJson, StringComparison.Ordinal);
    }

    [Fact]
    public void MissingAltText_IsAWarning()
    {
        var result = CompositionParser.Parse("""{"kind":"image-text","title":"T","image":{"path":"a.png"},"points":["x"]}""");
        Assert.Contains(result.Warnings, warning => warning.StartsWith("$.image.alt", StringComparison.Ordinal));
    }
}
