using System.Runtime.Versioning;
using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Review;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for slide composition against live PowerPoint COM.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Composition")]
[SupportedOSPlatform("windows")]
public class ComposeCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly ComposeCommands _compose = new();
    private readonly ReviewCommands _review = new();
    private readonly DeckCommands _deck = new();

    public ComposeCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    public static TheoryData<string> Kinds() => new(CompositionKinds.All.Select(kind => kind.Name));

    private void FreshWideDeck()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new PageSetupCommands().SetSize(_fixture.Batch, 960, 540).Success);
    }

    private static string ExampleFor(string kind)
    {
        var example = CompositionKinds.Find(kind)!.Example;
        if (kind != "image-text")
            return example;
        var path = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"photo-{Guid.NewGuid():N}.png");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using (var bitmap = new System.Drawing.Bitmap(1600, 900))
        {
            using var graphics = System.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(System.Drawing.Color.SteelBlue);
            bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }
        return example.Replace(@"C:\\Assets\\plant.jpg", JsonEncodedText.Encode(path).ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(Kinds))]
    public void Create_EveryKind_ProducesNativeObjectsWithoutLayoutErrors(string kind)
    {
        FreshWideDeck();

        var result = _compose.Create(_fixture.Batch, spec: ExampleFor(kind), profile: "default");

        Assert.True(result.Success, result.ErrorMessage + " " + string.Join("; ", result.Errors ?? []));
        var slide = Assert.Single(result.Slides!);
        Assert.NotEmpty(slide.Elements);
        Assert.All(slide.Elements, element => Assert.True(element.ShapeId > 0));
        Assert.Null(result.Overflow);

        var findings = _review.Validate(_fixture.Batch, slideIndex: slide.SlideIndex, rules: ["off-slide", "text-overflow", "text-overlap"]);
        Assert.True(findings.Success, findings.ErrorMessage);
        Assert.True(findings.TotalFindings == 0, string.Join(" | ", findings.Findings!.Select(finding => finding.Message)));
    }

    [Fact]
    public void Create_SemanticMapIdsAreAddressable()
    {
        FreshWideDeck();
        var result = _compose.Create(_fixture.Batch, spec: """{"kind":"cards","id":"prio","title":"Priorities","cards":[{"heading":"Growth","body":"Two markets"},{"heading":"People","body":"Hire 40"}]}""");
        Assert.True(result.Success, result.ErrorMessage);

        var body = _deck.Find(_fixture.Batch, "appid:prio/card-2.body", requireSingle: true);
        Assert.True(body.Success, body.ErrorMessage);
        Assert.Equal("Hire 40", body.Objects![0].Text);
        Assert.Single(_deck.Find(_fixture.Batch, "slideappid:prio role:title").Objects!);
        Assert.Equal(2, _deck.Find(_fixture.Batch, "slideappid:prio component:card role:card-heading").Objects!.Count);
    }

    [Fact]
    public void Create_LongSummary_SplitsByMeasurementAndNeverCutsText()
    {
        FreshWideDeck();
        var points = string.Join(",", Enumerable.Range(1, 16).Select(i =>
            $$"""{"text":"Пункт {{i}}: выручка выросла в регионе","detail":"Подробное пояснение номер {{i}}, которое занимает две строки текста на слайде шириной 960 точек."}"""));
        var spec = $$$"""{"kind":"executive-summary","id":"sum","title":"Итоги года","points":[{{{points}}}],"fit":{"policy":["split"]}}""";

        var result = _compose.Create(_fixture.Batch, spec: spec);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.Slides!.Count >= 2);
        Assert.Equal(["sum", "sum~2"], result.Slides.Take(2).Select(slide => slide.AppId));
        Assert.Null(result.Overflow);
        var allText = string.Join("\n", _deck.Find(_fixture.Batch, "slideappid:sum* role:point").Objects!.Select(item => item.Text));
        for (int i = 1; i <= 16; i++)
            Assert.Contains($"Подробное пояснение номер {i},", allText, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_DuplicateId_FailsUnlessReplace_ReplaceKeepsPosition()
    {
        FreshWideDeck();
        const string spec = """{"kind":"quote","id":"q","title":"Voice of the customer","quote":{"text":"Fast and simple.","author":"A. Customer"}}""";
        var first = _compose.Create(_fixture.Batch, spec: spec, insertAt: 1);
        Assert.True(first.Success, first.ErrorMessage);

        var duplicate = _compose.Create(_fixture.Batch, spec: spec);
        Assert.False(duplicate.Success);
        Assert.Contains("replace=true", duplicate.ErrorMessage, StringComparison.Ordinal);

        var replaced = _compose.Create(_fixture.Batch, spec: spec.Replace("Fast and simple.", "Faster than ever.", StringComparison.Ordinal), replace: true);
        Assert.True(replaced.Success, replaced.ErrorMessage);
        Assert.Equal(1, replaced.Slides![0].SlideIndex);
        Assert.Single(_deck.Find(_fixture.Batch, "slideappid:q role:quote").Objects!);
    }

    [Fact]
    public void GetSpec_RoundTripsTheComposition_AndUpdateTextKeepsFormatting()
    {
        FreshWideDeck();
        var spec = CompositionKinds.Find("kpis")!.Example.Replace("{\"kind\":\"kpis\"", "{\"kind\":\"kpis\",\"id\":\"q3\"", StringComparison.Ordinal);
        Assert.True(_compose.Create(_fixture.Batch, spec: spec).Success);

        var stored = _compose.GetSpec(_fixture.Batch, "q3");
        Assert.True(stored.Success, stored.ErrorMessage);
        Assert.Contains("\"Gross margin\"", stored.Spec, StringComparison.Ordinal);

        var before = _deck.Find(_fixture.Batch, "appid:q3/kpi-1.value", requireSingle: true).Objects![0];
        var update = _compose.UpdateText(_fixture.Batch, "q3/kpi-1.value", "$4.5M");
        Assert.True(update.Success, update.ErrorMessage);
        var after = _deck.Find(_fixture.Batch, "appid:q3/kpi-1.value", requireSingle: true).Objects![0];
        Assert.Equal("$4.5M", after.Text);
        Assert.Equal(before.MaxFontSize, after.MaxFontSize);
        Assert.NotNull(_compose.GetSpec(_fixture.Batch, "q3").Warnings);
    }

    [Fact]
    public void Plan_ChangesNothing()
    {
        FreshWideDeck();
        var revision = _deck.Fingerprint(_fixture.Batch).Revision;
        var plan = _compose.Plan(_fixture.Batch, spec: CompositionKinds.Find("timeline")!.Example);
        Assert.True(plan.Success, plan.ErrorMessage);
        Assert.NotEmpty(plan.Planned!);
        Assert.Equal(revision, _deck.Fingerprint(_fixture.Batch).Revision);
    }

    [Fact]
    public void Create_InvalidSpec_ReturnsAllErrorsAndCreatesNothing()
    {
        FreshWideDeck();
        var count = _deck.Summary(_fixture.Batch, includeTheme: false).SlideCount;
        var result = _compose.Create(_fixture.Batch, spec: """{"kind":"table","title":"T","table":{"header":["a","b"],"rows":[["1"],["1","2","3"]]}}""");
        Assert.False(result.Success);
        Assert.Equal(2, result.Errors!.Count);
        Assert.Equal(count, _deck.Summary(_fixture.Batch, includeTheme: false).SlideCount);
    }

    [Fact]
    public void Create_WithThemeProfile_UsesTheTemplatesTitlePlacement()
    {
        FreshWideDeck();
        var result = _compose.Create(_fixture.Batch, spec: CompositionKinds.Find("cards")!.Example);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.StartsWith("theme@", result.Profile, StringComparison.Ordinal);
    }
}
