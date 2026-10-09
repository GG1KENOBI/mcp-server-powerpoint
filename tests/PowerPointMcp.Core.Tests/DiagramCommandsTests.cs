using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Diagram;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Review;
using Sbroenne.PowerPointMcp.Core.Shape;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for diagrams against live PowerPoint COM.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Diagram")]
public class DiagramCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly DiagramCommands _diagrams = new();

    public DiagramCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    public static TheoryData<string> Types() => new(DiagramParser.Types);

    private void FreshWideDeck()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new PageSetupCommands().SetSize(_fixture.Batch, 960, 540).Success);
    }

    [Theory]
    [MemberData(nameof(Types))]
    public void Create_EveryType_GluesEveryConnector(string type)
    {
        FreshWideDeck();
        var result = _diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples[type], profile: "default");

        Assert.True(result.Success, result.ErrorMessage + " " + string.Join("; ", result.Errors ?? []));
        var inspect = _diagrams.Inspect(_fixture.Batch, result.DiagramId!);
        Assert.True(inspect.Success, inspect.ErrorMessage);
        Assert.All(inspect.Edges!, edge => Assert.True(edge.Attached, $"{edge.From}→{edge.To} is not glued"));
        Assert.Equal(result.Nodes!.Count, inspect.Nodes!.Count);

        var review = new ReviewCommands().Validate(_fixture.Batch, slideIndex: result.SlideIndex, rules: ["off-slide", "text-overflow"]);
        Assert.True(review.TotalFindings == 0, string.Join(" | ", review.Findings!.Select(finding => finding.Message)));
    }

    [Fact]
    public void MovingANode_KeepsConnectorsGlued()
    {
        FreshWideDeck();
        var result = _diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["flowchart"]);
        Assert.True(result.Success, result.ErrorMessage);
        var node = _deck.Find(_fixture.Batch, "appid:approval/node-cfo", requireSingle: true).Objects![0];
        Assert.True(new ShapeCommands().SetPosition(_fixture.Batch, result.SlideIndex!.Value, node.ShapeIndex!.Value, node.Left + 40, node.Top).Success);

        var inspect = _diagrams.Inspect(_fixture.Batch, "approval");
        Assert.All(inspect.Edges!, edge => Assert.True(edge.Attached));
    }

    private readonly DeckCommands _deck = new();

    [Fact]
    public void UpdateNode_ChangesTextInPlace_KeepsShapeId()
    {
        FreshWideDeck();
        var created = _diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["flowchart"]);
        var before = created.Nodes!.Single(node => node.NodeId == "cfo").ShapeId;

        var updated = _diagrams.UpdateNode(_fixture.Batch, "approval", "cfo", label: "Финансовый директор", detail: "до 3 дней");

        Assert.True(updated.Success, updated.ErrorMessage);
        Assert.Equal(before, updated.Nodes![0].ShapeId);
        var text = _deck.Find(_fixture.Batch, "appid:approval/node-cfo", requireSingle: true).Objects![0].Text;
        Assert.Contains("Финансовый директор", text, StringComparison.Ordinal);
        Assert.Contains("до 3 дней", text, StringComparison.Ordinal);
        Assert.Contains("Финансовый директор", _diagrams.Inspect(_fixture.Batch, "approval").Spec, StringComparison.Ordinal);
    }

    [Fact]
    public void AddAndRemoveNodes_RelayoutKeepsNodeIds()
    {
        FreshWideDeck();
        Assert.True(_diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["flowchart"]).Success);

        var added = _diagrams.AddNode(_fixture.Batch, "approval", "audit", "Audit log", connectFrom: "done");
        Assert.True(added.Success, added.ErrorMessage);
        Assert.Contains(added.Nodes!, node => node.NodeId == "audit");
        Assert.Contains(added.Edges!, edge => edge.From == "done" && edge.To == "audit");

        var removed = _diagrams.RemoveNode(_fixture.Batch, "approval", "auto");
        Assert.True(removed.Success, removed.ErrorMessage);
        Assert.DoesNotContain(removed.Nodes!, node => node.NodeId == "auto");
        Assert.DoesNotContain(removed.Edges!, edge => edge.From == "auto" || edge.To == "auto");
        Assert.Empty(_deck.Find(_fixture.Batch, "appid:approval/node-auto").Objects!);
    }

    [Fact]
    public void InvalidEdits_ChangeNothing()
    {
        FreshWideDeck();
        Assert.True(_diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["flowchart"]).Success);
        var revision = _deck.Fingerprint(_fixture.Batch).Revision;

        Assert.False(_diagrams.AddEdge(_fixture.Batch, "approval", "start", "missing").Success);
        Assert.False(_diagrams.AddNode(_fixture.Batch, "approval", "start", "Duplicate").Success);
        Assert.False(_diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["flowchart"]).Success);

        Assert.Equal(revision, _deck.Fingerprint(_fixture.Batch).Revision);
    }

    [Fact]
    public void List_FindsDiagrams()
    {
        FreshWideDeck();
        Assert.True(_diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["matrix"]).Success);
        Assert.True(_diagrams.Create(_fixture.Batch, spec: DiagramParser.Examples["hub-spoke"]).Success);
        var list = _diagrams.List(_fixture.Batch);
        Assert.Equal(["prio", "eco"], list.Diagrams!.Select(diagram => diagram.DiagramId));
    }
}
