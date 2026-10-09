using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Notes;
using Sbroenne.PowerPointMcp.Core.Shape;
using Sbroenne.PowerPointMcp.Core.Table;
using Sbroenne.PowerPointMcp.Core.TextFrame;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Real integration tests for deck-wide find/replace and rich text editing (paragraphs, runs,
/// range replace/format, paragraph and frame format) against live PowerPoint COM.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "TextFrame")]
public class RichTextCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly DeckCommands _deck = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TableCommands _tables = new();
    private readonly NotesCommands _notes = new();
    private readonly TextFrameCommands _text = new();

    public RichTextCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    private int AddTextBox(string text, float top = 40)
    {
        var box = _shapes.AddTextBox(_fixture.Batch, 1, 40, top, 400, 60, text);
        Assert.True(box.Success, box.ErrorMessage);
        return box.ShapeIndex!.Value;
    }

    [Fact]
    public void FindText_SearchesTextBoxesTableCellsAndNotes()
    {
        _fixture.CreateFreshPresentation();
        AddTextBox("Выручка выросла");
        var table = _tables.AddTable(_fixture.Batch, 1, 2, 2, 40, 140, 400, 80);
        Assert.True(table.Success, table.ErrorMessage);
        Assert.True(_tables.SetCellText(_fixture.Batch, 1, table.ShapeIndex!.Value, 2, 1, "выручка, млн").Success);
        Assert.True(_notes.SetNotesText(_fixture.Batch, 1, "Сказать про ВЫРУЧКУ").Success);

        var withoutNotes = _deck.FindText(_fixture.Batch, "выручк");
        var withNotes = _deck.FindText(_fixture.Batch, "выручк", includeNotes: true);

        Assert.True(withoutNotes.Success, withoutNotes.ErrorMessage);
        Assert.Equal(2, withoutNotes.TotalCount);
        var cell = Assert.Single(withoutNotes.TextHits!, hit => hit.Row is not null);
        Assert.Equal((2, 1), (cell.Row!.Value, cell.Column!.Value));
        Assert.Equal(3, withNotes.TotalCount);
        Assert.Contains(withNotes.TextHits!, hit => hit.InNotes == true && hit.Matched == "ВЫРУЧК");
    }

    [Fact]
    public void ReplaceText_DryRunChangesNothingAndExpectedCountGuards()
    {
        _fixture.CreateFreshPresentation();
        int index = AddTextBox("Q3 revenue; Q3 margin");

        var preview = _deck.ReplaceText(_fixture.Batch, "Q3", "Q4");
        var guarded = _deck.ReplaceText(_fixture.Batch, "Q3", "Q4", dryRun: false, expectedCount: 3);

        Assert.True(preview.Success, preview.ErrorMessage);
        Assert.Equal(0, preview.ReplacementCount);
        Assert.Equal(2, preview.TotalCount);
        Assert.Contains("«Q4»", preview.TextHits![0].Context, StringComparison.Ordinal);
        Assert.False(guarded.Success);
        Assert.Equal("Q3 revenue; Q3 margin", _text.GetText(_fixture.Batch, 1, index).Text);
    }

    [Fact]
    public void ReplaceText_AppliesAndKeepsOtherRunsFormatted()
    {
        _fixture.CreateFreshPresentation();
        int index = AddTextBox("Old name is bold here");
        Assert.True(_text.FormatRange(_fixture.Batch, 1, index, match: "bold", bold: true).Success);

        var result = _deck.ReplaceText(_fixture.Batch, "Old name", "Новое имя", dryRun: false, expectedCount: 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(1, result.ReplacementCount);
        var paragraphs = _text.GetParagraphs(_fixture.Batch, 1, index);
        Assert.Equal("Новое имя is bold here", paragraphs.Text);
        var boldRun = Assert.Single(paragraphs.Paragraphs![0].Runs!, run => run.Bold == true);
        Assert.Equal("bold", boldRun.Text);
    }

    [Fact]
    public void ReplaceText_RegexWithGroupsAndSelectorScope()
    {
        _fixture.CreateFreshPresentation();
        int first = AddTextBox("Period 2024-03", top: 40);
        int second = AddTextBox("Period 2025-01", top: 140);
        var info = _deck.InspectObjects(_fixture.Batch, slideIndex: 1, selector: "text:\"*2025*\"");
        var shapeId = Assert.Single(info.Objects!).ShapeId;

        var result = _deck.ReplaceText(_fixture.Batch, @"(\d{4})-(\d{2})", "$2/$1", selector: $"id:{shapeId}", useRegex: true, dryRun: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("Period 2024-03", _text.GetText(_fixture.Batch, 1, first).Text);
        Assert.Equal("Period 01/2025", _text.GetText(_fixture.Batch, 1, second).Text);
    }

    [Fact]
    public void ReplaceRange_KeepsFormattingOfReplacedText()
    {
        _fixture.CreateFreshPresentation();
        int index = AddTextBox("Total: 12 units");
        Assert.True(_text.FormatRange(_fixture.Batch, 1, index, match: "12", italic: true, color: "#C00000").Success);

        var result = _text.ReplaceRange(_fixture.Batch, 1, index, "1 250", match: "12");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("Total: 1 250 units", result.Text);
        var runs = _text.GetParagraphs(_fixture.Batch, 1, index).Paragraphs![0].Runs!;
        var number = Assert.Single(runs, run => run.Italic == true);
        Assert.Equal(("1 250", "#C00000"), (number.Text, number.Color));
    }

    [Fact]
    public void ParagraphFormatAndFrameLayout_RoundTrip()
    {
        _fixture.CreateFreshPresentation();
        int index = AddTextBox("First\nSecond\nThird");
        Assert.True(_text.ReplaceRange(_fixture.Batch, 1, index, "First\nSecond\nThird").Success);

        var numbered = _text.SetParagraphFormat(_fixture.Batch, 1, index, paragraph: 2, paragraphCount: 2,
            alignment: "center", spaceAfter: 6, lineSpacing: 1.2f, bulletStyle: "numbered", numberStyle: "roman-upper-period", startAt: 2);
        var frame = _text.SetTextFrame(_fixture.Batch, 1, index, marginLeft: 12, verticalAnchor: "middle", columnCount: 2, columnSpacing: 18, rotation: 15);
        var read = _text.GetParagraphs(_fixture.Batch, 1, index);

        Assert.True(numbered.Success, numbered.ErrorMessage);
        Assert.True(frame.Success, frame.ErrorMessage);
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(3, read.Paragraphs!.Count);
        Assert.Equal("none", read.Paragraphs[0].Bullet);
        Assert.Equal(("numbered", "roman-upper-period", 2), (read.Paragraphs[1].Bullet, read.Paragraphs[1].NumberStyle, read.Paragraphs[1].StartAt!.Value));
        Assert.Equal(("center", 6f, 1.2f), (read.Paragraphs[2].Alignment, read.Paragraphs[2].SpaceAfter!.Value, read.Paragraphs[2].LineSpacing!.Value));
        Assert.Equal((12f, "middle", 2, 18f, 15f), (read.Frame!.Margins[0], read.Frame.VerticalAnchor, read.Frame.Columns, read.Frame.ColumnSpacing, read.Frame.Rotation));
    }

    [Fact]
    public void FormatRange_HyperlinkCharacterSpacingAndSuperscript()
    {
        _fixture.CreateFreshPresentation();
        int index = AddTextBox("See docs, m2 area");

        Assert.True(_text.FormatRange(_fixture.Batch, 1, index, match: "docs", hyperlink: "https://example.org/docs").Success);
        Assert.True(_text.FormatRange(_fixture.Batch, 1, index, match: "2", baselineOffset: 0.3f).Success);
        Assert.True(_text.FormatRange(_fixture.Batch, 1, index, match: "See", characterSpacing: 2).Success);
        var runs = _text.GetParagraphs(_fixture.Batch, 1, index).Paragraphs![0].Runs!;

        Assert.Contains(runs, run => run.Text == "docs" && run.Hyperlink == "https://example.org/docs");
        Assert.Contains(runs, run => run.Text == "2" && run.BaselineOffset == 0.3f);
        Assert.Contains(runs, run => run.Text == "See" && run.CharacterSpacing == 2f);
    }

    [Fact]
    public void InvalidSelections_ReturnActionableErrors()
    {
        _fixture.CreateFreshPresentation();
        int index = AddTextBox("abc");

        var missing = _text.FormatRange(_fixture.Batch, 1, index, match: "zzz", bold: true);
        var nothing = _text.FormatRange(_fixture.Batch, 1, index);
        var badRegex = _deck.FindText(_fixture.Batch, "(", useRegex: true);

        Assert.False(missing.Success);
        Assert.Contains("does not occur", missing.ErrorMessage, StringComparison.Ordinal);
        Assert.False(nothing.Success);
        Assert.False(badRegex.Success);
        Assert.Contains("regular expression", badRegex.ErrorMessage, StringComparison.Ordinal);
    }
}
