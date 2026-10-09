namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>One find-text or replace-text match, with where it is and the text around it.</summary>
public sealed class DeckTextHit
{
    /// <summary>1-based position of the slide.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID (stable across reordering).</summary>
    public required int SlideId { get; init; }

    /// <summary>Shape.Id of the shape holding the text (the notes body placeholder for notes).</summary>
    public required int ShapeId { get; init; }

    /// <summary>Selection Pane name of the shape.</summary>
    public required string ShapeName { get; init; }

    /// <summary>Persistent PPTMCP_ID of the shape, when it has one.</summary>
    public string? AppId { get; init; }

    /// <summary>1-based table row, for matches inside a table cell.</summary>
    public int? Row { get; init; }

    /// <summary>1-based table column, for matches inside a table cell.</summary>
    public int? Column { get; init; }

    /// <summary>True when the match is in the slide's speaker notes.</summary>
    public bool? InNotes { get; init; }

    /// <summary>1-based PowerPoint character position of the match in its text frame.</summary>
    public required int Start { get; init; }

    /// <summary>Length of the match in characters.</summary>
    public required int Length { get; init; }

    /// <summary>The matched text.</summary>
    public required string Matched { get; init; }

    /// <summary>The text that replaces (or would replace) the match.</summary>
    public string? Replacement { get; init; }

    /// <summary>Surrounding text with the match (or its replacement) marked «like this».</summary>
    public required string Context { get; init; }
}
