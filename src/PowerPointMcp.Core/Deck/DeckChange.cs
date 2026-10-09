namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>One structural difference between two inspection snapshots.</summary>
public sealed class DeckChange
{
    /// <summary>slide-added, slide-removed, slide-moved, object-added, object-removed, or object-modified.</summary>
    public required string Change { get; init; }

    /// <summary>PowerPoint SlideID of the affected slide.</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based slide position after the change (before, for removals).</summary>
    public required int SlideIndex { get; init; }

    /// <summary>Shape.Id of the affected object, for object changes.</summary>
    public int? ShapeId { get; init; }

    /// <summary>Object name, for object changes.</summary>
    public string? Name { get; init; }

    /// <summary>Properties that differ, for object-modified (e.g. left, text, maxFontSize).</summary>
    public IReadOnlyList<string>? Fields { get; init; }
}
