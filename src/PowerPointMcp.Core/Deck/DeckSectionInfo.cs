namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>One presentation section.</summary>
public sealed class DeckSectionInfo
{
    /// <summary>1-based section index.</summary>
    public required int SectionIndex { get; init; }

    /// <summary>Section name.</summary>
    public required string Name { get; init; }

    /// <summary>1-based position of the section's first slide (0 when the section is empty).</summary>
    public required int FirstSlide { get; init; }

    /// <summary>Number of slides in the section.</summary>
    public required int SlideCount { get; init; }
}
