namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>Read-only snapshot of one slide.</summary>
public sealed class DeckSlideInfo
{
    /// <summary>1-based position in the presentation.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID: unique in the presentation and stable across reordering and save.</summary>
    public required int SlideId { get; init; }

    /// <summary>Slide name.</summary>
    public string? Name { get; init; }

    /// <summary>Text of the title placeholder, when the slide has one.</summary>
    public string? Title { get; init; }

    /// <summary>Custom layout name.</summary>
    public string? LayoutName { get; init; }

    /// <summary>Whether the slide is hidden in slide shows.</summary>
    public required bool Hidden { get; init; }

    /// <summary>1-based section index, when the presentation has sections.</summary>
    public int? SectionIndex { get; init; }

    /// <summary>Section name, when the presentation has sections.</summary>
    public string? SectionName { get; init; }

    /// <summary>Number of top-level shapes.</summary>
    public required int ShapeCount { get; init; }

    /// <summary>Speaker notes (at most 300 characters unless detailed).</summary>
    public string? Notes { get; init; }

    /// <summary>Persistent application identifier (PPTMCP_ID slide tag).</summary>
    public string? AppId { get; init; }

    /// <summary>Slide string tags (names upper-case, as PowerPoint stores them).</summary>
    public IReadOnlyDictionary<string, string>? Tags { get; init; }

    /// <summary>Content fingerprint of the slide; changes whenever its objects or text change.</summary>
    public string? Fingerprint { get; init; }

    /// <summary>Solid background color as #RRGGBB when the slide background is a solid fill.</summary>
    public string? BackgroundColor { get; init; }
}
