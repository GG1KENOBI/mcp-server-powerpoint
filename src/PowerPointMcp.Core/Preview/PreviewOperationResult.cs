using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Review;

namespace Sbroenne.PowerPointMcp.Core.Preview;

/// <summary>One slide on a contact sheet.</summary>
public sealed record ContactSheetSlide(int SlideIndex, int SlideId, string? Title, bool Hidden, int Column, int Row);

/// <summary>Result of a preview operation. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class PreviewOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Full path of the PNG rendered by PowerPoint (snapshot) or composed (contact sheet).</summary>
    public string? ImagePath { get; init; }

    /// <summary>Image width in pixels.</summary>
    public int? PixelWidth { get; init; }

    /// <summary>Image height in pixels.</summary>
    public int? PixelHeight { get; init; }

    /// <summary>Slide width in points.</summary>
    public float? SlideWidth { get; init; }

    /// <summary>Slide height in points.</summary>
    public float? SlideHeight { get; init; }

    /// <summary>Pixels per point in the image.</summary>
    public float? Scale { get; init; }

    /// <summary>1-based slide position.</summary>
    public int? SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID.</summary>
    public int? SlideId { get; init; }

    /// <summary>Slide title.</summary>
    public string? Title { get; init; }

    /// <summary>Slide fingerprint the image corresponds to.</summary>
    public string? Fingerprint { get; init; }

    /// <summary>Whether the image came from the preview cache.</summary>
    public bool? Cached { get; init; }

    /// <summary>Objects on the slide (compact).</summary>
    public IReadOnlyList<DeckObjectInfo>? Objects { get; init; }

    /// <summary>Validation findings for the slide.</summary>
    public IReadOnlyList<DeckFinding>? Findings { get; init; }

    /// <summary>Text wireframe of the slide for models without image input.</summary>
    public string? Sketch { get; init; }

    /// <summary>Slides on a contact sheet.</summary>
    public IReadOnlyList<ContactSheetSlide>? Slides { get; init; }

    /// <summary>Image comparison, for compare-images.</summary>
    public ImageDiffResult? Difference { get; init; }

    /// <summary>Changed area in slide points (left, top, width, height), for compare-images.</summary>
    public IReadOnlyList<float>? ChangedArea { get; init; }

    /// <summary>Files removed, for clear-cache.</summary>
    public int? Removed { get; init; }

    /// <summary>Non-fatal notes.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
