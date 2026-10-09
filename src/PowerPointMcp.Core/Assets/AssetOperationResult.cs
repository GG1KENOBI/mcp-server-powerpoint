namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <summary>A picture in the deck with its quality and accessibility checks.</summary>
public sealed class DeckPictureInfo
{
    /// <summary>1-based slide position.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID.</summary>
    public required int SlideId { get; init; }

    /// <summary>Shape.Id.</summary>
    public required int ShapeId { get; init; }

    /// <summary>Shape name.</summary>
    public required string Name { get; init; }

    /// <summary>Persistent PPTMCP_ID.</summary>
    public string? AppId { get; init; }

    /// <summary>Whether the picture is linked to a file instead of embedded.</summary>
    public bool Linked { get; init; }

    /// <summary>Linked file path.</summary>
    public string? LinkSource { get; init; }

    /// <summary>Whether the linked file exists on this computer.</summary>
    public bool? LinkExists { get; init; }

    /// <summary>File the picture was placed from (PPTMCP_IMG_SOURCE tag).</summary>
    public string? Source { get; init; }

    /// <summary>Source pixel size (width, height) when known.</summary>
    public IReadOnlyList<int>? PixelSize { get; init; }

    /// <summary>Visible size on the slide in points (width, height).</summary>
    public required IReadOnlyList<float> DisplaySize { get; init; }

    /// <summary>Effective resolution in pixels per inch of the visible picture, when the pixel size is known.</summary>
    public int? EffectivePpi { get; init; }

    /// <summary>Whether part of the picture is cropped away.</summary>
    public bool Cropped { get; init; }

    /// <summary>Aspect-ratio distortion in percent (0 = undistorted), when the pixel size is known.</summary>
    public float? DistortionPercent { get; init; }

    /// <summary>Alternative text.</summary>
    public string? AltText { get; init; }

    /// <summary>Marked decorative (no alt text needed).</summary>
    public bool Decorative { get; init; }

    /// <summary>Attribution or license text (PPTMCP_ATTRIBUTION tag).</summary>
    public string? Attribution { get; init; }

    /// <summary>missing-alt-text, low-resolution, soft-resolution, distorted, broken-link, linked, unknown-resolution.</summary>
    public IReadOnlyList<string>? Issues { get; init; }
}

/// <summary>Result of an asset command. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class AssetOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>The catalog file.</summary>
    public string? CatalogPath { get; init; }

    /// <summary>Folder scan counts.</summary>
    public AssetScanSummary? Scan { get; init; }

    /// <summary>Catalog entries (search results are ordered by relevance).</summary>
    public IReadOnlyList<AssetEntry>? Assets { get; init; }

    /// <summary>Groups of identical files (same content hash).</summary>
    public IReadOnlyList<IReadOnlyList<string>>? DuplicateGroups { get; init; }

    /// <summary>Pictures in the deck.</summary>
    public IReadOnlyList<DeckPictureInfo>? Pictures { get; init; }

    /// <summary>Total before the limit was applied.</summary>
    public int? TotalCount { get; init; }

    /// <summary>Slide of the placed or changed picture.</summary>
    public int? SlideIndex { get; init; }

    /// <summary>Shape.Id of the placed or changed picture.</summary>
    public int? ShapeId { get; init; }

    /// <summary>Name of the placed or changed picture.</summary>
    public string? ShapeName { get; init; }

    /// <summary>Planned or applied changes, one line each.</summary>
    public IReadOnlyList<string>? Changes { get; init; }

    /// <summary>Notes and follow-up advice.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
