using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <summary>
/// Local images: a searchable catalog of image folders, and picture operations in the deck —
/// inspection (resolution, distortion, alt text, links), placing and replacing pictures without
/// distortion, alt text and attribution, embedding linked pictures, and repairing broken links.
/// Only local files are used; nothing is downloaded or generated.
/// </summary>
[ServiceCategory("asset", "Asset")]
[McpTool("asset", Title = "Images and Assets", Destructive = true, Category = "content",
    Description = "Local image assets. Catalog: scan-folder, search (words, orientation, min width, tag), tag-asset, duplicates. Deck pictures: inspect (effective PPI, distortion, crop, alt text, broken or linked files), place (contain or cover with a focal point, never stretched), replace (new image in the same frame, z-order, name, and tags), set-alt-text (alt text, decorative, attribution), embed-linked, fix-links. Never downloads or generates images.")]
[McpReadOnlyActions("search", "duplicates", "inspect")]
public interface IAssetCommands
{
    /// <summary>Adds or refreshes every image (PNG, JPEG, GIF, BMP, TIFF, WebP, SVG, EMF, WMF) under a local folder in the asset catalog: pixel size, orientation, and content hash.</summary>
    /// <param name="folder">Full path of a local folder.</param>
    /// <param name="recursive">Include subfolders (default true).</param>
    AssetOperationResult ScanFolder(IPresentationBatch batch, string folder, bool recursive = true);

    /// <summary>Searches the catalog by words in file names, tags, descriptions, and folders; filter by orientation, minimum width, or tag.</summary>
    /// <param name="query">Words to look for, e.g. "team office" (empty lists everything that passes the filters).</param>
    /// <param name="orientation">landscape, portrait, or square.</param>
    /// <param name="minWidth">Smallest pixel width.</param>
    /// <param name="tag">Only assets with this tag.</param>
    /// <param name="limit">Maximum results (1-200, default 20).</param>
    AssetOperationResult Search(IPresentationBatch batch, string? query = null, string? orientation = null, int minWidth = 0, string? tag = null, int limit = 20);

    /// <summary>Sets tags, a description (also the default alt text when placed), and attribution for an image; adds it to the catalog if needed.</summary>
    /// <param name="path">Full path of the image file.</param>
    /// <param name="tags">Comma-separated tags to add; prefix a tag with - to remove it.</param>
    /// <param name="description">What the image shows.</param>
    /// <param name="attribution">Credit or license text, e.g. "Photo: Jane Doe, CC BY 4.0".</param>
    AssetOperationResult TagAsset(IPresentationBatch batch, string path, string? tags = null, string? description = null, string? attribution = null);

    /// <summary>Lists catalog files with identical content, and deck pictures placed from the same source file.</summary>
    AssetOperationResult Duplicates(IPresentationBatch batch);

    /// <summary>
    /// Lists the deck's pictures with effective resolution (PPI), crop, distortion, alt text,
    /// decorative flag, attribution, and linked-file status, flagging issues.
    /// </summary>
    /// <param name="selector">Optional filter, e.g. slide:3 or name:"Logo".</param>
    AssetOperationResult Inspect(IPresentationBatch batch, string? selector = null);

    /// <summary>
    /// Places a local image on a slide, embedded, inside a frame (default: the content area
    /// below the title). fit=contain shows the whole image; fit=cover fills the frame and crops
    /// around the focal point. Never stretches. Alt text defaults to the catalog description.
    /// </summary>
    /// <param name="slideIndex">1-based slide position.</param>
    /// <param name="left">Frame left in points (pass all four of left, top, width, height, or none).</param>
    /// <param name="top">Frame top in points.</param>
    /// <param name="width">Frame width in points.</param>
    /// <param name="height">Frame height in points.</param>
    /// <param name="fit">contain (default for place) or cover (default for replace).</param>
    /// <param name="focalX">Horizontal focal point for cover, 0 (left) to 1 (right), default 0.5.</param>
    /// <param name="focalY">Vertical focal point for cover, 0 (top) to 1 (bottom), default 0.5.</param>
    /// <param name="altText">Alternative text for screen readers.</param>
    /// <param name="decorative">Mark as decorative (no alt text needed).</param>
    /// <param name="name">Shape name.</param>
    AssetOperationResult Place(
        IPresentationBatch batch,
        int slideIndex,
        string path,
        float? left = null,
        float? top = null,
        float? width = null,
        float? height = null,
        string fit = "contain",
        float focalX = 0.5f,
        float focalY = 0.5f,
        string? altText = null,
        string? attribution = null,
        bool decorative = false,
        string? name = null);

    /// <summary>
    /// Replaces a picture (or picture placeholder) with another local image in the same frame,
    /// keeping its position, size, z-order, name, tags, and alt text (unless alt_text is given).
    /// Animations on the old picture are not carried over (reported).
    /// </summary>
    /// <param name="target">Selector for exactly one picture, e.g. "slide:2 name:\"Hero\"" or "appid:hero".</param>
    AssetOperationResult Replace(
        IPresentationBatch batch,
        string target,
        string path,
        string fit = "cover",
        float focalX = 0.5f,
        float focalY = 0.5f,
        string? altText = null);

    /// <summary>Sets alt text, the decorative flag, and/or attribution on one object.</summary>
    AssetOperationResult SetAltText(IPresentationBatch batch, string target, string? altText = null, bool? decorative = null, string? attribution = null);

    /// <summary>Replaces linked pictures whose files exist with embedded copies (same frame and crop) so the deck no longer depends on the files. Dry run by default.</summary>
    /// <param name="dryRun">Only list what would change (default true).</param>
    AssetOperationResult EmbedLinked(IPresentationBatch batch, string? selector = null, bool dryRun = true);

    /// <summary>Finds files of broken links by file name under a local folder and relinks those with exactly one match. Dry run by default.</summary>
    /// <param name="searchFolder">Local folder to search (recursively) for the missing files.</param>
    AssetOperationResult FixLinks(IPresentationBatch batch, string searchFolder, string? selector = null, bool dryRun = true);
}
