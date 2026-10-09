using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Preview;

/// <summary>
/// Visual feedback rendered by PowerPoint itself: slide snapshots (PNG plus objects, findings,
/// fingerprint, and a text wireframe for text-only models), contact sheets, and image comparison.
/// MCP clients receive the PNG as image content unless include_image=false or
/// PPTMCP_PREVIEW_IMAGES=off. Geometry and findings are not visual understanding: look at the
/// image when it matters.
/// </summary>
[ServiceCategory("preview", "Preview")]
[McpTool("preview", Title = "Slide Preview", Destructive = true, Category = "content",
    Description = "Render a slide with PowerPoint to PNG and get the image (when the client supports images), its file path, the objects, validation findings, a text wireframe, and a fingerprint; build contact sheets of many slides; compare two renders. Set include_image=false for text-only models.")]
[McpReadOnlyActions("snapshot", "contact-sheet", "compare-images")]
public interface IPreviewCommands
{
    /// <summary>
    /// Renders one slide with PowerPoint (Slide.Export) and returns the PNG path, pixel size,
    /// slide size, scale, fingerprint, objects, findings, and a text wireframe. Reuses a cached
    /// render when the slide is unchanged (see refresh).
    /// </summary>
    /// <param name="slideIndex">1-based slide to render.</param>
    /// <param name="slideId">PowerPoint SlideID to render.</param>
    /// <param name="width">Image width in pixels (160-3840, default 1280).</param>
    /// <param name="outputPath">Full path of a .png to write (default: the preview cache folder).</param>
    /// <param name="overwrite">Allow replacing an existing output file (default false).</param>
    /// <param name="includeFindings">Validate the slide and include findings (default true).</param>
    /// <param name="includeSketch">Include the text wireframe (default true).</param>
    /// <param name="refresh">Render again even if a cached image matches (default false).</param>
    /// <param name="includeImage">Attach the PNG as MCP image content (default true; false for text-only models).</param>
    PreviewOperationResult Snapshot(
        IPresentationBatch batch,
        int? slideIndex = null,
        int? slideId = null,
        int width = 1280,
        string? outputPath = null,
        bool overwrite = false,
        bool includeFindings = true,
        bool includeSketch = true,
        bool refresh = false,
        bool includeImage = true);

    /// <summary>
    /// Renders thumbnails with PowerPoint and composes them into one PNG grid labelled with slide
    /// numbers and titles (hidden slides marked).
    /// </summary>
    /// <param name="startSlide">First slide (default 1).</param>
    /// <param name="maxSlides">Slides on the sheet (1-48, default 24).</param>
    /// <param name="columns">Columns (1-8, default 4).</param>
    /// <param name="thumbnailWidth">Thumbnail width in pixels (120-640, default 320).</param>
    PreviewOperationResult ContactSheet(
        IPresentationBatch batch,
        int startSlide = 1,
        int maxSlides = 24,
        int columns = 4,
        int thumbnailWidth = 320,
        string? outputPath = null,
        bool overwrite = false,
        bool includeImage = true);

    /// <summary>
    /// Compares two PNG renders of the same size (e.g. before and after an edit) and reports the
    /// share of changed pixels and the changed area in pixels and slide points.
    /// </summary>
    /// <param name="firstPath">First PNG.</param>
    /// <param name="secondPath">Second PNG.</param>
    /// <param name="threshold">Per-channel difference treated as unchanged (0-255, default 24).</param>
    PreviewOperationResult CompareImages(IPresentationBatch batch, string firstPath, string secondPath, int threshold = 24);

    /// <summary>Deletes cached previews older than ten minutes, or all with all=true.</summary>
    /// <param name="all">Delete every cached preview.</param>
    PreviewOperationResult ClearCache(IPresentationBatch batch, bool all = false);
}
