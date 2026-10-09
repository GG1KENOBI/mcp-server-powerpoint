using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Template;

/// <summary>
/// Template workflows: inspect masters and layouts with their placeholders, add slides from a
/// layout by name, fill placeholders by role, import slides from another deck with a report, and
/// preview how a template would map the deck's layouts before applying it.
/// </summary>
[ServiceCategory("template", "Template")]
[McpTool("template", Title = "Templates and Layouts", Destructive = true, Category = "content",
    Description = "Work with corporate templates: list-layouts shows every master layout with its placeholders; add-slide creates a slide from a layout by name and fills placeholders by role (title, subtitle, body, body2, picture, notes); fill-placeholders fills an existing slide; import-slides copies slides from another local deck and reports layouts, missing fonts, and layout problems; preview-template compares a template with the deck before presentation apply-template.")]
[McpReadOnlyActions("list-layouts", "preview-template")]
public interface ITemplateCommands
{
    /// <summary>Lists masters and their layouts: name, slides using each, and placeholders (role, type, name, bounds) in reading order.</summary>
    /// <param name="masterIndex">Only this 1-based master (default all).</param>
    TemplateOperationResult ListLayouts(IPresentationBatch batch, int? masterIndex = null);

    /// <summary>
    /// Adds a slide using a layout found by name (exact, ignoring case, else a unique partial
    /// match) and optionally fills its placeholders. Content JSON: {"title": "...", "subtitle":
    /// "...", "body": ["point", {"text": "detail", "level": 2}], "body2": "...", "picture":
    /// "C:\\images\\photo.jpg", "notes": "..."}. Pictures are local files, embedded, never stretched.
    /// </summary>
    /// <param name="layout">Layout name, e.g. "Title and Content" or "Two Content".</param>
    /// <param name="position">1-based position of the new slide (default: after the last slide).</param>
    /// <param name="content">Placeholder content JSON keyed by role (title, subtitle, body, body2, ..., picture, notes).</param>
    TemplateOperationResult AddSlide(IPresentationBatch batch, string layout, int? position = null, string? content = null, int? masterIndex = null);

    /// <summary>Fills an existing slide's placeholders by role from content JSON (same format as add-slide).</summary>
    /// <param name="slideIndex">1-based slide position.</param>
    /// <param name="removeEmpty">Delete placeholders that remain empty (default false).</param>
    TemplateOperationResult FillPlaceholders(IPresentationBatch batch, int slideIndex, string content, bool removeEmpty = false);

    /// <summary>
    /// Imports slides from another local presentation (they take this deck's theme) and reports
    /// each slide's source and new layout, fonts that are not installed, and validation findings
    /// such as overflow or off-slide objects.
    /// </summary>
    /// <param name="sourcePath">Full path of the .pptx/.potx to import from.</param>
    /// <param name="slides">Source slides, e.g. "1-3,5" (default all).</param>
    TemplateOperationResult ImportSlides(IPresentationBatch batch, string sourcePath, string? slides = null, int? position = null);

    /// <summary>
    /// Opens a template read-only and reports what applying it would change: each layout in use
    /// and the template layout it maps to (or missing), and differences in theme colors, fonts,
    /// and slide size. Changes nothing; apply with presentation apply-template.
    /// </summary>
    /// <param name="templatePath">Full path of the .potx/.pptx template.</param>
    TemplateOperationResult PreviewTemplate(IPresentationBatch batch, string templatePath);
}
