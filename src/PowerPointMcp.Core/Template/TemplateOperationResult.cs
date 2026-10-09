namespace Sbroenne.PowerPointMcp.Core.Template;

/// <summary>A placeholder on a layout.</summary>
public sealed class TemplatePlaceholderInfo
{
    /// <summary>Role used in content keys: title, subtitle, body, picture, or another placeholder role (date, footer, chart, table, ...).</summary>
    public required string Role { get; init; }

    /// <summary>Native PpPlaceholderType name.</summary>
    public required string Type { get; init; }

    /// <summary>Placeholder name.</summary>
    public required string Name { get; init; }

    /// <summary>Left, top, width, height in points.</summary>
    public required IReadOnlyList<float> Bounds { get; init; }
}

/// <summary>A slide layout of a master.</summary>
public sealed class TemplateLayoutInfo
{
    /// <summary>1-based master (design) index.</summary>
    public required int MasterIndex { get; init; }

    /// <summary>Master (design) name.</summary>
    public required string MasterName { get; init; }

    /// <summary>1-based layout index within the master.</summary>
    public required int LayoutIndex { get; init; }

    /// <summary>Layout name, as add-slide takes it.</summary>
    public required string Name { get; init; }

    /// <summary>Slides that use the layout.</summary>
    public required int SlidesUsing { get; init; }

    /// <summary>Placeholders in reading order.</summary>
    public IReadOnlyList<TemplatePlaceholderInfo>? Placeholders { get; init; }
}

/// <summary>A slide added by import-slides.</summary>
public sealed class TemplateImportedSlide
{
    /// <summary>1-based position in this presentation.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID.</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based position in the source presentation.</summary>
    public required int SourceSlide { get; init; }

    /// <summary>Layout in the source presentation.</summary>
    public string? SourceLayout { get; init; }

    /// <summary>Layout it uses here.</summary>
    public string? Layout { get; init; }

    /// <summary>Slide title.</summary>
    public string? Title { get; init; }
}

/// <summary>How the deck's layouts map onto a template.</summary>
public sealed class TemplateLayoutMapping
{
    /// <summary>Layout name in the open deck.</summary>
    public required string Layout { get; init; }

    /// <summary>Slides that use it.</summary>
    public required IReadOnlyList<int> Slides { get; init; }

    /// <summary>matched (same name), similar (name contains or is contained), or missing.</summary>
    public required string Status { get; init; }

    /// <summary>The template layout it maps to.</summary>
    public string? TemplateLayout { get; init; }
}

/// <summary>Result of a template command. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class TemplateOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Layouts (list-layouts, preview-template).</summary>
    public IReadOnlyList<TemplateLayoutInfo>? Layouts { get; init; }

    /// <summary>Slide added or filled.</summary>
    public int? SlideIndex { get; init; }

    /// <summary>SlideID of the slide added or filled.</summary>
    public int? SlideId { get; init; }

    /// <summary>Layout used.</summary>
    public string? LayoutName { get; init; }

    /// <summary>Content keys and the placeholders they filled ("body2 → Content Placeholder 3").</summary>
    public IReadOnlyList<string>? Filled { get; init; }

    /// <summary>Content keys that matched no placeholder.</summary>
    public IReadOnlyList<string>? Unmatched { get; init; }

    /// <summary>Placeholders left empty (they show prompt text in Edit view but not in slide shows).</summary>
    public IReadOnlyList<string>? EmptyPlaceholders { get; init; }

    /// <summary>Slides imported.</summary>
    public IReadOnlyList<TemplateImportedSlide>? ImportedSlides { get; init; }

    /// <summary>Fonts used by imported slides that are not installed on this computer.</summary>
    public IReadOnlyList<string>? MissingFonts { get; init; }

    /// <summary>Validation findings on imported slides, as "slide N code: message".</summary>
    public IReadOnlyList<string>? Findings { get; init; }

    /// <summary>preview-template: where each layout in use would go.</summary>
    public IReadOnlyList<TemplateLayoutMapping>? LayoutMapping { get; init; }

    /// <summary>preview-template: theme colors, fonts, and slide size that differ.</summary>
    public IReadOnlyList<string>? Differences { get; init; }

    /// <summary>Notes and follow-up advice.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
