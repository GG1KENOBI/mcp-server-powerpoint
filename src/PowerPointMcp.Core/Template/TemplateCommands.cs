extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Assets;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Review;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Template;

/// <inheritdoc cref="ITemplateCommands"/>
public sealed class TemplateCommands : ITemplateCommands
{
    private const int MaxValidatedSlides = 30;

    private readonly ReviewCommands _review = new();

    /// <inheritdoc/>
    public TemplateOperationResult ListLayouts(IPresentationBatch batch, int? masterIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return batch.Execute((ctx, ct) =>
        {
            var usage = LayoutUsage(ctx.Presentation, ct);
            var layouts = ReadLayouts(ctx.Presentation, masterIndex, usage, withPlaceholders: true, out var error);
            return error is not null ? Fail(error) : new TemplateOperationResult { Success = true, Layouts = layouts };
        });
    }

    /// <inheritdoc/>
    public TemplateOperationResult AddSlide(IPresentationBatch batch, string layout, int? position = null, string? content = null, int? masterIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        IReadOnlyList<ContentItem> items = [];
        string? notes = null;
        if (!string.IsNullOrWhiteSpace(content))
        {
            try
            {
                (items, notes) = PlaceholderContent.Parse(content);
            }
            catch (ArgumentException ex)
            {
                return Fail(ex.Message);
            }
        }

        return batch.Execute((ctx, ct) =>
        {
            var layouts = ReadLayouts(ctx.Presentation, masterIndex, null, withPlaceholders: false, out var error);
            if (error is not null)
                return Fail(error);
            var (found, findError) = TemplateText.FindLayout(layouts.Select(item => item.Name).ToList(), layout);
            if (findError is not null)
                return Fail(findError);
            var chosen = layouts[found];

            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.CustomLayout? customLayout = null;
            try
            {
                slides = ctx.Presentation.Slides;
                int at = position ?? slides.Count + 1;
                if (at < 1 || at > slides.Count + 1)
                    return Fail($"position must be between 1 and {slides.Count + 1}.");
                customLayout = FindCustomLayout(ctx.Presentation, chosen.MasterIndex, chosen.LayoutIndex);
                slide = slides.AddSlide(at, customLayout);
                var warnings = new List<string>();
                if (layouts.Count(item => string.Equals(item.Name, chosen.Name, StringComparison.OrdinalIgnoreCase)) > 1)
                    warnings.Add($"Several masters have a layout named '{chosen.Name}'; master {chosen.MasterIndex} was used. Pass master_index to choose.");
                var fill = Fill(slide, items, notes, removeEmpty: false, warnings);
                return new TemplateOperationResult
                {
                    Success = true,
                    SlideIndex = slide.SlideIndex,
                    SlideId = slide.SlideID,
                    LayoutName = chosen.Name,
                    Filled = fill.Filled,
                    Unmatched = NullIfEmpty(fill.Unmatched),
                    EmptyPlaceholders = NullIfEmpty(fill.Empty),
                    Warnings = NullIfEmpty(warnings),
                };
            }
            finally
            {
                if (customLayout is not null) ComUtilities.Release(ref customLayout);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public TemplateOperationResult FillPlaceholders(IPresentationBatch batch, int slideIndex, string content, bool removeEmpty = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        IReadOnlyList<ContentItem> items;
        string? notes;
        try
        {
            (items, notes) = PlaceholderContent.Parse(content);
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.CustomLayout? layout = null;
            try
            {
                slides = ctx.Presentation.Slides;
                if (slideIndex < 1 || slideIndex > slides.Count)
                    return Fail($"Slide index {slideIndex} is out of range (1-{slides.Count}).");
                slide = slides[slideIndex];
                layout = slide.CustomLayout;
                var warnings = new List<string>();
                var fill = Fill(slide, items, notes, removeEmpty, warnings);
                return new TemplateOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    SlideId = slide.SlideID,
                    LayoutName = layout.Name,
                    Filled = fill.Filled,
                    Unmatched = NullIfEmpty(fill.Unmatched),
                    EmptyPlaceholders = NullIfEmpty(fill.Empty),
                    Warnings = NullIfEmpty(warnings),
                };
            }
            finally
            {
                if (layout is not null) ComUtilities.Release(ref layout);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public TemplateOperationResult ImportSlides(IPresentationBatch batch, string sourcePath, string? slides = null, int? position = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(sourcePath))
            return Fail("source_path is required.");

        var imported = batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? deckSlides = null;
            try
            {
                deckSlides = ctx.Presentation.Slides;
                int count = deckSlides.Count;
                int at = position ?? count + 1;
                if (at < 1 || at > count + 1)
                    return (Result: Fail($"position must be between 1 and {count + 1}."), Ids: (List<int>?)null);

                string fullPath;
                IReadOnlyList<int> wanted;
                var sourceInfo = new Dictionary<int, string?>();
                try
                {
                    using var source = SourcePresentation.Open(ctx.Presentation, sourcePath);
                    fullPath = source.Presentation.FullName;
                    PowerPoint.Slides? sourceSlides = null;
                    try
                    {
                        sourceSlides = source.Presentation.Slides;
                        wanted = TemplateText.ParseSlides(slides, sourceSlides.Count);
                        foreach (var number in wanted)
                            sourceInfo[number] = LayoutNameOf(sourceSlides, number);
                    }
                    finally
                    {
                        if (sourceSlides is not null) ComUtilities.Release(ref sourceSlides);
                    }
                }
                catch (ArgumentException ex)
                {
                    return (Result: Fail(ex.Message), Ids: null);
                }
                if (wanted.Count == 0)
                    return (Result: Fail("No slides selected."), Ids: null);

                var added = new List<TemplateImportedSlide>();
                int after = at - 1;
                var sourceOrder = new Queue<int>(wanted);
                foreach (var (first, last) in TemplateText.Runs(wanted))
                {
                    ct.ThrowIfCancellationRequested();
                    int inserted = deckSlides.InsertFromFile(fullPath, after, first, last);
                    for (int offset = 1; offset <= inserted; offset++)
                    {
                        PowerPoint.Slide? slide = null;
                        PowerPoint.CustomLayout? layout = null;
                        try
                        {
                            slide = deckSlides[after + offset];
                            layout = slide.CustomLayout;
                            int sourceSlide = sourceOrder.Count > 0 ? sourceOrder.Dequeue() : first + offset - 1;
                            added.Add(new TemplateImportedSlide
                            {
                                SlideIndex = after + offset,
                                SlideId = slide.SlideID,
                                SourceSlide = sourceSlide,
                                SourceLayout = sourceInfo.GetValueOrDefault(sourceSlide),
                                Layout = layout.Name,
                                Title = TitleOf(slide),
                            });
                        }
                        finally
                        {
                            if (layout is not null) ComUtilities.Release(ref layout);
                            if (slide is not null) ComUtilities.Release(ref slide);
                        }
                    }
                    after += inserted;
                }

                var ids = added.Select(item => item.SlideId).ToHashSet();
                var scan = DesignScanner.Scan(ctx.Presentation, null, ids.Contains, ct);
                var installed = FontCatalog.InstalledFamilies();
                var warnings = new List<string>();
                List<string>? missing = null;
                if (installed is null)
                {
                    warnings.Add("Installed fonts could not be read; missing fonts were not checked.");
                }
                else
                {
                    missing = scan.Runs.Select(run => run.FontName).Where(name => !string.IsNullOrWhiteSpace(name) && !name.StartsWith('+'))
                        .Distinct(StringComparer.OrdinalIgnoreCase).Where(name => !installed.Contains(name)).Order(StringComparer.OrdinalIgnoreCase).ToList();
                    if (missing.Count > 0)
                        warnings.Add("Text in missing fonts is shown with a substitute font and may wrap differently; use design normalize-typography to switch to the profile fonts.");
                }
                var changedLayouts = added.Where(item => item.SourceLayout is not null && !string.Equals(item.SourceLayout, item.Layout, StringComparison.OrdinalIgnoreCase)).ToList();
                if (changedLayouts.Count > 0)
                    warnings.Add($"{changedLayouts.Count} slide(s) now use a different layout than in the source; check their placeholders.");
                warnings.Add("Imported slides take this presentation's theme. The deck is changed in memory only.");
                return (Result: new TemplateOperationResult
                {
                    Success = true,
                    ImportedSlides = added,
                    MissingFonts = NullIfEmpty(missing ?? []),
                    Warnings = warnings,
                }, Ids: added.Select(item => item.SlideIndex).ToList());
            }
            finally
            {
                if (deckSlides is not null) ComUtilities.Release(ref deckSlides);
            }
        });

        if (!imported.Result.Success || imported.Ids is null)
            return imported.Result;

        // Validation runs its own Execute calls, so it happens after the import.
        var findings = new List<string>();
        foreach (var slideIndex in imported.Ids.Take(MaxValidatedSlides))
        {
            var review = _review.Validate(batch, slideIndex, minSeverity: "warning");
            if (review.Success)
                findings.AddRange((review.Findings ?? []).Select(finding => $"slide {finding.SlideIndex.ToString(CultureInfo.InvariantCulture)} {finding.Code}: {finding.Message}"));
        }
        var notes = imported.Result.Warnings?.ToList() ?? [];
        if (imported.Ids.Count > MaxValidatedSlides)
            notes.Add($"Only the first {MaxValidatedSlides} imported slides were validated; run review validate for the rest.");
        return new TemplateOperationResult
        {
            Success = true,
            ImportedSlides = imported.Result.ImportedSlides,
            MissingFonts = imported.Result.MissingFonts,
            Findings = NullIfEmpty(findings),
            Warnings = notes,
        };
    }

    /// <inheritdoc/>
    public TemplateOperationResult PreviewTemplate(IPresentationBatch batch, string templatePath)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(templatePath))
            return Fail("template_path is required.");
        return batch.Execute((ctx, ct) =>
        {
            var usage = LayoutUsage(ctx.Presentation, ct);
            var (colors, heading, body) = DesignApplier.ReadTheme(ctx.Presentation);
            var (width, height) = SlideSize(ctx.Presentation);

            List<TemplateLayoutInfo> templateLayouts;
            Dictionary<string, string> templateColors;
            string? templateHeading, templateBody;
            float templateWidth, templateHeight;
            try
            {
                using var source = SourcePresentation.Open(ctx.Presentation, templatePath);
                templateLayouts = ReadLayouts(source.Presentation, null, null, withPlaceholders: true, out _);
                (templateColors, templateHeading, templateBody) = DesignApplier.ReadTheme(source.Presentation);
                (templateWidth, templateHeight) = SlideSize(source.Presentation);
            }
            catch (ArgumentException ex)
            {
                return Fail(ex.Message);
            }

            var names = templateLayouts.Select(item => item.Name).ToList();
            var mapping = new List<TemplateLayoutMapping>();
            foreach (var group in usage.GroupBy(entry => entry.Value.Layout, StringComparer.OrdinalIgnoreCase).OrderBy(group => group.Min(entry => entry.Key)))
            {
                var exact = names.FirstOrDefault(name => string.Equals(name, group.Key, StringComparison.OrdinalIgnoreCase));
                var (similar, _) = exact is null ? TemplateText.FindLayout(names, group.Key) : (-1, null);
                mapping.Add(new TemplateLayoutMapping
                {
                    Layout = group.Key,
                    Slides = group.Select(entry => entry.Key).Order().ToList(),
                    Status = exact is not null ? "matched" : similar >= 0 ? "similar" : "missing",
                    TemplateLayout = exact ?? (similar >= 0 ? names[similar] : null),
                });
            }

            var differences = new List<string>();
            foreach (var (role, hex) in colors.OrderBy(pair => pair.Key, StringComparer.Ordinal))
            {
                if (templateColors.TryGetValue(role, out var other) && !string.Equals(hex, other, StringComparison.OrdinalIgnoreCase))
                    differences.Add($"theme color {role}: {hex} → {other}");
            }
            if (!string.Equals(heading, templateHeading, StringComparison.OrdinalIgnoreCase))
                differences.Add($"heading font: {heading} → {templateHeading}");
            if (!string.Equals(body, templateBody, StringComparison.OrdinalIgnoreCase))
                differences.Add($"body font: {body} → {templateBody}");
            if (MathF.Abs(width - templateWidth) > 0.5f || MathF.Abs(height - templateHeight) > 0.5f)
                differences.Add(string.Create(CultureInfo.InvariantCulture, $"slide size: {width:0.#}×{height:0.#} pt → {templateWidth:0.#}×{templateHeight:0.#} pt ({DeckGeometry.AspectRatio(width, height)} → {DeckGeometry.AspectRatio(templateWidth, templateHeight)})"));

            var warnings = new List<string>();
            if (mapping.Any(item => item.Status == "missing"))
                warnings.Add("Slides on missing layouts get a layout of the same kind or keep their current one; check them after presentation apply-template.");
            if (differences.Any(item => item.StartsWith("slide size", StringComparison.Ordinal)))
                warnings.Add("The slide size differs; content positioned for the current size may need re-layout.");
            warnings.Add("Nothing was changed. Save a copy (file save-as) before presentation apply-template to keep the original.");
            return new TemplateOperationResult
            {
                Success = true,
                Layouts = templateLayouts,
                LayoutMapping = mapping,
                Differences = differences,
                Warnings = warnings,
            };
        });
    }

    private sealed record FillOutcome(List<string> Filled, List<string> Unmatched, List<string> Empty);

    private static FillOutcome Fill(PowerPoint.Slide slide, IReadOnlyList<ContentItem> items, string? notes, bool removeEmpty, List<string> warnings)
    {
        var filled = new List<string>();
        var unmatched = new List<string>();
        var slots = ReadSlots(slide);
        var assignment = PlaceholderContent.Assign(items, slots);
        unmatched.AddRange(assignment.Unmatched.Select(item => item.Key));
        if (assignment.Unmatched.Count > 0)
            warnings.Add($"No placeholder for {string.Join(", ", assignment.Unmatched.Select(item => item.Key))} on this layout. Placeholders: {string.Join(", ", slots.Select(slot => $"{slot.Role}{(slot.Ordinal > 1 ? slot.Ordinal.ToString(CultureInfo.InvariantCulture) : "")} ({slot.Name})"))}.");

        var removed = new HashSet<int>();
        foreach (var (item, slot) in assignment.Matches)
        {
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            try
            {
                shapes = slide.Shapes;
                shape = DeckShapeLocator.FindShape(slide, slot.ShapeId);
                if (shape is null)
                    continue;
                if (item.PicturePath is { } path)
                {
                    var error = PlacePicture(shapes, shape, slot, path);
                    if (error is not null)
                    {
                        unmatched.Add(item.Key);
                        warnings.Add(error);
                        continue;
                    }
                    removed.Add(slot.ShapeId);
                }
                else
                {
                    WriteParagraphs(shape, item.Paragraphs!);
                }
                filled.Add($"{item.Key} → {slot.Name}");
            }
            finally
            {
                if (shape is not null) ComUtilities.Release(ref shape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
            }
        }
        if (notes is not null)
        {
            if (WriteNotes(slide, notes))
                filled.Add("notes → speaker notes");
            else
                warnings.Add("The slide has no notes placeholder; notes were not written.");
        }

        var empty = new List<string>();
        foreach (var slot in ReadSlots(slide).Where(slot => !slot.HasText && !removed.Contains(slot.ShapeId)))
        {
            if (removeEmpty)
            {
                var shape = DeckShapeLocator.FindShape(slide, slot.ShapeId);
                if (shape is not null)
                {
                    try
                    {
                        shape.Delete();
                    }
                    finally
                    {
                        ComUtilities.Release(ref shape!);
                    }
                }
            }
            else
            {
                empty.Add($"{slot.Role}: {slot.Name}");
            }
        }
        if (empty.Count > 0)
            warnings.Add("Empty placeholders show prompt text while editing; fill them or pass remove_empty=true.");
        return new FillOutcome(filled, unmatched, empty);
    }

    private static string? PlacePicture(PowerPoint.Shapes shapes, PowerPoint.Shape placeholder, PlaceholderSlot slot, string path)
    {
        string full;
        try
        {
            full = System.IO.Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return $"Invalid picture path '{path}': {ex.Message}";
        }
        if (!File.Exists(full))
            return $"The picture '{full}' does not exist on this computer (pictures must be local files; nothing is downloaded).";
        var header = ImageHeaderReader.ReadFile(full);
        if (header is null || header.Aspect <= 0)
            return $"'{full}' is not a supported image (PNG, JPEG, GIF, BMP, TIFF, WebP, SVG, EMF, WMF).";
        var frame = new Box(slot.Left, slot.Top, slot.Left + slot.Width, slot.Top + slot.Height);
        var fit = slot.Role == "picture" ? "cover" : "contain";
        var placement = PictureFit.Compute(frame, header.Aspect, fit);
        var picture = PicturePlacer.Place(shapes, new PicturePlan(full, placement, null, header.DisplayWidth, header.DisplayHeight, fit, null));
        try
        {
            picture.Name = slot.Name;
            placeholder.Delete();
        }
        finally
        {
            ComUtilities.Release(ref picture!);
        }
        return null;
    }

    private static void WriteParagraphs(PowerPoint.Shape shape, IReadOnlyList<ContentParagraph> paragraphs)
    {
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            frame = shape.TextFrame;
            range = frame.TextRange;
            range.Text = string.Join("\r", paragraphs.Select(paragraph => paragraph.Text));
            for (int index = 1; index <= paragraphs.Count; index++)
            {
                if (paragraphs[index - 1].Level == 1)
                    continue;
                PowerPoint.TextRange? paragraph = null;
                try
                {
                    paragraph = range.Paragraphs(index, 1);
                    paragraph.IndentLevel = paragraphs[index - 1].Level;
                }
                finally
                {
                    if (paragraph is not null) ComUtilities.Release(ref paragraph);
                }
            }
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    private static bool WriteNotes(PowerPoint.Slide slide, string notes)
    {
        PowerPoint.SlideRange? page = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            page = slide.NotesPage;
            shapes = page.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.PlaceholderFormat? format = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        continue;
                    format = shape.PlaceholderFormat;
                    if (format.Type != PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        continue;
                    WriteParagraphs(shape, [new ContentParagraph(TextFrame.TextRanges.ToPowerPointParagraphs(notes), 1)]);
                    return true;
                }
                finally
                {
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
            return false;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (page is not null) ComUtilities.Release(ref page);
        }
    }

    private static List<PlaceholderSlot> ReadSlots(PowerPoint.Slide slide)
    {
        var raw = new List<(int Id, string Name, string Role, float Left, float Top, float Width, float Height, bool HasText)>();
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.PlaceholderFormat? format = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        continue;
                    format = shape.PlaceholderFormat;
                    var role = SlotRole(format.Type);
                    if (role is null)
                        continue;
                    bool hasText = shape.HasTextFrame == Office.MsoTriState.msoTrue && HasText(shape);
                    raw.Add((shape.Id, shape.Name, role, shape.Left, shape.Top, shape.Width, shape.Height, hasText));
                }
                finally
                {
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
        return raw.GroupBy(item => item.Role)
            .SelectMany(group => group.OrderBy(item => MathF.Round(item.Top / 10f)).ThenBy(item => item.Left)
                .Select((item, order) => new PlaceholderSlot(item.Id, item.Name, item.Role, order + 1, item.Left, item.Top, item.Width, item.Height, item.HasText)))
            .ToList();
    }

    private static bool HasText(PowerPoint.Shape shape)
    {
        PowerPoint.TextFrame? frame = null;
        try
        {
            frame = shape.TextFrame;
            return frame.HasText == Office.MsoTriState.msoTrue;
        }
        finally
        {
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    /// <summary>Content role of a placeholder type, or null for placeholders content does not fill (date, footer, slide number, ...).</summary>
    private static string? SlotRole(PowerPoint.PpPlaceholderType type) => DeckRoles.FromPlaceholder(type.ToString()) switch
    {
        "title" => "title",
        "subtitle" => "subtitle",
        "body" or "object" => "body",
        "picture" => "picture",
        _ => null,
    };

    private static List<TemplateLayoutInfo> ReadLayouts(
        PowerPoint.Presentation presentation,
        int? masterIndex,
        Dictionary<int, (string Master, string Layout)>? usage,
        bool withPlaceholders,
        out string? error)
    {
        error = null;
        var result = new List<TemplateLayoutInfo>();
        PowerPoint.Designs? designs = null;
        try
        {
            designs = presentation.Designs;
            if (masterIndex is { } only && (only < 1 || only > designs.Count))
            {
                error = $"master_index {only} is out of range (1-{designs.Count}).";
                return result;
            }
            for (int index = 1; index <= designs.Count; index++)
            {
                if (masterIndex is { } wanted && wanted != index)
                    continue;
                PowerPoint.Design? design = null;
                PowerPoint.Master? master = null;
                PowerPoint.CustomLayouts? layouts = null;
                try
                {
                    design = designs[index];
                    master = design.SlideMaster;
                    layouts = master.CustomLayouts;
                    string masterName = design.Name;
                    for (int layoutIndex = 1; layoutIndex <= layouts.Count; layoutIndex++)
                    {
                        PowerPoint.CustomLayout? layout = null;
                        try
                        {
                            layout = layouts[layoutIndex];
                            string name = layout.Name;
                            result.Add(new TemplateLayoutInfo
                            {
                                MasterIndex = index,
                                MasterName = masterName,
                                LayoutIndex = layoutIndex,
                                Name = name,
                                SlidesUsing = usage?.Values.Count(entry => entry.Master == masterName && entry.Layout == name) ?? 0,
                                Placeholders = withPlaceholders ? ReadLayoutPlaceholders(layout) : null,
                            });
                        }
                        finally
                        {
                            if (layout is not null) ComUtilities.Release(ref layout);
                        }
                    }
                }
                finally
                {
                    if (layouts is not null) ComUtilities.Release(ref layouts);
                    if (master is not null) ComUtilities.Release(ref master);
                    if (design is not null) ComUtilities.Release(ref design);
                }
            }
        }
        finally
        {
            if (designs is not null) ComUtilities.Release(ref designs);
        }
        return result;
    }

    private static List<TemplatePlaceholderInfo> ReadLayoutPlaceholders(PowerPoint.CustomLayout layout)
    {
        var result = new List<(TemplatePlaceholderInfo Info, float Top, float Left)>();
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = layout.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.PlaceholderFormat? format = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        continue;
                    format = shape.PlaceholderFormat;
                    var type = format.Type;
                    var role = SlotRole(type) ?? DeckRoles.FromPlaceholder(type.ToString()) ?? "other";
                    result.Add((new TemplatePlaceholderInfo
                    {
                        Role = role,
                        Type = type.ToString(),
                        Name = shape.Name,
                        Bounds = [Round(shape.Left), Round(shape.Top), Round(shape.Width), Round(shape.Height)],
                    }, shape.Top, shape.Left));
                }
                finally
                {
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
        return result.OrderBy(entry => MathF.Round(entry.Top / 10f)).ThenBy(entry => entry.Left).Select(entry => entry.Info).ToList();
    }

    /// <summary>Master name and layout name of every slide, by slide index.</summary>
    private static Dictionary<int, (string Master, string Layout)> LayoutUsage(PowerPoint.Presentation presentation, CancellationToken cancellationToken)
    {
        var usage = new Dictionary<int, (string, string)>();
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (int index = 1; index <= slides.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPoint.Slide? slide = null;
                PowerPoint.Design? design = null;
                PowerPoint.CustomLayout? layout = null;
                try
                {
                    slide = slides[index];
                    design = slide.Design;
                    layout = slide.CustomLayout;
                    usage[index] = (design.Name, layout.Name);
                }
                finally
                {
                    if (layout is not null) ComUtilities.Release(ref layout);
                    if (design is not null) ComUtilities.Release(ref design);
                    if (slide is not null) ComUtilities.Release(ref slide);
                }
            }
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
        return usage;
    }

    private static PowerPoint.CustomLayout FindCustomLayout(PowerPoint.Presentation presentation, int masterIndex, int layoutIndex)
    {
        PowerPoint.Designs? designs = null;
        PowerPoint.Design? design = null;
        PowerPoint.Master? master = null;
        PowerPoint.CustomLayouts? layouts = null;
        try
        {
            designs = presentation.Designs;
            design = designs[masterIndex];
            master = design.SlideMaster;
            layouts = master.CustomLayouts;
            return layouts[layoutIndex];
        }
        finally
        {
            if (layouts is not null) ComUtilities.Release(ref layouts);
            if (master is not null) ComUtilities.Release(ref master);
            if (design is not null) ComUtilities.Release(ref design);
            if (designs is not null) ComUtilities.Release(ref designs);
        }
    }

    private static string? LayoutNameOf(PowerPoint.Slides slides, int index)
    {
        PowerPoint.Slide? slide = null;
        PowerPoint.CustomLayout? layout = null;
        try
        {
            slide = slides[index];
            layout = slide.CustomLayout;
            return layout.Name;
        }
        finally
        {
            if (layout is not null) ComUtilities.Release(ref layout);
            if (slide is not null) ComUtilities.Release(ref slide);
        }
    }

    private static string? TitleOf(PowerPoint.Slide slide)
    {
        PowerPoint.Shapes? shapes = null;
        PowerPoint.Shape? title = null;
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            shapes = slide.Shapes;
            if (shapes.HasTitle != Office.MsoTriState.msoTrue)
                return null;
            title = shapes.Title;
            frame = title.TextFrame;
            range = frame.TextRange;
            var text = range.Text.Replace('\r', ' ').Replace('\v', ' ').Trim();
            return text.Length == 0 ? null : text;
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
            if (title is not null) ComUtilities.Release(ref title);
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    private static (float Width, float Height) SlideSize(PowerPoint.Presentation presentation)
    {
        PowerPoint.PageSetup? setup = null;
        try
        {
            setup = presentation.PageSetup;
            return (setup.SlideWidth, setup.SlideHeight);
        }
        finally
        {
            if (setup is not null) ComUtilities.Release(ref setup);
        }
    }

    private static float Round(float value) => MathF.Round(value * 10f) / 10f;

    private static List<string>? NullIfEmpty(List<string> values) => values.Count == 0 ? null : values;

    private static TemplateOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
