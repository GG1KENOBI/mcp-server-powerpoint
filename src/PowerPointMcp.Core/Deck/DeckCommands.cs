using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Master;
using Sbroenne.PowerPointMcp.Core.Presentation;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>Deck inspection, addressing, and identity commands.</summary>
public sealed partial class DeckCommands : IDeckCommands
{
    private const int MaxPageSize = 200;

    private readonly MasterCommands _masters = new();
    private readonly PresentationCommands _presentation = new();

    /// <inheritdoc/>
    public DeckOperationResult Summary(IPresentationBatch batch, int startSlide = 1, int maxSlides = 25, bool includeTheme = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (maxSlides is < 1 or > 100)
            return Fail("max_slides must be between 1 and 100.");
        if (startSlide < 1)
            return Fail("start_slide must be 1 or greater.");

        var core = batch.Execute((ctx, ct) =>
        {
            PowerPoint.PageSetup? pageSetup = null;
            PowerPoint.Slides? slides = null;
            try
            {
                pageSetup = ctx.Presentation.PageSetup;
                slides = ctx.Presentation.Slides;
                var sectionNames = DeckSnapshotReader.ReadSectionNames(ctx.Presentation);
                var sections = ReadSections(ctx.Presentation, sectionNames.Count);
                int count = slides.Count;
                var page = new List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>();
                for (int index = startSlide; index <= Math.Min(count, startSlide + maxSlides - 1); index++)
                {
                    ct.ThrowIfCancellationRequested();
                    PowerPoint.Slide? slide = null;
                    try
                    {
                        slide = slides[index];
                        page.Add(DeckSnapshotReader.ReadSlideWithObjects(slide, index, sectionNames, detailed: false, ct));
                    }
                    finally
                    {
                        if (slide is not null) ComUtilities.Release(ref slide);
                    }
                }

                return new DeckOperationResult
                {
                    Success = true,
                    SlideWidth = pageSetup.SlideWidth,
                    SlideHeight = pageSetup.SlideHeight,
                    AspectRatio = DeckGeometry.AspectRatio(pageSetup.SlideWidth, pageSetup.SlideHeight),
                    SlideCount = count,
                    Sections = sections,
                    Slides = page.Select(entry => entry.Slide).ToList(),
                    // The revision covers every slide, so it is only known when this page is the whole deck.
                    Revision = startSlide == 1 && page.Count == count ? Revision(page) : null,
                    TotalCount = count,
                    Offset = startSlide - 1,
                    HasMore = startSlide - 1 + page.Count < count,
                };
            }
            finally
            {
                if (slides is not null) ComUtilities.Release(ref slides);
                if (pageSetup is not null) ComUtilities.Release(ref pageSetup);
            }
        });

        if (!includeTheme)
            return core;

        // Reuse the existing domain commands for theme, masters, and properties.
        var warnings = new List<string>();
        var theme = _presentation.GetThemeName(batch);
        var masters = _masters.ListMasters(batch);
        var colors = _masters.GetThemeColors(batch, 1);
        var fonts = _masters.GetThemeFonts(batch, 1);
        if (!masters.Success) warnings.Add($"Masters not read: {masters.ErrorMessage}");
        if (!colors.Success) warnings.Add($"Theme colors not read: {colors.ErrorMessage}");
        if (!fonts.Success) warnings.Add($"Theme fonts not read: {fonts.ErrorMessage}");

        var properties = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var name in new[] { "Title", "Subject", "Author" })
        {
            var property = _presentation.GetDocumentProperty(batch, name);
            if (property.Success && !string.IsNullOrEmpty(property.PropertyValue))
                properties[name] = property.PropertyValue;
        }

        return new DeckOperationResult
        {
            Success = true,
            SlideWidth = core.SlideWidth,
            SlideHeight = core.SlideHeight,
            AspectRatio = core.AspectRatio,
            SlideCount = core.SlideCount,
            Sections = core.Sections,
            ThemeName = theme.ThemeName,
            ThemeColors = colors.ThemeColors,
            ThemeFonts = fonts.Success
                ? new Dictionary<string, string?>(StringComparer.Ordinal)
                {
                    ["heading"] = fonts.MajorThemeFonts?.GetValueOrDefault("latin"),
                    ["body"] = fonts.MinorThemeFonts?.GetValueOrDefault("latin"),
                }
                : null,
            Masters = masters.Masters,
            Properties = properties,
            Slides = core.Slides,
            Revision = core.Revision,
            TotalCount = core.TotalCount,
            Offset = core.Offset,
            HasMore = core.HasMore,
            Warnings = warnings.Count > 0 ? warnings : null,
        };
    }

    /// <inheritdoc/>
    public DeckOperationResult InspectObjects(
        IPresentationBatch batch,
        int? slideIndex = null,
        int? slideId = null,
        string? selector = null,
        string detail = "compact",
        int offset = 0,
        int limit = 50)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (detail is not ("compact" or "full"))
            return Fail("detail must be compact or full.");
        if (offset < 0)
            return Fail("offset must be 0 or greater.");
        if (limit is < 1 or > MaxPageSize)
            return Fail($"limit must be between 1 and {MaxPageSize}.");
        if (slideIndex is null && slideId is null && string.IsNullOrWhiteSpace(selector))
            return Fail("Pass slide_index, slide_id, or a selector (for example kind:table) to choose what to inspect.");

        ObjectSelector? parsed;
        try
        {
            parsed = string.IsNullOrWhiteSpace(selector) ? null : ObjectSelector.Parse(selector);
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }

        return batch.Execute((ctx, ct) =>
        {
            var target = ReadTargetSlides(ctx.Presentation, slideIndex, slideId, ct, out var error);
            if (error is not null)
                return Fail(error);

            var matches = Select(target, parsed);
            var page = matches.Skip(offset).Take(limit)
                .Select(item => detail == "full" ? item : DeckSnapshotReader.Compact(item))
                .ToList();
            return new DeckOperationResult
            {
                Success = true,
                Slides = target.Select(entry => entry.Slide).ToList(),
                Objects = page,
                TotalCount = matches.Count,
                Offset = offset,
                HasMore = offset + page.Count < matches.Count,
                DuplicateAppIds = NullIfEmpty(AppIdPlanner.FindDuplicates(target)),
            };
        });
    }

    /// <inheritdoc/>
    public DeckOperationResult Find(IPresentationBatch batch, string selector, bool requireSingle = false, int limit = 20)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (limit is < 1 or > MaxPageSize)
            return Fail($"limit must be between 1 and {MaxPageSize}.");

        ObjectSelector parsed;
        try
        {
            parsed = ObjectSelector.Parse(selector ?? "");
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }

        return batch.Execute((ctx, ct) =>
        {
            var deck = DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: true, ct);
            var matches = Select(deck, parsed);
            if (requireSingle)
            {
                try
                {
                    ObjectSelector.RequireSingle(selector!, matches, deck.Sum(entry => entry.Objects.Count));
                }
                catch (ArgumentException ex)
                {
                    return new DeckOperationResult
                    {
                        Success = false,
                        ErrorMessage = ex.Message,
                        Objects = matches.Take(limit).Select(DeckSnapshotReader.Compact).ToList(),
                        TotalCount = matches.Count,
                    };
                }
            }

            return new DeckOperationResult
            {
                Success = true,
                Objects = matches.Take(limit).Select(DeckSnapshotReader.Compact).ToList(),
                TotalCount = matches.Count,
                Offset = 0,
                HasMore = matches.Count > limit,
                DuplicateAppIds = NullIfEmpty(AppIdPlanner.FindDuplicates(deck)),
            };
        });
    }

    /// <inheritdoc/>
    public DeckOperationResult Fingerprint(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return batch.Execute((ctx, ct) =>
        {
            var deck = DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: false, ct);
            return new DeckOperationResult
            {
                Success = true,
                SlideCount = deck.Count,
                Revision = Revision(deck),
                DuplicateAppIds = NullIfEmpty(AppIdPlanner.FindDuplicates(deck)),
                Slides = deck.Select(entry => new DeckSlideInfo
                {
                    SlideIndex = entry.Slide.SlideIndex,
                    SlideId = entry.Slide.SlideId,
                    Hidden = entry.Slide.Hidden,
                    ShapeCount = entry.Slide.ShapeCount,
                    Title = entry.Slide.Title,
                    Fingerprint = entry.Slide.Fingerprint,
                }).ToList(),
            };
        });
    }

    /// <inheritdoc/>
    public DeckOperationResult AssignIds(
        IPresentationBatch batch,
        string? selector = null,
        string? appId = null,
        string prefix = "id",
        bool dryRun = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(prefix) || prefix.Any(char.IsWhiteSpace))
            return Fail("prefix must be a non-empty word without spaces.");
        if (appId is not null && (string.IsNullOrWhiteSpace(appId) || appId.Any(char.IsWhiteSpace)))
            return Fail("app_id must be a non-empty word without spaces.");
        if (appId is not null && string.IsNullOrWhiteSpace(selector))
            return Fail("app_id needs a selector that matches exactly one object.");

        ObjectSelector? parsed;
        try
        {
            parsed = string.IsNullOrWhiteSpace(selector) ? null : ObjectSelector.Parse(selector);
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }

        return batch.Execute((ctx, ct) =>
        {
            var deck = DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: true, ct);
            IReadOnlyList<AppIdAssignment> plan;
            if (appId is not null)
            {
                var matches = Select(deck, parsed);
                DeckObjectInfo target;
                try
                {
                    target = ObjectSelector.RequireSingle(selector!, matches, deck.Sum(entry => entry.Objects.Count));
                }
                catch (ArgumentException ex)
                {
                    return Fail(ex.Message);
                }
                var owner = deck.SelectMany(entry => entry.Objects)
                    .FirstOrDefault(item => item.AppId == appId && (item.SlideId, item.ShapeId) != (target.SlideId, target.ShapeId));
                if (owner is not null)
                    return Fail($"app_id '{appId}' is already used by {ObjectSelector.Describe(owner)}.");
                plan =
                [
                    new AppIdAssignment
                    {
                        SlideId = target.SlideId,
                        SlideIndex = target.SlideIndex,
                        ShapeId = target.ShapeId,
                        OldAppId = target.AppId,
                        NewAppId = appId,
                        Reason = "requested",
                    },
                ];
            }
            else
            {
                var toTag = parsed is null ? null : Select(deck, parsed);
                plan = AppIdPlanner.Plan(deck, toTag, prefix);
            }

            if (!dryRun)
                ApplyAssignments(ctx.Presentation, plan);

            return new DeckOperationResult
            {
                Success = true,
                Assignments = plan,
                DuplicateAppIds = dryRun ? NullIfEmpty(AppIdPlanner.FindDuplicates(deck)) : null,
                Warnings = dryRun ? ["Dry run: nothing was changed."] : null,
            };
        });
    }

    private static void ApplyAssignments(PowerPoint.Presentation presentation, IReadOnlyList<AppIdAssignment> plan)
    {
        foreach (var assignment in plan)
        {
            PowerPoint.Slide? slide = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.Tags? tags = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(presentation, assignment.SlideId)
                    ?? throw new InvalidOperationException($"Slide {assignment.SlideId} disappeared while assigning ids.");
                if (assignment.ShapeId is { } shapeId)
                {
                    shape = DeckShapeLocator.FindShape(slide, shapeId)
                        ?? throw new InvalidOperationException($"Shape {shapeId} disappeared while assigning ids.");
                    tags = shape.Tags;
                }
                else
                {
                    tags = slide.Tags;
                }
                tags.Add(DeckRoles.IdTag, assignment.NewAppId);
            }
            finally
            {
                if (tags is not null) ComUtilities.Release(ref tags);
                if (shape is not null) ComUtilities.Release(ref shape);
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        }
    }

    internal static List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> ReadTargetSlides(
        PowerPoint.Presentation presentation,
        int? slideIndex,
        int? slideId,
        CancellationToken cancellationToken,
        out string? error)
    {
        error = null;
        if (slideIndex is null && slideId is null)
            return DeckSnapshotReader.ReadAll(presentation, detailed: true, cancellationToken);

        PowerPoint.Slides? slides = null;
        PowerPoint.Slide? slide = null;
        try
        {
            slides = presentation.Slides;
            int index;
            if (slideId is { } id)
            {
                slide = DeckShapeLocator.FindSlide(presentation, id);
                if (slide is null)
                {
                    error = $"No slide has slide_id {id}. Use deck summary to list slide ids.";
                    return [];
                }
                index = slide.SlideIndex;
                if (slideIndex is { } requested && requested != index)
                {
                    error = $"slide_index {requested} and slide_id {id} refer to different slides (slide_id {id} is at position {index}).";
                    return [];
                }
            }
            else
            {
                index = slideIndex!.Value;
                if (index < 1 || index > slides.Count)
                {
                    error = $"Slide index {index} is out of range. The presentation has {slides.Count} slide(s) (valid range: 1-{slides.Count}).";
                    return [];
                }
                slide = slides[index];
            }

            var sectionNames = DeckSnapshotReader.ReadSectionNames(presentation);
            return [DeckSnapshotReader.ReadSlideWithObjects(slide, index, sectionNames, detailed: true, cancellationToken)];
        }
        finally
        {
            if (slide is not null) ComUtilities.Release(ref slide);
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }

    internal static List<DeckObjectInfo> Select(
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck,
        ObjectSelector? selector) =>
        selector is null
            ? deck.SelectMany(entry => entry.Objects).ToList()
            : selector.Find(deck).ToList();

    internal static string Revision(IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck) =>
        DeckFingerprint.ForDeck(deck.Select(entry => (entry.Slide.SlideId, entry.Slide.Fingerprint ?? "")));

    private static List<DeckSectionInfo> ReadSections(PowerPoint.Presentation presentation, int count)
    {
        PowerPoint.SectionProperties? sections = null;
        try
        {
            sections = presentation.SectionProperties;
            var result = new List<DeckSectionInfo>(count);
            for (int index = 1; index <= count; index++)
            {
                result.Add(new DeckSectionInfo
                {
                    SectionIndex = index,
                    Name = sections.Name(index),
                    FirstSlide = sections.SlidesCount(index) > 0 ? sections.FirstSlide(index) : 0,
                    SlideCount = sections.SlidesCount(index),
                });
            }
            return result;
        }
        finally
        {
            if (sections is not null) ComUtilities.Release(ref sections);
        }
    }

    internal static DeckOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };

    private static IReadOnlyList<string>? NullIfEmpty(IReadOnlyList<string> values) => values.Count == 0 ? null : values;
}
