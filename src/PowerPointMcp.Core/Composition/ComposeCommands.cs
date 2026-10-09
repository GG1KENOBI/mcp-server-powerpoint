extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Assets;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>Slide composition commands.</summary>
public sealed class ComposeCommands : IComposeCommands
{
    private const int MaxRenderPasses = 12;
    private const int MaxSpecBytes = 512 * 1024;

    private static readonly string[] EditedWarning =
        ["Parts of this slide were edited after composition (update-text); the stored JSON does not include those edits."];


    /// <inheritdoc/>
    public ComposeOperationResult Kinds(IPresentationBatch batch) =>
        new() { Success = true, Schema = CompositionSpec.SchemaId, Kinds = CompositionKinds.All };

    /// <inheritdoc/>
    public ComposeOperationResult Plan(IPresentationBatch batch, string? spec = null, string? specPath = null, string? profile = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var prepared = Prepare(batch, spec, specPath, profile);
        if (prepared.Failure is not null)
            return prepared.Failure;

        CompositionPlan plan;
        try
        {
            plan = CompositionPlanner.Plan(prepared.Spec!, prepared.Profile!, prepared.Deck!.Width, prepared.Deck.Height, prepared.Context!);
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message, prepared.Warnings);
        }

        return new ComposeOperationResult
        {
            Success = true,
            Profile = $"{prepared.Profile!.Name}@{prepared.Profile.Version}",
            Planned = plan.Slides.Select(slide => new PlannedSlideSummary
            {
                AppId = slide.AppId,
                Title = slide.Title,
                Items = plan.ItemName is null ? null : slide.ItemCount,
                Elements = slide.Elements.Select(element => new PlannedElementSummary(
                    element.Key,
                    element.Role,
                    element.Type,
                    [element.Box.Left, element.Box.Top, element.Box.Width, element.Box.Height],
                    element.EstimatedTextHeight is { } height ? MathF.Round(height, 1) : null,
                    element.Paragraphs is { Count: > 0 } paragraphs ? paragraphs.Min(paragraph => paragraph.Size) : element.Table?.FontSize,
                    element.Paragraphs is { Count: > 0 } text ? Preview(string.Join(" / ", text.Select(paragraph => paragraph.Text))) : null)).ToList(),
            }).ToList(),
            Warnings = [.. prepared.Warnings, .. plan.Warnings, "Dry run: nothing was created. Heights are estimates; create measures with PowerPoint."],
        };
    }

    /// <inheritdoc/>
    public ComposeOperationResult Create(IPresentationBatch batch, string? spec = null, string? specPath = null, string? profile = null, int? insertAt = null, bool replace = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (insertAt is < 1)
            return Fail("insert_at must be 1 or greater.");
        var prepared = Prepare(batch, spec, specPath, profile);
        if (prepared.Failure is not null)
            return prepared.Failure;
        var deck = prepared.Deck!;
        var compositionSpec = prepared.Spec!;
        var appId = prepared.Context!.AppId;

        if (deck.ExistingSlideIds.Contains(appId) && !replace)
        {
            return Fail($"A slide with id '{appId}' already exists. Pass replace=true to rebuild it in place, or give the composition another id.",
                prepared.Warnings);
        }

        var warnings = new List<string>(prepared.Warnings);
        IReadOnlyList<int>? partition = null;
        for (int pass = 1; pass <= MaxRenderPasses; pass++)
        {
            CompositionPlan plan;
            try
            {
                plan = CompositionPlanner.Plan(compositionSpec, prepared.Profile!, deck.Width, deck.Height, new PlanContext
                {
                    AppId = appId,
                    TemplateTitle = prepared.Context.TemplateTitle,
                    ImageSizes = prepared.Context.ImageSizes,
                    Partition = partition,
                });
            }
            catch (ArgumentException ex)
            {
                return Fail(ex.Message, warnings);
            }

            var outcome = batch.Execute((ctx, ct) =>
            {
                var position = insertAt ?? compositionSpec.InsertAt ?? int.MaxValue;
                if (replace && pass == 1)
                    position = RemoveExisting(ctx.Presentation, appId, position);
                var layouts = LayoutPicker.Pick(ctx.Presentation);
                var rendered = new List<RenderedSlide>();
                for (int i = 0; i < plan.Slides.Count; i++)
                {
                    var slidePosition = position == int.MaxValue ? int.MaxValue : position + i;
                    var result = SlideRenderer.Render(ctx.Presentation, plan.Slides[i], plan, prepared.Profile!, slidePosition, layouts,
                        applyFooter: prepared.Profile!.Footer.Text is not null, ct);
                    rendered.Add(result);
                    if (result.SplitAfter is not null)
                    {
                        // Remove this pass's slides (created in this call) and re-plan with the measured split.
                        foreach (var created in rendered)
                            DeleteSlide(ctx.Presentation, created.SlideId);
                        return (Rendered: rendered, Split: (Slide: i, Fit: result.SplitAfter.Value), Position: position);
                    }
                }
                return (Rendered: rendered, Split: ((int Slide, int Fit)?)null, Position: position);
            });

            if (replace && pass == 1)
                insertAt = outcome.Position == int.MaxValue ? null : outcome.Position;

            if (outcome.Split is { } split)
            {
                var counts = plan.Slides.Select(slide => slide.ItemCount).ToList();
                var remaining = counts.Skip(split.Slide).Sum() - split.Fit;
                partition = [.. counts.Take(split.Slide), split.Fit, .. remaining > 0 ? new[] { remaining } : []];
                warnings.Add($"Pass {pass}: measured text did not fit on slide {split.Slide + 1}; {split.Fit} {plan.ItemName ?? "items"} fit, the rest continue on the next slide.");
                continue;
            }

            warnings.AddRange(plan.Warnings.Where(warning => !warning.StartsWith("Content is estimated", StringComparison.Ordinal) || pass == 1));
            warnings.AddRange(outcome.Rendered.SelectMany(slide => slide.Warnings));
            var overflow = outcome.Rendered.SelectMany(slide => slide.Overflow).ToList();
            return new ComposeOperationResult
            {
                Success = true,
                Profile = $"{prepared.Profile!.Name}@{prepared.Profile.Version}",
                Slides = outcome.Rendered.Select(slide => new ComposedSlide
                {
                    SlideId = slide.SlideId,
                    SlideIndex = slide.SlideIndex,
                    AppId = slide.AppId,
                    Elements = slide.Elements,
                }).ToList(),
                Overflow = overflow.Count > 0 ? overflow : null,
                RenderPasses = pass,
                Warnings = warnings.Count > 0 ? warnings.Distinct().ToList() : null,
            };
        }

        return Fail($"The content could not be fitted after {MaxRenderPasses} measure-and-split passes; nothing was left on the slides from those passes. Split the composition into smaller parts.", warnings);
    }

    /// <inheritdoc/>
    public ComposeOperationResult GetSpec(IPresentationBatch batch, string slideAppId)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(slideAppId))
            return Fail("slide_app_id is required.");
        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            try
            {
                slides = ctx.Presentation.Slides;
                for (int index = 1; index <= slides.Count; index++)
                {
                    PowerPoint.Slide? slide = null;
                    try
                    {
                        slide = slides[index];
                        var tags = DeckSnapshotReader.ReadTags(slide.Tags);
                        if (tags.GetValueOrDefault(DeckRoles.IdTag) != slideAppId)
                            continue;
                        if (!tags.TryGetValue("PPTMCP_SPEC", out var json))
                            return Fail($"Slide {index} has id '{slideAppId}' but stores no composition (it was not created by compose, or it is a continuation slide).");
                        var warnings = tags.ContainsKey("PPTMCP_EDITED") ? EditedWarning : null;
                        return new ComposeOperationResult { Success = true, Spec = json, Schema = CompositionSpec.SchemaId, Warnings = warnings };
                    }
                    finally
                    {
                        if (slide is not null) ComUtilities.Release(ref slide);
                    }
                }
                return Fail($"No slide has id '{slideAppId}'. Use deck find with slideappid:{slideAppId} or deck summary to list ids.");
            }
            finally
            {
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public ComposeOperationResult UpdateText(IPresentationBatch batch, string appId, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);
        if (string.IsNullOrWhiteSpace(appId))
            return Fail("app_id is required.");
        var deckCommands = new DeckCommands();
        var found = deckCommands.Find(batch, $"appid:\"{appId}\"", requireSingle: true);
        if (!found.Success)
            return Fail(found.ErrorMessage ?? $"No object has app id '{appId}'.");
        var target = found.Objects![0];
        if (!target.HasText && target.Kind is not ("text-box" or "placeholder" or "auto-shape"))
            return Fail($"'{appId}' is a {target.Kind} without text.");

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slide? slide = null;
            PowerPoint.Shape? shape = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(ctx.Presentation, target.SlideId);
                shape = slide is null ? null : DeckShapeLocator.FindShape(slide, target.ShapeId);
                if (shape is null || shape.HasTextFrame != Office.MsoTriState.msoTrue)
                    return Fail($"'{appId}' no longer has a text frame.");
                TextRuns.WithRange(shape, (_, range) =>
                {
                    // Assigning Text keeps the formatting of the first character for the new text.
                    range.Text = text.Replace("\r\n", "\v", StringComparison.Ordinal).Replace('\n', '\v');
                    return true;
                });
                var tags = slide!.Tags;
                try
                {
                    tags.Add("PPTMCP_EDITED", "1");
                }
                finally
                {
                    ComUtilities.Release(ref tags);
                }
                var needed = TextRuns.NeededHeight(shape) ?? 0;
                var warnings = needed > shape.Height + 1
                    ? new[] { $"The new text needs {needed.ToString("0.#", CultureInfo.InvariantCulture)} pt but the box is {shape.Height.ToString("0.#", CultureInfo.InvariantCulture)} pt tall. Run review repair (codes text-overflow) or re-create the slide with compose create replace=true." }
                    : null;
                return new ComposeOperationResult
                {
                    Success = true,
                    MeasuredHeight = MathF.Round(needed, 1),
                    BoxHeight = MathF.Round(shape.Height, 1),
                    Warnings = warnings,
                };
            }
            finally
            {
                if (shape is not null) ComUtilities.Release(ref shape);
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    private sealed record DeckInfo(float Width, float Height, int SlideCount, HashSet<string> ExistingSlideIds, Box? TitleBox);

    private sealed class Prepared
    {
        public ComposeOperationResult? Failure { get; init; }

        public CompositionSpec? Spec { get; init; }

        public ResolvedProfile? Profile { get; init; }

        public DeckInfo? Deck { get; init; }

        public PlanContext? Context { get; init; }

        public List<string> Warnings { get; init; } = [];
    }

    /// <summary>Parses the spec, reads the deck, resolves the profile, and reads image headers.</summary>
    private static Prepared Prepare(IPresentationBatch batch, string? spec, string? specPath, string? profileName)
    {
        string json;
        if (spec is not null && specPath is not null)
            return new Prepared { Failure = Fail("Pass spec or spec_path, not both.") };
        if (specPath is not null)
        {
            if (!File.Exists(specPath))
                return new Prepared { Failure = Fail($"spec_path '{specPath}' does not exist.") };
            if (new FileInfo(specPath).Length > MaxSpecBytes)
                return new Prepared { Failure = Fail($"spec_path is larger than {MaxSpecBytes / 1024} KB.") };
            json = File.ReadAllText(specPath);
        }
        else if (spec is not null)
        {
            json = spec;
        }
        else
        {
            return new Prepared { Failure = Fail("Pass the composition JSON in spec, or a file path in spec_path. Call compose kinds for examples.") };
        }

        var parsed = CompositionParser.Parse(json);
        var warnings = parsed.Warnings.ToList();
        if (parsed.Spec is null)
        {
            return new Prepared
            {
                Failure = new ComposeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"The composition is invalid ({parsed.Errors.Count} problem(s)); nothing was created. First: {parsed.Errors[0]}",
                    Errors = parsed.Errors,
                    Warnings = warnings,
                },
            };
        }
        var compositionSpec = parsed.Spec;

        var deck = batch.Execute((ctx, ct) =>
        {
            PowerPoint.PageSetup? setup = null;
            try
            {
                setup = ctx.Presentation.PageSetup;
                var ids = new HashSet<string>(StringComparer.Ordinal);
                int count = ReadSlideIds(ctx.Presentation, ids);
                var layouts = LayoutPicker.Pick(ctx.Presentation);
                return new DeckInfo(setup.SlideWidth, setup.SlideHeight, count, ids, layouts.TitleBox);
            }
            finally
            {
                if (setup is not null) ComUtilities.Release(ref setup);
            }
        });

        var (resolvedProfile, profileError, profileWarnings) = ProfileResolver.Resolve(batch, profileName ?? compositionSpec.Profile);
        warnings.AddRange(profileWarnings);
        if (profileError is not null)
            return new Prepared { Failure = Fail(profileError, warnings) };
        var resolved = resolvedProfile!;

        var images = new Dictionary<string, (int, int)>(StringComparer.Ordinal);
        if (compositionSpec.Image is { } image)
        {
            if (!File.Exists(image.Path))
                return new Prepared { Failure = Fail($"$.image.path: '{image.Path}' does not exist on this computer. Use a full local path; images are never downloaded.", warnings) };
            var info = ImageHeaderReader.ReadFile(image.Path);
            if (info is null)
                return new Prepared { Failure = Fail($"$.image.path: '{image.Path}' is not a readable PNG, JPEG, GIF, BMP, WebP, TIFF, SVG, EMF, or WMF file.", warnings) };
            if (info.Format == "svg")
                warnings.Add("SVG pictures need PowerPoint 2016 or newer (Microsoft 365); older versions reject them. Convert to PNG if creation fails.");
            images[image.Path] = (info.DisplayWidth, info.DisplayHeight);
        }

        var appId = compositionSpec.Id ?? UniqueId(compositionSpec.Kind, deck.ExistingSlideIds);
        return new Prepared
        {
            Spec = compositionSpec,
            Profile = resolved,
            Deck = deck,
            Warnings = warnings,
            Context = new PlanContext
            {
                AppId = appId,
                TemplateTitle = resolved.Name == ThemeProfile.Name ? deck.TitleBox : null,
                ImageSizes = images,
            },
        };
    }

    private static int ReadSlideIds(PowerPoint.Presentation presentation, HashSet<string> ids)
    {
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (int index = 1; index <= slides.Count; index++)
            {
                PowerPoint.Slide? slide = null;
                try
                {
                    slide = slides[index];
                    if (DeckSnapshotReader.ReadTags(slide.Tags).GetValueOrDefault(DeckRoles.IdTag) is { } id)
                        ids.Add(id);
                }
                finally
                {
                    if (slide is not null) ComUtilities.Release(ref slide);
                }
            }
            return slides.Count;
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }

    private static string UniqueId(string kind, HashSet<string> existing)
    {
        for (int n = 1; ; n++)
        {
            var candidate = $"{kind}-{n.ToString(CultureInfo.InvariantCulture)}";
            if (!existing.Contains(candidate) && !existing.Any(id => id.StartsWith(candidate + "~", StringComparison.Ordinal)))
                return candidate;
        }
    }

    /// <summary>Deletes slides with the id or its continuations (id~N); returns the first one's position.</summary>
    private static int RemoveExisting(PowerPoint.Presentation presentation, string appId, int requested)
    {
        var doomed = new List<int>();
        int first = requested;
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (int index = 1; index <= slides.Count; index++)
            {
                PowerPoint.Slide? slide = null;
                try
                {
                    slide = slides[index];
                    var id = DeckSnapshotReader.ReadTags(slide.Tags).GetValueOrDefault(DeckRoles.IdTag);
                    if (id == appId || (id is not null && id.StartsWith(appId + "~", StringComparison.Ordinal)))
                    {
                        if (doomed.Count == 0 && requested == int.MaxValue)
                            first = index;
                        doomed.Add(slide.SlideID);
                    }
                }
                finally
                {
                    if (slide is not null) ComUtilities.Release(ref slide);
                }
            }
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
        foreach (var slideId in doomed)
            DeleteSlide(presentation, slideId);
        return first;
    }

    private static void DeleteSlide(PowerPoint.Presentation presentation, int slideId)
    {
        var slide = DeckShapeLocator.FindSlide(presentation, slideId);
        if (slide is null)
            return;
        try
        {
            slide.Delete();
        }
        finally
        {
            ComUtilities.Release(ref slide);
        }
    }

    private static string Preview(string text) => text.Length <= 80 ? text : text[..80] + "…";

    private static ComposeOperationResult Fail(string message, IReadOnlyList<string>? warnings = null) =>
        new() { Success = false, ErrorMessage = message, Warnings = warnings is { Count: > 0 } ? warnings : null };
}
