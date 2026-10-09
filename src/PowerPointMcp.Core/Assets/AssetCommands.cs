extern alias OfficeInterop;

using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <inheritdoc cref="IAssetCommands"/>
public sealed class AssetCommands : IAssetCommands
{
    internal const string DecorativeTag = "PPTMCP_DECORATIVE";
    internal const string SourceTag = "PPTMCP_IMG_SOURCE";

    /// <inheritdoc/>
    public AssetOperationResult ScanFolder(IPresentationBatch batch, string folder, bool recursive = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(folder))
            return Fail("folder is required.");
        return WithCatalog(catalog =>
        {
            var summary = catalog.Scan(folder, recursive);
            catalog.Save();
            return new AssetOperationResult
            {
                Success = true,
                CatalogPath = catalog.File,
                Scan = summary,
                TotalCount = catalog.Entries.Count,
                Warnings = summary.Notes.Count > 0 ? summary.Notes : null,
            };
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult Search(IPresentationBatch batch, string? query = null, string? orientation = null, int minWidth = 0, string? tag = null, int limit = 20)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (limit is < 1 or > 200)
            return Fail("limit must be between 1 and 200.");
        if (orientation is not null && orientation is not ("landscape" or "portrait" or "square"))
            return Fail("orientation must be landscape, portrait, or square.");
        return WithCatalog(catalog =>
        {
            if (catalog.Entries.Count == 0)
                return new AssetOperationResult { Success = true, CatalogPath = catalog.File, Assets = [], TotalCount = 0, Warnings = ["The catalog is empty; run asset scan-folder on a local image folder first."] };
            var hits = catalog.Search(query, orientation, minWidth, tag, int.MaxValue);
            return new AssetOperationResult
            {
                Success = true,
                CatalogPath = catalog.File,
                Assets = hits.Take(limit).Select(hit => hit.Asset).ToList(),
                TotalCount = hits.Count,
            };
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult TagAsset(IPresentationBatch batch, string path, string? tags = null, string? description = null, string? attribution = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(path))
            return Fail("path is required.");
        if (tags is null && description is null && attribution is null)
            return Fail("Pass tags, description, or attribution.");
        return WithCatalog(catalog =>
        {
            var entry = catalog.Find(path) ?? catalog.Add(path);
            foreach (var raw in (tags ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var tag = raw.TrimStart('-').ToLowerInvariant();
                if (tag.Length == 0)
                    continue;
                if (raw.StartsWith('-'))
                    entry.Tags.Remove(tag);
                else if (!entry.Tags.Contains(tag))
                    entry.Tags.Add(tag);
            }
            if (description is not null)
                entry.Description = description.Length == 0 ? null : description;
            if (attribution is not null)
                entry.Attribution = attribution.Length == 0 ? null : attribution;
            catalog.Save();
            return new AssetOperationResult { Success = true, CatalogPath = catalog.File, Assets = [entry] };
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult Duplicates(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);
        AssetCatalog catalog;
        try
        {
            catalog = AssetCatalog.Load();
        }
        catch (InvalidDataException ex)
        {
            return Fail(ex.Message);
        }
        var groups = catalog.Duplicates().Select(group => (IReadOnlyList<string>)group.Select(entry => entry.Path).ToList()).ToList();
        var pictures = batch.Execute((ctx, ct) => ReadPictures(ctx.Presentation, null, ct, out _));
        var placed = pictures.Where(picture => picture.Source is not null)
            .GroupBy(picture => picture.Source!, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => $"{group.Key} is placed {group.Count().ToString(CultureInfo.InvariantCulture)} times (slides {string.Join(", ", group.Select(picture => picture.SlideIndex).Distinct())}); PowerPoint stores the image once.")
            .ToList();
        return new AssetOperationResult
        {
            Success = true,
            CatalogPath = catalog.File,
            DuplicateGroups = groups,
            Warnings = placed.Count > 0 ? placed : null,
        };
    }

    /// <inheritdoc/>
    public AssetOperationResult Inspect(IPresentationBatch batch, string? selector = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
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
            var pictures = ReadPictures(ctx.Presentation, parsed, ct, out var notes);
            var counts = pictures.SelectMany(picture => picture.Issues ?? []).GroupBy(issue => issue).OrderBy(group => group.Key, StringComparer.Ordinal)
                .Select(group => $"{group.Key}: {group.Count().ToString(CultureInfo.InvariantCulture)}").ToList();
            if (counts.Count > 0)
                notes.Insert(0, "Issues — " + string.Join(", ", counts) + ".");
            if (pictures.Any(picture => picture.Issues?.Contains("unknown-resolution") == true))
                notes.Add("Resolution is known for pictures placed or replaced with the asset tool, and for linked files.");
            return new AssetOperationResult { Success = true, Pictures = pictures, TotalCount = pictures.Count, Warnings = notes.Count > 0 ? notes : null };
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult Place(
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
        string? name = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        int given = (left is null ? 0 : 1) + (top is null ? 0 : 1) + (width is null ? 0 : 1) + (height is null ? 0 : 1);
        if (given is not (0 or 4))
            return Fail("Pass all of left, top, width, and height, or none of them for the default content area.");
        if (width is <= 0 || height is <= 0)
            return Fail("width and height must be greater than 0.");
        var (image, error) = ReadImage(path, fit, focalX, focalY);
        if (error is not null)
            return Fail(error);
        var entry = TryCatalogEntry(image!.Path);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? picture = null;
            try
            {
                slides = ctx.Presentation.Slides;
                if (slideIndex < 1 || slideIndex > slides.Count)
                    return Fail($"Slide index {slideIndex} is out of range (1-{slides.Count}).");
                slide = slides[slideIndex];
                shapes = slide.Shapes;
                var frame = given == 4 ? new Box(left!.Value, top!.Value, left.Value + width!.Value, top.Value + height!.Value) : ContentArea(ctx.Presentation);
                var placement = PictureFit.Compute(frame, image.Header.Aspect, fit, focalX, focalY);
                var credit = attribution ?? entry?.Attribution;
                picture = PicturePlacer.Place(shapes, new PicturePlan(image.Path, placement, null, image.Header.DisplayWidth, image.Header.DisplayHeight, fit, credit));
                if (!string.IsNullOrWhiteSpace(name))
                    picture.Name = name;
                var warnings = new List<string>();
                ApplyAccessibility(picture, decorative ? "" : altText ?? entry?.Description, decorative, null, warnings);
                if (!decorative && string.IsNullOrWhiteSpace(altText ?? entry?.Description))
                    warnings.Add("The picture has no alt text; add it with asset set-alt-text or mark it decorative.");
                AddResolutionWarning(image.Header, placement, warnings);
                return new AssetOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    ShapeId = picture.Id,
                    ShapeName = picture.Name,
                    Warnings = warnings.Count > 0 ? warnings : null,
                };
            }
            finally
            {
                if (picture is not null) ComUtilities.Release(ref picture);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult Replace(
        IPresentationBatch batch,
        string target,
        string path,
        string fit = "cover",
        float focalX = 0.5f,
        float focalY = 0.5f,
        string? altText = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var (image, error) = ReadImage(path, fit, focalX, focalY);
        if (error is not null)
            return Fail(error);
        var entry = TryCatalogEntry(image!.Path);

        return batch.Execute((ctx, ct) =>
        {
            var (item, findError) = FindSingle(ctx.Presentation, target, ct);
            if (findError is not null)
                return Fail(findError);
            if (item!.Kind != "picture" && item.HasPicture != true)
                return Fail($"{ObjectSelector.Describe(item)} is a {item.Kind}, not a picture; use asset place to add a picture.");
            if (item.GroupShapeId is not null)
                return Fail($"{ObjectSelector.Describe(item)} is inside a group; ungroup it first (shape ungroup) or place a new picture.");

            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? old = null;
            PowerPoint.Shape? picture = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(ctx.Presentation, item.SlideId)!;
                shapes = slide.Shapes;
                old = DeckShapeLocator.FindShape(slide, item.ShapeId)!;
                var warnings = new List<string>();
                if (HasAnimation(slide, item.ShapeId))
                    warnings.Add("The old picture had animations; they were not carried over (add them again with the animation tool).");
                var frame = new Box(old.Left, old.Top, old.Left + old.Width, old.Top + old.Height);
                var placement = PictureFit.Compute(frame, image.Header.Aspect, fit, focalX, focalY);
                var keptName = old.Name;
                var keptAlt = old.AlternativeText;
                var keptTags = ReadTags(old);
                picture = PicturePlacer.Place(shapes, new PicturePlan(image.Path, placement, null, image.Header.DisplayWidth, image.Header.DisplayHeight, fit, entry?.Attribution));
                MoveBelowTop(picture, old.ZOrderPosition + 1);
                old.Delete();
                ComUtilities.Release(ref old!);
                old = null;
                picture.Name = keptName;
                CopyTags(picture, keptTags, skip: [DeckRoles.ImagePixelsTag, SourceTag, DeckRoles.AttributionTag]);
                picture.AlternativeText = altText ?? (string.IsNullOrEmpty(keptAlt) ? entry?.Description ?? "" : keptAlt);
                if (string.IsNullOrWhiteSpace(picture.AlternativeText) && keptTags.GetValueOrDefault(DecorativeTag) != "1")
                    warnings.Add("The picture has no alt text; add it with asset set-alt-text.");
                if (altText is null && !string.IsNullOrEmpty(keptAlt))
                    warnings.Add("The previous alt text was kept; pass alt_text if it no longer describes the new image.");
                AddResolutionWarning(image.Header, placement, warnings);
                return new AssetOperationResult
                {
                    Success = true,
                    SlideIndex = item.SlideIndex,
                    ShapeId = picture.Id,
                    ShapeName = picture.Name,
                    Warnings = warnings.Count > 0 ? warnings : null,
                };
            }
            finally
            {
                if (picture is not null) ComUtilities.Release(ref picture);
                if (old is not null) ComUtilities.Release(ref old);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult SetAltText(IPresentationBatch batch, string target, string? altText = null, bool? decorative = null, string? attribution = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (altText is null && decorative is null && attribution is null)
            return Fail("Pass alt_text, decorative, or attribution.");
        if (decorative == true && !string.IsNullOrWhiteSpace(altText))
            return Fail("A decorative object has no alt text; pass either alt_text or decorative=true.");
        return batch.Execute((ctx, ct) =>
        {
            var (item, findError) = FindSingle(ctx.Presentation, target, ct);
            if (findError is not null)
                return Fail(findError);
            PowerPoint.Slide? slide = null;
            PowerPoint.Shape? shape = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(ctx.Presentation, item!.SlideId)!;
                shape = DeckShapeLocator.FindShape(slide, item.ShapeId)!;
                var warnings = new List<string>();
                ApplyAccessibility(shape, decorative == true ? "" : altText, decorative, attribution, warnings);
                return new AssetOperationResult { Success = true, SlideIndex = item.SlideIndex, ShapeId = item.ShapeId, ShapeName = item.Name, Warnings = warnings.Count > 0 ? warnings : null };
            }
            finally
            {
                if (shape is not null) ComUtilities.Release(ref shape);
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult EmbedLinked(IPresentationBatch batch, string? selector = null, bool dryRun = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
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
            var linked = ReadPictures(ctx.Presentation, parsed, ct, out _).Where(picture => picture.Linked).ToList();
            var changes = new List<string>();
            var warnings = new List<string>();
            foreach (var picture in linked)
            {
                if (picture.LinkExists != true)
                {
                    warnings.Add($"'{picture.Name}' (slide {picture.SlideIndex}) links to a missing file {picture.LinkSource}; use asset fix-links first.");
                    continue;
                }
                changes.Add($"slide {picture.SlideIndex} '{picture.Name}': embed {picture.LinkSource}");
                if (!dryRun)
                    EmbedOne(ctx.Presentation, picture, warnings);
            }
            if (linked.Count == 0)
                warnings.Add("No linked pictures found.");
            if (dryRun && changes.Count > 0)
                warnings.Add("Dry run: nothing was changed. Pass dry_run=false to embed.");
            return new AssetOperationResult { Success = true, Changes = changes, TotalCount = changes.Count, Warnings = warnings.Count > 0 ? warnings : null };
        });
    }

    /// <inheritdoc/>
    public AssetOperationResult FixLinks(IPresentationBatch batch, string searchFolder, string? selector = null, bool dryRun = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(searchFolder) || !Directory.Exists(searchFolder))
            return Fail($"The folder '{searchFolder}' does not exist.");
        ObjectSelector? parsed;
        try
        {
            parsed = string.IsNullOrWhiteSpace(selector) ? null : ObjectSelector.Parse(selector);
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }
        var index = Directory.EnumerateFiles(Path.GetFullPath(searchFolder), "*", new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true })
            .Take(100_000)
            .GroupBy(file => Path.GetFileName(file), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);

        return batch.Execute((ctx, ct) =>
        {
            var broken = ReadPictures(ctx.Presentation, parsed, ct, out _).Where(picture => picture.Linked && picture.LinkExists == false).ToList();
            var changes = new List<string>();
            var warnings = new List<string>();
            foreach (var picture in broken)
            {
                var fileName = Path.GetFileName(picture.LinkSource ?? "");
                var candidates = index.GetValueOrDefault(fileName) ?? [];
                if (candidates.Count != 1)
                {
                    warnings.Add(candidates.Count == 0
                        ? $"'{picture.Name}' (slide {picture.SlideIndex}): no file named {fileName} under {searchFolder}."
                        : $"'{picture.Name}' (slide {picture.SlideIndex}): {candidates.Count} files named {fileName}; relink it by hand: {string.Join("; ", candidates.Take(5))}.");
                    continue;
                }
                changes.Add($"slide {picture.SlideIndex} '{picture.Name}': {picture.LinkSource} → {candidates[0]}");
                if (!dryRun)
                    Relink(ctx.Presentation, picture, candidates[0]);
            }
            if (broken.Count == 0)
                warnings.Add("No broken picture links found.");
            if (dryRun && changes.Count > 0)
                warnings.Add("Dry run: nothing was changed. Pass dry_run=false to relink.");
            return new AssetOperationResult { Success = true, Changes = changes, TotalCount = changes.Count, Warnings = warnings.Count > 0 ? warnings : null };
        });
    }

    private sealed record LocalImage(string Path, ImageHeaderInfo Header);

    private static (LocalImage? Image, string? Error) ReadImage(string path, string fit, float focalX, float focalY)
    {
        if (fit is not ("contain" or "cover"))
            return (null, "fit must be contain or cover (pictures are never stretched).");
        if (focalX is < 0 or > 1 || focalY is < 0 or > 1)
            return (null, "focal_x and focal_y must be between 0 and 1.");
        if (string.IsNullOrWhiteSpace(path))
            return (null, "path is required (a local image file).");
        if (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || path.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            return (null, "Only local files are supported; nothing is downloaded. Save the image locally first.");
        string full;
        try
        {
            full = System.IO.Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return (null, $"Invalid path '{path}': {ex.Message}");
        }
        if (!File.Exists(full))
            return (null, $"The image '{full}' does not exist on this computer.");
        var header = ImageHeaderReader.ReadFile(full);
        if (header is null || header.Aspect <= 0)
            return (null, $"'{full}' is not a supported image or has no size (PNG, JPEG, GIF, BMP, TIFF, WebP, SVG with width/height or viewBox, EMF, WMF).");
        return (new LocalImage(full, header), null);
    }

    private static AssetEntry? TryCatalogEntry(string path)
    {
        try
        {
            return AssetCatalog.Load().Find(path);
        }
        catch (InvalidDataException)
        {
            return null;
        }
    }

    private static AssetOperationResult WithCatalog(Func<AssetCatalog, AssetOperationResult> action)
    {
        try
        {
            return action(AssetCatalog.Load());
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }
        catch (InvalidDataException ex)
        {
            return Fail(ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fail($"The asset catalog could not be read or written: {ex.Message}");
        }
    }

    private static Box ContentArea(PowerPoint.Presentation presentation)
    {
        PowerPoint.PageSetup? setup = null;
        try
        {
            setup = presentation.PageSetup;
            var profile = DesignProfiles.ResolveBuiltIn("default");
            float left = profile.Margins["left"], right = setup.SlideWidth - profile.Margins["right"];
            float top = profile.Margins["top"] + profile.Areas["title"], bottom = setup.SlideHeight - profile.Margins["bottom"] - profile.Areas["footer"];
            return new Box(left, top, right, bottom);
        }
        finally
        {
            if (setup is not null) ComUtilities.Release(ref setup);
        }
    }

    private static void AddResolutionWarning(ImageHeaderInfo header, PicturePlacement placement, List<string> warnings)
    {
        if (header.IsVector || placement.PictureWidth <= 0)
            return;
        int ppi = (int)Math.Round(header.DisplayWidth / (placement.PictureWidth / 72.0));
        if (ppi < PictureQuality.LowPpi)
            warnings.Add($"The image is shown at about {ppi} pixels per inch and will look blurry; use a larger image (at least {PictureQuality.SoftPpi} PPI).");
        else if (ppi < PictureQuality.SoftPpi)
            warnings.Add($"The image is shown at about {ppi} pixels per inch and may look soft when projected.");
    }

    private static (DeckObjectInfo? Item, string? Error) FindSingle(PowerPoint.Presentation presentation, string target, CancellationToken cancellationToken)
    {
        ObjectSelector parsed;
        try
        {
            parsed = ObjectSelector.Parse(target ?? "");
        }
        catch (ArgumentException ex)
        {
            return (null, ex.Message);
        }
        var deck = DeckSnapshotReader.ReadAll(presentation, detailed: true, cancellationToken);
        var matches = DeckCommands.Select(deck, parsed);
        try
        {
            return (ObjectSelector.RequireSingle(target ?? "", matches, deck.Sum(entry => entry.Objects.Count)), null);
        }
        catch (ArgumentException ex)
        {
            return (null, ex.Message);
        }
    }

    private static List<DeckPictureInfo> ReadPictures(PowerPoint.Presentation presentation, ObjectSelector? selector, CancellationToken cancellationToken, out List<string> notes)
    {
        notes = [];
        var deck = DeckSnapshotReader.ReadAll(presentation, detailed: true, cancellationToken);
        var items = DeckCommands.Select(deck, selector).Where(item => item.Kind == "picture" || item.HasPicture == true).ToList();
        var result = new List<DeckPictureInfo>();
        foreach (var bySlide in items.GroupBy(item => item.SlideId))
        {
            PowerPoint.Slide? slide = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(presentation, bySlide.Key);
                if (slide is null)
                    continue;
                foreach (var item in bySlide)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var (pictureWidth, pictureHeight) = ReadPictureSize(slide, item.ShapeId) ?? (item.Width, item.Height);
                    result.Add(PictureQuality.Describe(item, pictureWidth, pictureHeight));
                }
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        }
        return result;
    }

    private static (float Width, float Height)? ReadPictureSize(PowerPoint.Slide slide, int shapeId)
    {
        PowerPoint.Shape? shape = null;
        PowerPoint.PictureFormat? format = null;
        Office.Crop? crop = null;
        try
        {
            shape = DeckShapeLocator.FindShape(slide, shapeId);
            if (shape is null)
                return null;
            format = shape.PictureFormat;
            crop = format.Crop;
            return (crop.PictureWidth, crop.PictureHeight);
        }
        catch (COMException)
        {
            // Picture placeholders without a picture and some linked formats have no crop data.
            return null;
        }
        finally
        {
            if (crop is not null) ComUtilities.Release(ref crop);
            if (format is not null) ComUtilities.Release(ref format);
            if (shape is not null) ComUtilities.Release(ref shape);
        }
    }

    private static void ApplyAccessibility(PowerPoint.Shape shape, string? altText, bool? decorative, string? attribution, List<string> warnings)
    {
        if (altText is not null)
            shape.AlternativeText = altText;
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            if (decorative is { } isDecorative)
            {
                if (isDecorative)
                    tags.Add(DecorativeTag, "1");
                else
                    tags.Delete(DecorativeTag);
                try
                {
                    // PIA gap: Shape.Decorative (Office 2019+) is not in the PowerPoint 15 PIA; set it late-bound when the installed PowerPoint has it.
                    ((dynamic)shape).Decorative = isDecorative ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                }
                catch (Exception ex) when (ex is RuntimeBinderException or COMException)
                {
                    warnings.Add("This PowerPoint version has no Decorative flag; the PPTMCP_DECORATIVE tag records it instead.");
                }
            }
            if (attribution is not null)
            {
                if (attribution.Length == 0)
                    tags.Delete(DeckRoles.AttributionTag);
                else
                    tags.Add(DeckRoles.AttributionTag, attribution);
            }
        }
        finally
        {
            if (tags is not null) ComUtilities.Release(ref tags);
        }
    }

    private static Dictionary<string, string> ReadTags(PowerPoint.Shape shape)
    {
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            return DeckSnapshotReader.ReadTags(tags);
        }
        finally
        {
            if (tags is not null) ComUtilities.Release(ref tags);
        }
    }

    private static void CopyTags(PowerPoint.Shape shape, Dictionary<string, string> values, IReadOnlyCollection<string> skip)
    {
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            foreach (var (name, value) in values)
            {
                if (!skip.Contains(name, StringComparer.OrdinalIgnoreCase))
                    tags.Add(name, value);
            }
        }
        finally
        {
            if (tags is not null) ComUtilities.Release(ref tags);
        }
    }

    /// <summary>Sends a newly added (top-most) shape backward until it sits at <paramref name="position"/>.</summary>
    private static void MoveBelowTop(PowerPoint.Shape shape, int position)
    {
        for (int guard = 0; guard < 5000 && shape.ZOrderPosition > position; guard++)
            shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
    }

    private static bool HasAnimation(PowerPoint.Slide slide, int shapeId)
    {
        PowerPoint.TimeLine? timeline = null;
        PowerPoint.Sequence? sequence = null;
        try
        {
            timeline = slide.TimeLine;
            sequence = timeline.MainSequence;
            for (int index = 1; index <= sequence.Count; index++)
            {
                PowerPoint.Effect? effect = null;
                PowerPoint.Shape? animated = null;
                try
                {
                    effect = sequence[index];
                    animated = effect.Shape;
                    if (animated.Id == shapeId)
                        return true;
                }
                finally
                {
                    if (animated is not null) ComUtilities.Release(ref animated);
                    if (effect is not null) ComUtilities.Release(ref effect);
                }
            }
            return false;
        }
        finally
        {
            if (sequence is not null) ComUtilities.Release(ref sequence);
            if (timeline is not null) ComUtilities.Release(ref timeline);
        }
    }

    private static void EmbedOne(PowerPoint.Presentation presentation, DeckPictureInfo picture, List<string> warnings)
    {
        PowerPoint.Slide? slide = null;
        PowerPoint.Shapes? shapes = null;
        PowerPoint.Shape? old = null;
        PowerPoint.Shape? embedded = null;
        PowerPoint.PictureFormat? oldFormat = null, newFormat = null;
        Office.Crop? oldCrop = null, newCrop = null;
        try
        {
            slide = DeckShapeLocator.FindSlide(presentation, picture.SlideId);
            old = slide is null ? null : DeckShapeLocator.FindShape(slide, picture.ShapeId);
            if (slide is null || old is null)
                return;
            if (HasAnimation(slide, picture.ShapeId))
                warnings.Add($"'{picture.Name}' (slide {picture.SlideIndex}) had animations that were not carried over.");
            shapes = slide.Shapes;
            oldFormat = old.PictureFormat;
            oldCrop = oldFormat.Crop;
            embedded = shapes.AddPicture(picture.LinkSource!, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, oldCrop.ShapeLeft, oldCrop.ShapeTop, oldCrop.PictureWidth, oldCrop.PictureHeight);
            newFormat = embedded.PictureFormat;
            newCrop = newFormat.Crop;
            newCrop.PictureWidth = oldCrop.PictureWidth;
            newCrop.PictureHeight = oldCrop.PictureHeight;
            newCrop.ShapeLeft = oldCrop.ShapeLeft;
            newCrop.ShapeTop = oldCrop.ShapeTop;
            newCrop.ShapeWidth = oldCrop.ShapeWidth;
            newCrop.ShapeHeight = oldCrop.ShapeHeight;
            newCrop.PictureOffsetX = oldCrop.PictureOffsetX;
            newCrop.PictureOffsetY = oldCrop.PictureOffsetY;
            embedded.Rotation = old.Rotation;
            var keptTags = ReadTags(old);
            var name = old.Name;
            var alt = old.AlternativeText;
            MoveBelowTop(embedded, old.ZOrderPosition + 1);
            old.Delete();
            embedded.Name = name;
            embedded.AlternativeText = alt;
            CopyTags(embedded, keptTags, skip: []);
            var header = ImageHeaderReader.ReadFile(picture.LinkSource!);
            var tags = embedded.Tags;
            try
            {
                tags.Add(SourceTag, picture.LinkSource!);
                if (header is not null)
                    tags.Add(DeckRoles.ImagePixelsTag, $"{header.DisplayWidth.ToString(CultureInfo.InvariantCulture)}x{header.DisplayHeight.ToString(CultureInfo.InvariantCulture)}");
            }
            finally
            {
                ComUtilities.Release(ref tags);
            }
        }
        finally
        {
            if (newCrop is not null) ComUtilities.Release(ref newCrop);
            if (newFormat is not null) ComUtilities.Release(ref newFormat);
            if (oldCrop is not null) ComUtilities.Release(ref oldCrop);
            if (oldFormat is not null) ComUtilities.Release(ref oldFormat);
            if (embedded is not null) ComUtilities.Release(ref embedded);
            if (old is not null) ComUtilities.Release(ref old);
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (slide is not null) ComUtilities.Release(ref slide);
        }
    }

    private static void Relink(PowerPoint.Presentation presentation, DeckPictureInfo picture, string path)
    {
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.LinkFormat? link = null;
        try
        {
            slide = DeckShapeLocator.FindSlide(presentation, picture.SlideId);
            shape = slide is null ? null : DeckShapeLocator.FindShape(slide, picture.ShapeId);
            if (shape is null)
                return;
            link = shape.LinkFormat;
            link.SourceFullName = path;
        }
        finally
        {
            if (link is not null) ComUtilities.Release(ref link);
            if (shape is not null) ComUtilities.Release(ref shape);
            if (slide is not null) ComUtilities.Release(ref slide);
        }
    }

    private static AssetOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
