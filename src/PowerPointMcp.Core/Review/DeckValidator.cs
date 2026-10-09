using System.Globalization;
using Sbroenne.PowerPointMcp.Core.Deck;
using static Sbroenne.PowerPointMcp.Core.Deck.DeckGeometry;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Result of <see cref="DeckValidator.Validate"/>.</summary>
public sealed class ValidationReport
{
    /// <summary>Findings, most severe first, then by slide and shape.</summary>
    public required IReadOnlyList<DeckFinding> Findings { get; init; }

    /// <summary>Rule codes that ran.</summary>
    public required IReadOnlyList<string> RulesRun { get; init; }

    /// <summary>Rules that did not run and why.</summary>
    public required IReadOnlyList<string> NotChecked { get; init; }
}

/// <summary>
/// Deterministic and heuristic checks over inspection snapshots. Pure: no COM access, so every
/// rule is unit-testable. Objects inherited from the slide master and layout are not part of the
/// slide's shapes and are not checked individually.
/// </summary>
public static class DeckValidator
{
    private const float Tolerance = 1f;
    private const float BackdropCoverage = 0.9f;
    private const float MinOverlap = 2f;
    private const float DefaultFrameMargins = 7.2f;
    private const float SuggestedGap = 12f;
    private const float LargeTextSize = 18f;
    private const double DistortionTolerance = 0.02;

    private static readonly HashSet<string> FooterRoles = new(StringComparer.Ordinal)
    {
        "footer", "slide-number", "date", "header", "source",
    };

    private static readonly HashSet<string> DescribedKinds = new(StringComparer.Ordinal)
    {
        "picture", "chart", "smart-art", "media",
    };

    /// <summary>Every rule with its certainty and detection method.</summary>
    public static readonly IReadOnlyList<ValidationRule> Rules =
    [
        Rule("off-slide", "deterministic", "geometry: rotated object bounds vs slide size", "Objects partly or entirely outside the slide.", true, true),
        Rule("outside-safe-area", "deterministic", "geometry: object bounds vs safe_margin", "Content closer to the slide edge than the safe margin (only with safe_margin).", true, false),
        Rule("text-overflow", "deterministic", "measured: PowerPoint TextRange bounds vs shape bounds", "Text that does not fit its box.", true, true),
        Rule("small-text", "deterministic", "font: smallest run size vs threshold", "Text below the readable size.", true, false),
        Rule("text-overlap", "deterministic", "geometry: intersection of two text-bearing objects", "Two text objects overlap.", true, true),
        Rule("partial-overlap", "heuristic", "geometry: partial intersection, not containment", "Objects partly overlap, which is often accidental.", false, false),
        Rule("near-misaligned", "heuristic", "geometry: edges within the alignment band", "Edges that almost line up.", true, true),
        Rule("inconsistent-title-position", "heuristic", "geometry: title position vs the most common position for the same layout", "Titles that jump between slides of the same layout.", true, true),
        Rule("inconsistent-margins", "heuristic", "geometry: leftmost content edge vs the deck's most common edge", "Slides whose content starts at a different left margin.", false, false),
        Rule("image-distortion", "deterministic", "image header or PPTMCP_IMG_PX tag vs displayed aspect ratio", "Pictures stretched out of their source aspect ratio.", true, true),
        Rule("missing-asset", "deterministic", "file system: linked source file exists", "Linked pictures or media whose file is missing.", false, false),
        Rule("missing-font", "deterministic", "fonts: run font names vs installed font families", "Fonts used in text that are not installed on this computer.", false, false),
        Rule("empty-placeholder", "deterministic", "placeholder without text, picture, chart, or table", "Empty placeholders that show prompt text while editing.", false, false),
        Rule("footer-inconsistent", "heuristic", "deck: footer, date, and slide-number objects across content slides", "Footers missing on some slides or with different text.", false, false),
        Rule("duplicate-footer", "heuristic", "slide: several footer objects with the same text", "A footer repeated on one slide.", false, false),
        Rule("missing-title", "deterministic", "accessibility: slide has no title placeholder text", "Slides without a title (screen readers announce titles).", false, false),
        Rule("duplicate-title", "heuristic", "accessibility: identical titles on several slides", "Slides sharing a title.", false, false),
        Rule("missing-alt-text", "deterministic", "accessibility: pictures, charts, SmartArt, and media without alternative text", "Visuals without alternative text.", false, false),
        Rule("low-contrast", "heuristic", "color: first-character text color vs fill, container, or slide background (WCAG ratio)", "Text with low contrast against its background.", false, false),
        Rule("duplicate-app-id", "deterministic", "tags: PPTMCP_ID used by several slides or objects", "Persistent ids duplicated by copying.", true, true),
    ];

    /// <summary>Codes repair includes when the caller names none.</summary>
    public static IReadOnlySet<string> DefaultRepairCodes { get; } =
        Rules.Where(rule => rule.RepairedByDefault).Select(rule => rule.Code).ToHashSet(StringComparer.Ordinal);

    /// <summary>Validates a deck snapshot read in detailed mode.</summary>
    public static ValidationReport Validate(
        float slideWidth,
        float slideHeight,
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck,
        ValidationOptions options)
    {
        ArgumentNullException.ThrowIfNull(deck);
        ArgumentNullException.ThrowIfNull(options);

        var context = new Context(slideWidth, slideHeight, options);
        foreach (var (slide, objects) in deck)
        {
            ValidateSlide(context, slide, objects);
        }
        ValidateDeck(context, deck);

        var notChecked = new List<string>
        {
            "Objects inherited from the slide master and layout are not checked individually.",
        };
        if (options.InstalledFonts is null)
            notChecked.Add("missing-font: installed fonts are only listed on Windows.");
        if (options.SafeMargin is null)
            notChecked.Add("outside-safe-area: pass safe_margin to check it.");
        if (!options.IncludeHeuristic)
            notChecked.Add("Heuristic rules were skipped (include_heuristic=false).");

        var ran = Rules.Where(rule => context.Runs(rule.Code)).Select(rule => rule.Code).ToList();
        return new ValidationReport
        {
            Findings = context.Findings
                .OrderBy(finding => SeverityRank(finding.Severity))
                .ThenBy(finding => finding.SlideIndex)
                .ThenBy(finding => finding.ShapeIds is { Count: > 0 } ids ? ids[0] : 0)
                .ThenBy(finding => finding.Code, StringComparer.Ordinal)
                .ToList(),
            RulesRun = ran,
            NotChecked = notChecked,
        };
    }

    private static void ValidateSlide(Context context, DeckSlideInfo slide, IReadOnlyList<DeckObjectInfo> objects)
    {
        var topLevel = objects.Where(item => item.ShapeIndex is not null && item.Visible).ToList();
        var visible = objects.Where(item => item.Visible && IsVisibleThroughGroups(item, objects)).ToList();

        foreach (var item in topLevel)
        {
            CheckSlideBounds(context, slide, item);
            CheckSafeArea(context, slide, item);
        }

        foreach (var item in visible)
        {
            CheckTextOverflow(context, slide, item);
            CheckFontSize(context, slide, item);
            CheckEmptyPlaceholder(context, slide, item);
            CheckImageDistortion(context, slide, item);
            CheckMissingAsset(context, slide, item);
            CheckAltText(context, slide, item);
            CheckContrast(context, slide, item, visible);
        }

        CheckMissingFonts(context, slide, visible);
        CheckMissingTitle(context, slide, objects);
        CheckDuplicateFooter(context, slide, topLevel);

        var blocks = topLevel.Where(item => IsContentBlock(context, item)).OrderBy(item => item.ShapeIndex).ToList();
        for (var i = 0; i < blocks.Count; i++)
        {
            for (var j = i + 1; j < blocks.Count; j++)
            {
                CheckOverlap(context, slide, blocks[i], blocks[j]);
                CheckAlignment(context, slide, blocks[i], blocks[j]);
            }
        }
    }

    private static bool IsVisibleThroughGroups(DeckObjectInfo item, IReadOnlyList<DeckObjectInfo> objects)
    {
        var current = item;
        while (current.GroupShapeId is { } parentId)
        {
            var parent = objects.FirstOrDefault(candidate => candidate.ShapeId == parentId);
            if (parent is null)
                return true;
            if (!parent.Visible)
                return false;
            current = parent;
        }
        return true;
    }

    private static void CheckSlideBounds(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (!context.Runs("off-slide") || (item.Width <= 0 && item.Height <= 0))
            return;

        var box = BoundsOf(item);
        var evidence = Evidence(("bounds", BoxText(box)), ("slide", $"{P(context.Width)} x {P(context.Height)}"));
        if (box.Right <= Tolerance || box.Left >= context.Width - Tolerance || box.Bottom <= Tolerance || box.Top >= context.Height - Tolerance)
        {
            context.Add("off-slide", "error", slide, [item],
                $"{Describe(item)} is entirely outside the {P(context.Width)} x {P(context.Height)} pt slide and is not shown in the slide show.",
                "It may be parked there on purpose. Otherwise move it onto the slide with shape set-position; repair does not move objects that are entirely off the slide.",
                evidence, autoRepairable: false);
            return;
        }

        if (IsBackdrop(context, item))
            return;

        if (box.Left < -Tolerance || box.Top < -Tolerance || box.Right > context.Width + Tolerance || box.Bottom > context.Height + Tolerance)
        {
            bool fits = box.Width <= context.Width + Tolerance && box.Height <= context.Height + Tolerance;
            context.Add("off-slide", "warning", slide, [item],
                $"{Describe(item)} extends past the slide edge: it spans {BoxText(box)} on a {P(context.Width)} x {P(context.Height)} pt slide.",
                fits
                    ? "Move it fully onto the slide (review repair does this, or shape set-position)."
                    : $"It is larger than the slide; resize it to fit within {P(context.Width)} x {P(context.Height)} pt (shape set-size).",
                evidence, autoRepairable: fits);
        }
    }

    private static void CheckSafeArea(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (context.Options.SafeMargin is not { } margin || !context.Runs("outside-safe-area") || IsBackdrop(context, item) ||
            item.Kind is "line" or "connector" || IsFooterLike(item))
        {
            return;
        }

        var box = BoundsOf(item);
        if (box.Right <= 0 || box.Left >= context.Width || box.Bottom <= 0 || box.Top >= context.Height)
            return; // Reported as off-slide.
        if (box.Left < margin - Tolerance || box.Top < margin - Tolerance ||
            box.Right > context.Width - margin + Tolerance || box.Bottom > context.Height - margin + Tolerance)
        {
            bool fits = box.Width <= context.Width - (2 * margin) && box.Height <= context.Height - (2 * margin);
            context.Add("outside-safe-area", "info", slide, [item],
                $"{Describe(item)} is within {P(margin)} pt of the slide edge ({BoxText(box)}).",
                fits ? "Move it inside the safe area (review repair with codes=outside-safe-area)." : "Shrink it to fit inside the safe area.",
                Evidence(("bounds", BoxText(box)), ("safeMargin", P(margin))), autoRepairable: fits);
        }
    }

    private static void CheckTextOverflow(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (!context.Runs("text-overflow") || item.Text is null || item.TextBounds is not { Count: 4 } bounds ||
            !IsUpright(item.Rotation) || item.TableRows is not null)
        {
            return;
        }

        var (textLeft, textTop, textWidth, textHeight) = (bounds[0], bounds[1], bounds[2], bounds[3]);
        var shrinking = item.AutoSize == "shrink-text-on-overflow" ? " PowerPoint is already shrinking this text to fit." : "";
        var margins = VerticalMargins(item);
        if (textTop + textHeight > item.Top + item.Height + Tolerance || textTop < item.Top - Tolerance)
        {
            context.Add("text-overflow", "error", slide, [item],
                $"Text in {Describe(item)} overflows its box vertically: the text is {P(textHeight)} pt tall but the box is {P(item.Height)} pt tall.{shrinking}",
                $"Make the box at least {P(MathF.Ceiling(textHeight + margins))} pt tall, scale the font down (not below the minimum), or split the content. Review repair tries these in order of fit_policy.",
                Evidence(("textHeight", P(textHeight)), ("boxHeight", P(item.Height)), ("textTop", P(textTop)), ("boxTop", P(item.Top)),
                    ("minFontSize", item.MinFontSize is { } min ? P(min) : "unknown")),
                autoRepairable: true);
        }

        if (item.WordWrap != true && (textLeft + textWidth > item.Left + item.Width + Tolerance || textLeft < item.Left - Tolerance))
        {
            context.Add("text-overflow", "error", slide, [item],
                $"Text in {Describe(item)} does not wrap and overflows horizontally: the text is {P(textWidth)} pt wide but the box is {P(item.Width)} pt wide.",
                $"Turn on word wrap, widen the box to at least {P(MathF.Ceiling(textWidth + HorizontalMargins(item)))} pt, or scale the font down.",
                Evidence(("textWidth", P(textWidth)), ("boxWidth", P(item.Width)), ("wordWrap", "false")),
                autoRepairable: true);
        }
    }

    private static void CheckFontSize(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (!context.Runs("small-text") || item.Text is null || item.MinFontSize is not { } size || size <= 0)
            return;
        var threshold = IsFooterLike(item) ? context.Options.MinFootnoteFontSize : context.Options.MinFontSize;
        if (size < threshold - 0.01f)
        {
            context.Add("small-text", "warning", slide, [item],
                $"{Describe(item)} uses {P(size)} pt text, below the {P(threshold)} pt minimum.",
                $"Scale the text up to at least {P(threshold)} pt (review repair with codes=small-text scales runs proportionally), or move the detail into speaker notes.",
                Evidence(("minFontSize", P(size)), ("threshold", P(threshold))), autoRepairable: true);
        }
    }

    private static void CheckEmptyPlaceholder(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (context.Runs("empty-placeholder") && IsEmptyPlaceholder(item))
        {
            context.Add("empty-placeholder", "warning", slide, [item],
                $"Placeholder {Describe(item)} is empty; it shows prompt text while editing and reserves space.",
                "Fill it (shape set-placeholder-text or set-placeholder-image) or delete it with shape delete. Repair never deletes objects.",
                Evidence(("placeholderType", item.PlaceholderType ?? "")), autoRepairable: false);
        }
    }

    private static void CheckImageDistortion(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (!context.Runs("image-distortion") || item.Kind != "picture" || item.SourcePixelSize is not { Count: 2 } pixels ||
            item.Width <= 0 || item.Height <= 0)
        {
            return;
        }

        var crop = item.Crop ?? [0, 0, 0, 0];
        bool cropped = crop.Any(value => Math.Abs(value) > 0.1f);
        // Crop values are in points of the picture's original size; assume 96 dpi when cropped.
        double sourceWidth = pixels[0] * 0.75;
        double sourceHeight = pixels[1] * 0.75;
        double visibleWidth = sourceWidth - crop[0] - crop[2];
        double visibleHeight = sourceHeight - crop[1] - crop[3];
        if (visibleWidth <= 0 || visibleHeight <= 0)
            return;
        double expected = visibleWidth / visibleHeight;
        double actual = item.Width / (double)item.Height;
        double difference = Math.Abs(actual - expected) / expected;
        if (difference <= DistortionTolerance)
            return;

        context.Add("image-distortion", "warning", slide, [item],
            $"Picture {Describe(item)} is stretched: it is shown at aspect {actual:0.###} but its {(cropped ? "cropped " : "")}source is {expected:0.###} ({pixels[0]}x{pixels[1]} px), a {difference:P0} distortion.",
            "Restore the aspect ratio inside the current frame (review repair does this, keeping the center), or crop instead of stretching (asset place with fit=cover).",
            Evidence(("displayedAspect", actual.ToString("0.###", CultureInfo.InvariantCulture)), ("sourceAspect", expected.ToString("0.###", CultureInfo.InvariantCulture)),
                ("sourcePixels", $"{pixels[0]}x{pixels[1]}"), ("cropAssumes96Dpi", cropped ? "true" : "false")),
            autoRepairable: true, certainty: cropped ? "heuristic" : null);
    }

    private static void CheckMissingAsset(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (context.Runs("missing-asset") && item.LinkSourceExists == false)
        {
            context.Add("missing-asset", "error", slide, [item],
                $"{Describe(item)} links to a file that does not exist on this computer: {item.LinkSource}.",
                "Restore the file at that path, relink it (shape update-link after restoring), or replace the picture with asset replace.",
                Evidence(("linkSource", item.LinkSource ?? "")), autoRepairable: false);
        }
    }

    private static void CheckAltText(Context context, DeckSlideInfo slide, DeckObjectInfo item)
    {
        if (!context.Runs("missing-alt-text") || !DescribedKinds.Contains(item.Kind) || !string.IsNullOrWhiteSpace(item.AltText) ||
            item.Tags?.GetValueOrDefault("PPTMCP_DECORATIVE") == "1")
        {
            return;
        }
        context.Add("missing-alt-text", "warning", slide, [item],
            $"{Describe(item)} has no alternative text.",
            "Describe it with asset set-alt-text, or mark it decorative with asset set-alt-text decorative=true.",
            null, autoRepairable: false);
    }

    private static void CheckContrast(Context context, DeckSlideInfo slide, DeckObjectInfo item, IReadOnlyList<DeckObjectInfo> visible)
    {
        if (!context.Runs("low-contrast") || item.Text is null || item.TextColor is null || item.TableRows is not null)
            return;

        var background = item.FillColor;
        var source = "fill";
        if (background is null)
        {
            var box = BoundsOf(item);
            var container = visible
                .Where(other => other.ShapeId != item.ShapeId && other.FillColor is not null && other.ZOrder < item.ZOrder &&
                    other.ShapeIndex is not null && BoundsOf(other).Contains(box))
                .OrderByDescending(other => other.ZOrder)
                .FirstOrDefault();
            (background, source) = container is not null ? (container.FillColor, $"shape {container.ShapeId}") : (slide.BackgroundColor, "slide background");
        }
        if (ColorContrast.Ratio(item.TextColor, background) is not { } ratio)
            return;

        var required = item.MinFontSize is >= LargeTextSize ? 3.0 : context.Options.MinContrast;
        if (ratio < required)
        {
            context.Add("low-contrast", "warning", slide, [item],
                $"Text in {Describe(item)} ({item.TextColor}) on {source} ({background}) has contrast {ratio:0.##}:1, below {required:0.#}:1.",
                "Change the text or background color (textframe set-font-color, shape set-fill), or apply the design profile's colors.",
                Evidence(("textColor", item.TextColor), ("background", background ?? ""), ("backgroundSource", source),
                    ("ratio", ratio.ToString("0.##", CultureInfo.InvariantCulture)), ("required", required.ToString("0.#", CultureInfo.InvariantCulture))),
                autoRepairable: false);
        }
    }

    private static void CheckMissingFonts(Context context, DeckSlideInfo slide, IReadOnlyList<DeckObjectInfo> visible)
    {
        if (!context.Runs("missing-font") || context.Options.InstalledFonts is not { } installed)
            return;

        var byFont = visible
            .SelectMany(item => (item.FontNames ?? []).Select(font => (Font: font, Item: item)))
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Font) && !entry.Font.StartsWith('+') && !installed.Contains(entry.Font))
            .GroupBy(entry => entry.Font, StringComparer.OrdinalIgnoreCase);
        foreach (var group in byFont)
        {
            var items = group.Select(entry => entry.Item).DistinctBy(item => item.ShapeId).ToList();
            context.Add("missing-font", "warning", slide, items,
                $"Font '{group.Key}' is used on this slide but is not installed on this computer; PowerPoint substitutes another font, which changes text width and wrapping.",
                "Install the font, or replace it (textframe set-font-name, or design apply-profile to normalize typography).",
                Evidence(("font", group.Key)), autoRepairable: false);
        }
    }

    private static void CheckMissingTitle(Context context, DeckSlideInfo slide, IReadOnlyList<DeckObjectInfo> objects)
    {
        if (context.Runs("missing-title") && string.IsNullOrWhiteSpace(slide.Title) && objects.Count > 0)
        {
            context.Add("missing-title", "warning", slide, [],
                $"Slide {slide.SlideIndex} has no title. Screen readers and the slide navigator use titles.",
                "Use a layout with a title placeholder and fill it, or set a title placeholder's text. A hidden title is acceptable for full-bleed visuals.",
                null, autoRepairable: false);
        }
    }

    private static void CheckDuplicateFooter(Context context, DeckSlideInfo slide, IReadOnlyList<DeckObjectInfo> topLevel)
    {
        if (!context.Runs("duplicate-footer"))
            return;
        var bottomBand = context.Height * 0.85f;
        var candidates = topLevel
            .Where(item => item.Text is { Length: > 0 } && (IsFooterLike(item) || item.Top >= bottomBand))
            .GroupBy(item => Normalize(item.Text!), StringComparer.Ordinal)
            .Where(group => group.Count() > 1 && group.Any(IsFooterLike));
        foreach (var group in candidates)
        {
            context.Add("duplicate-footer", "warning", slide, group.ToList(),
                $"The footer text \"{Shorten(group.First().Text!, 60)}\" appears {group.Count()} times on slide {slide.SlideIndex}.",
                "Keep the footer placeholder and delete the extra copy (shape delete) after confirming with a preview.",
                null, autoRepairable: false);
        }
    }

    private static void CheckOverlap(Context context, DeckSlideInfo slide, DeckObjectInfo first, DeckObjectInfo second)
    {
        if (AllowsOverlap(first) || AllowsOverlap(second))
            return;
        var a = BoundsOf(first);
        var b = BoundsOf(second);
        var overlap = a.Intersect(b);
        if (overlap.Width <= MinOverlap || overlap.Height <= MinOverlap)
            return;

        var evidence = Evidence(("overlap", $"{P(overlap.Width)} x {P(overlap.Height)}"), ("first", BoxText(a)), ("second", BoxText(b)));
        if (IsTextual(first) && IsTextual(second) && context.Runs("text-overlap"))
        {
            var (upper, lower) = a.Top <= b.Top ? (first, second) : (second, first);
            var target = BoundsOf(upper).Bottom + SuggestedGap;
            bool room = target + lower.Height <= context.Height;
            context.Add("text-overlap", "error", slide, [first, second],
                $"Text in {Describe(first)} and {Describe(second)} overlaps by {P(overlap.Width)} x {P(overlap.Height)} pt.",
                room
                    ? $"Move {Describe(lower)} down to top={P(target)} (review repair does this when it creates no new overlap), or place them side by side."
                    : "There is no room below: shorten or shrink one, place them side by side, or move content to another slide.",
                evidence, autoRepairable: room);
            return;
        }

        if (a.Contains(b) || b.Contains(a) || !context.Runs("partial-overlap"))
            return; // Containment is deliberate layering (text on a card, caption on a photo).

        context.Add("partial-overlap", "warning", slide, [first, second],
            $"{Describe(first)} and {Describe(second)} partly overlap by {P(overlap.Width)} x {P(overlap.Height)} pt.",
            "Check a preview. If the layering is intended, tag one with PPTMCP_ALLOW_OVERLAP=1; otherwise move or resize one of them.",
            evidence, autoRepairable: false);
    }

    private static void CheckAlignment(Context context, DeckSlideInfo slide, DeckObjectInfo first, DeckObjectInfo second)
    {
        if (!context.Runs("near-misaligned") || !IsUpright(first.Rotation) || !IsUpright(second.Rotation) || IsFooterLike(first) || IsFooterLike(second))
            return;
        Near("left", first.Left, second.Left);
        Near("top", first.Top, second.Top);
        Near("right", first.Left + first.Width, second.Left + second.Width);

        void Near(string edge, float firstValue, float secondValue)
        {
            var difference = MathF.Abs(firstValue - secondValue);
            if (difference <= 0.5f || difference > context.Options.AlignmentBand)
                return;
            context.Add("near-misaligned", "info", slide, [first, second],
                $"The {edge} edges of {Describe(first)} and {Describe(second)} differ by {P(difference)} pt, which looks accidental.",
                $"Align {edge} edges: set {Describe(second)} to {edge}={P(firstValue)} (review repair does this) or use shape align.",
                Evidence(("edge", edge), ("first", P(firstValue)), ("second", P(secondValue))), autoRepairable: true,
                idSuffix: edge);
        }
    }

    private static void ValidateDeck(Context context, IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        CheckTitlePositions(context, deck);
        CheckMargins(context, deck);
        CheckFooters(context, deck);
        CheckDuplicateTitles(context, deck);
        CheckDuplicateIds(context, deck);
    }

    private static void CheckTitlePositions(Context context, IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        if (!context.Runs("inconsistent-title-position"))
            return;
        var titles = deck
            .Select(entry => (entry.Slide, Title: entry.Objects.FirstOrDefault(item => item.Role == "title" && item.ShapeIndex is not null && item.Visible)))
            .Where(entry => entry.Title is not null && entry.Slide.LayoutName is not null)
            .GroupBy(entry => entry.Slide.LayoutName!, StringComparer.Ordinal)
            .Where(group => group.Count() >= 3);
        foreach (var layout in titles)
        {
            var mode = layout
                .GroupBy(entry => (MathF.Round(entry.Title!.Left), MathF.Round(entry.Title!.Top)))
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key.Item2)
                .First();
            if (mode.Count() < 2)
                continue;
            var (left, top) = (mode.First().Title!.Left, mode.First().Title!.Top);
            foreach (var (slide, title) in layout)
            {
                var dx = Math.Abs(title!.Left - left);
                var dy = Math.Abs(title.Top - top);
                if ((dx > 1f || dy > 1f) && dx < 72f && dy < 72f)
                {
                    context.Add("inconsistent-title-position", "warning", slide, [title],
                        $"The title on slide {slide.SlideIndex} sits at left={P(title.Left)} top={P(title.Top)}, but {mode.Count()} slides with layout '{layout.Key}' put it at left={P(left)} top={P(top)}.",
                        $"Move it to left={P(left)} top={P(top)} (review repair does this) so titles do not jump between slides.",
                        Evidence(("left", P(title.Left)), ("top", P(title.Top)), ("expectedLeft", P(left)), ("expectedTop", P(top)), ("layout", layout.Key)),
                        autoRepairable: true);
                }
            }
        }
    }

    private static void CheckMargins(Context context, IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        if (!context.Runs("inconsistent-margins"))
            return;
        var edges = deck
            .Select(entry => (entry.Slide, Edge: entry.Objects
                .Where(item => item.ShapeIndex is not null && item.Visible && IsTextual(item) && !IsFooterLike(item) &&
                    item.Width < context.Width * 0.9f && item.Left >= 0)
                .Select(item => (float?)item.Left)
                .Min()))
            .Where(entry => entry.Edge is not null)
            .ToList();
        if (edges.Count < 3)
            return;
        var mode = edges.GroupBy(entry => MathF.Round(entry.Edge!.Value / 2f) * 2f).OrderByDescending(group => group.Count()).First();
        if (mode.Count() < Math.Max(2, edges.Count / 2))
            return;
        foreach (var (slide, edge) in edges)
        {
            var difference = Math.Abs(edge!.Value - mode.Key);
            if (difference > context.Options.AlignmentBand && difference < 48f)
            {
                context.Add("inconsistent-margins", "info", slide, [],
                    $"Content on slide {slide.SlideIndex} starts at left={P(edge.Value)} pt; most slides ({mode.Count()} of {edges.Count}) start at {P(mode.Key)} pt.",
                    $"Shift the slide's content to start at left={P(mode.Key)} pt, or apply a composition so margins come from the layout grid.",
                    Evidence(("left", P(edge.Value)), ("deckLeft", P(mode.Key))), autoRepairable: false);
            }
        }
    }

    private static void CheckFooters(Context context, IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        if (!context.Runs("footer-inconsistent"))
            return;
        var content = deck
            .Where(entry => entry.Slide.SlideIndex > 1 && !(entry.Slide.LayoutName ?? "").Contains("Title Slide", StringComparison.OrdinalIgnoreCase) &&
                !(entry.Slide.LayoutName ?? "").Contains("Section", StringComparison.OrdinalIgnoreCase))
            .Select(entry => (entry.Slide, Footer: entry.Objects.FirstOrDefault(item => item.Role == "footer" && item.Visible && item.Text is { Length: > 0 })))
            .ToList();
        var with = content.Where(entry => entry.Footer is not null).ToList();
        if (with.Count == 0 || content.Count < 2)
            return;

        var without = content.Where(entry => entry.Footer is null).Select(entry => entry.Slide.SlideIndex).ToList();
        if (without.Count > 0)
        {
            context.AddDeck("footer-inconsistent", "info", without,
                $"{with.Count} content slides have a footer but {without.Count} do not (slides {string.Join(", ", without.Take(20))}).",
                "Turn footers on consistently (insert header and footer in PowerPoint, or a design profile with footer text).",
                Evidence(("withFooter", with.Count.ToString(CultureInfo.InvariantCulture)), ("withoutFooter", without.Count.ToString(CultureInfo.InvariantCulture))));
        }

        var variants = with.GroupBy(entry => Normalize(entry.Footer!.Text!), StringComparer.Ordinal).ToList();
        if (variants.Count > 1)
        {
            var dominant = variants.OrderByDescending(group => group.Count()).First();
            var odd = variants.Where(group => group != dominant).SelectMany(group => group).Select(entry => entry.Slide.SlideIndex).ToList();
            context.AddDeck("footer-inconsistent", "warning", odd,
                $"Footer text differs: {variants.Count} variants. Most slides use \"{Shorten(dominant.First().Footer!.Text!, 60)}\"; slides {string.Join(", ", odd.Take(20))} differ.",
                "Use one footer text across the deck (deck find-text and replace-text, or set the footer placeholders).",
                Evidence(("variants", string.Join(" || ", variants.Select(group => Shorten(group.First().Footer!.Text!, 40))))),
                idSuffix: "text");
        }
    }

    private static void CheckDuplicateTitles(Context context, IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        if (!context.Runs("duplicate-title"))
            return;
        var duplicates = deck
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Slide.Title) && !IsContinuation(entry.Slide.Title!))
            .GroupBy(entry => Normalize(entry.Slide.Title!), StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1);
        foreach (var group in duplicates)
        {
            var slides = group.Select(entry => entry.Slide.SlideIndex).ToList();
            context.AddDeck("duplicate-title", "info", slides,
                $"Slides {string.Join(", ", slides)} share the title \"{Shorten(group.First().Slide.Title!, 60)}\".",
                "Give each slide a distinct title, or mark continuations with \"(cont.)\".",
                null, idSuffix: string.Join('-', slides));
        }
    }

    private static void CheckDuplicateIds(Context context, IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        if (!context.Runs("duplicate-app-id"))
            return;
        foreach (var duplicate in AppIdPlanner.FindDuplicates(deck))
        {
            var id = duplicate[(duplicate.IndexOf(':', StringComparison.Ordinal) + 1)..];
            var slides = duplicate.StartsWith("slide:", StringComparison.Ordinal)
                ? deck.Where(entry => entry.Slide.AppId == id).Select(entry => entry.Slide.SlideIndex).ToList()
                : deck.Where(entry => entry.Objects.Any(item => item.AppId == id)).Select(entry => entry.Slide.SlideIndex).Distinct().ToList();
            context.AddDeck("duplicate-app-id", "warning", slides,
                $"The persistent id '{id}' is used more than once ({duplicate}), usually after copying a slide or object; id-based edits would be ambiguous.",
                "Run deck assign-ids (or review repair) to give later copies a ~N suffix.",
                Evidence(("id", duplicate)), idSuffix: duplicate, autoRepairable: true);
        }
    }

    internal static bool IsBackdrop(float width, float height, DeckObjectInfo item)
    {
        if (IsTextual(item))
            return false;
        var box = BoundsOf(item);
        var visibleWidth = MathF.Min(box.Right, width) - MathF.Max(box.Left, 0);
        var visibleHeight = MathF.Min(box.Bottom, height) - MathF.Max(box.Top, 0);
        return visibleWidth > 0 && visibleHeight > 0 && visibleWidth * visibleHeight >= BackdropCoverage * width * height;
    }

    private static bool IsBackdrop(Context context, DeckObjectInfo item) => IsBackdrop(context.Width, context.Height, item);

    internal static bool IsContentBlock(float width, float height, DeckObjectInfo item) =>
        item.Kind is not ("line" or "connector") && item.Width > 0 && item.Height > 0 &&
        !IsBackdrop(width, height, item) && !IsEmptyPlaceholder(item);

    private static bool IsContentBlock(Context context, DeckObjectInfo item) => IsContentBlock(context.Width, context.Height, item);

    internal static bool IsEmptyPlaceholder(DeckObjectInfo item) =>
        item.Kind == "placeholder" && item.Text is null && item.HasPicture != true && item.ChartType is null &&
        item.TableRows is null && !IsFooterLike(item);

    internal static bool IsTextual(DeckObjectInfo item) => item.Text is not null || item.TableRows is not null;

    internal static bool IsFooterLike(DeckObjectInfo item) => item.Role is { } role && FooterRoles.Contains(role);

    private static bool AllowsOverlap(DeckObjectInfo item) => item.Tags?.GetValueOrDefault(DeckRoles.AllowOverlapTag) == "1";

    private static bool IsContinuation(string title) =>
        title.Contains("(cont", StringComparison.OrdinalIgnoreCase) || title.Contains("(продолж", StringComparison.OrdinalIgnoreCase);

    internal static float VerticalMargins(DeckObjectInfo item) =>
        item.TextMargins is { Count: 4 } margins ? margins[1] + margins[3] : DefaultFrameMargins;

    internal static float HorizontalMargins(DeckObjectInfo item) =>
        item.TextMargins is { Count: 4 } margins ? margins[0] + margins[2] : 14.4f;

    internal static string Describe(DeckObjectInfo item)
    {
        var appId = item.AppId is { Length: > 0 } id ? $", appid {id}" : "";
        return $"{item.Kind} '{item.Name}' (id {item.ShapeId}{appId})";
    }

    private static string BoxText(Box box) => $"{P(box.Left)}..{P(box.Right)} x {P(box.Top)}..{P(box.Bottom)}";

    private static string Normalize(string text) => string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static string Shorten(string text, int length)
    {
        var flat = Normalize(text);
        return flat.Length <= length ? flat : flat[..length] + "…";
    }

    private static int SeverityRank(string severity) => severity switch
    {
        "error" => 0,
        "warning" => 1,
        _ => 2,
    };

    private static Dictionary<string, string> Evidence(params (string Key, string Value)[] values) =>
        values.ToDictionary(value => value.Key, value => value.Value, StringComparer.Ordinal);

    private static ValidationRule Rule(string code, string certainty, string method, string description, bool auto, bool byDefault) =>
        new() { Code = code, Certainty = certainty, Method = method, Description = description, AutoRepairable = auto, RepairedByDefault = byDefault };

    private sealed class Context(float width, float height, ValidationOptions options)
    {
        private static readonly Dictionary<string, ValidationRule> ByCode = Rules.ToDictionary(rule => rule.Code, StringComparer.Ordinal);
        private readonly HashSet<string> _ids = new(StringComparer.Ordinal);

        public float Width { get; } = width;

        public float Height { get; } = height;

        public ValidationOptions Options { get; } = options;

        public List<DeckFinding> Findings { get; } = [];

        public bool Runs(string code) =>
            (Options.Rules is null || Options.Rules.Contains(code)) &&
            (Options.IncludeHeuristic || ByCode[code].Certainty != "heuristic");

        public void Add(
            string code, string severity, DeckSlideInfo slide, IReadOnlyList<DeckObjectInfo> items, string message, string suggestion,
            IReadOnlyDictionary<string, string>? evidence, bool autoRepairable, string? certainty = null, string? idSuffix = null)
        {
            var rule = ByCode[code];
            var effectiveCertainty = certainty ?? rule.Certainty;
            if (!Options.IncludeHeuristic && effectiveCertainty == "heuristic")
                return;
            var shapeIds = items.Select(item => item.ShapeId).ToList();
            var id = $"{code}:{slide.SlideId}:{string.Join('-', shapeIds)}{(idSuffix is null ? "" : ":" + idSuffix)}";
            if (!_ids.Add(id))
                return;
            Findings.Add(new DeckFinding
            {
                Id = id,
                Code = code,
                Severity = severity,
                Certainty = effectiveCertainty,
                Method = rule.Method,
                SlideIndex = slide.SlideIndex,
                SlideId = slide.SlideId,
                ShapeIds = shapeIds.Count > 0 ? shapeIds : null,
                ShapeIndexes = items.Where(item => item.ShapeIndex is not null).Select(item => item.ShapeIndex!.Value).ToList() is { Count: > 0 } indexes ? indexes : null,
                AppIds = items.Where(item => item.AppId is not null).Select(item => item.AppId!).ToList() is { Count: > 0 } appIds ? appIds : null,
                Message = message,
                Evidence = evidence,
                Suggestion = suggestion,
                AutoRepairable = autoRepairable,
            });
        }

        public void AddDeck(
            string code, string severity, IReadOnlyList<int> slideIndexes, string message, string suggestion,
            IReadOnlyDictionary<string, string>? evidence, string? idSuffix = null, bool autoRepairable = false)
        {
            var rule = ByCode[code];
            var id = $"{code}:deck{(idSuffix is null ? "" : ":" + idSuffix)}";
            if (!_ids.Add(id))
                return;
            Findings.Add(new DeckFinding
            {
                Id = id,
                Code = code,
                Severity = severity,
                Certainty = rule.Certainty,
                Method = rule.Method,
                SlideIndex = 0,
                SlideId = 0,
                SlideIndexes = slideIndexes,
                Message = message,
                Evidence = evidence,
                Suggestion = suggestion,
                AutoRepairable = autoRepairable,
            });
        }
    }
}
