using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Slide;

/// <summary>
/// Deterministic layout checks over a slide snapshot (see <see cref="SlideShapeInfo"/>). Pure
/// geometry with no COM access, so every rule is testable without PowerPoint. Each issue carries
/// the measured values and a concrete fix so an agent can repair the slide without a vision model.
/// </summary>
public static class SlideLayoutAnalyzer
{
    /// <summary>Text smaller than this (points) is reported as small-text.</summary>
    public const float MinReadableFontSize = 10f;

    /// <summary>Geometry differences up to this many points are treated as equal.</summary>
    public const float Tolerance = 1f;

    /// <summary>Edges closer than this (but not equal) are reported as near-misaligned.</summary>
    public const float AlignmentBand = 4f;

    // Shapes covering at least this share of the slide are backdrops, not content blocks.
    private const float BackdropCoverage = 0.9f;

    // Overlaps thinner than this in either direction are touching edges, not overlaps.
    private const float MinOverlap = 2f;

    // Default PowerPoint text-frame margins (top + bottom, left + right) in points.
    private const float FrameMargins = 7.2f;

    // Gap suggested between stacked shapes when separating an overlap.
    private const float SuggestedGap = 12f;

    private static readonly HashSet<string> NonContentPlaceholders = new(StringComparer.Ordinal)
    {
        "ppPlaceholderDate", "ppPlaceholderFooter", "ppPlaceholderSlideNumber", "ppPlaceholderHeader",
    };

    /// <summary>Returns the layout issues on one slide, most severe first.</summary>
    public static IReadOnlyList<SlideLayoutIssue> Analyze(
        int slideIndex,
        float slideWidth,
        float slideHeight,
        IReadOnlyList<SlideShapeInfo> shapes)
    {
        ArgumentNullException.ThrowIfNull(shapes);

        var issues = new List<SlideLayoutIssue>();
        var visible = shapes.Where(shape => shape.Visible).ToArray();
        foreach (var shape in visible)
        {
            CheckTextOverflow(slideIndex, shape, issues);
            CheckSlideBounds(slideIndex, slideWidth, slideHeight, shape, issues);
            CheckFontSize(slideIndex, shape, issues);
            CheckEmptyPlaceholder(slideIndex, shape, issues);
        }

        var blocks = visible
            .Where(shape => IsContentBlock(shape, slideWidth, slideHeight))
            .OrderBy(shape => shape.ShapeIndex)
            .ToArray();
        for (var i = 0; i < blocks.Length; i++)
        {
            for (var j = i + 1; j < blocks.Length; j++)
            {
                CheckOverlap(slideIndex, slideHeight, blocks[i], blocks[j], issues);
                CheckAlignment(slideIndex, blocks[i], blocks[j], issues);
            }
        }

        return issues
            .OrderBy(issue => SeverityRank(issue.Severity))
            .ThenBy(issue => issue.ShapeIndexes[0])
            .ThenBy(issue => issue.Code, StringComparer.Ordinal)
            .ToArray();
    }

    private static void CheckTextOverflow(int slideIndex, SlideShapeInfo shape, List<SlideLayoutIssue> issues)
    {
        if (!shape.HasText || !IsUpright(shape) ||
            shape.TextLeft is not { } textLeft || shape.TextTop is not { } textTop ||
            shape.TextWidth is not { } textWidth || shape.TextHeight is not { } textHeight)
        {
            return;
        }

        var shrinking = shape.AutoSize == "shrink-text-on-overflow" ? " PowerPoint is already shrinking this text." : "";
        if (textTop + textHeight > shape.Top + shape.Height + Tolerance || textTop < shape.Top - Tolerance)
        {
            var needed = MathF.Ceiling(textHeight + FrameMargins);
            issues.Add(Issue("text-overflow", "error", slideIndex, [shape.ShapeIndex],
                $"Text in shape {shape.ShapeIndex} ({Describe(shape)}) overflows its box vertically: the text is {P(textHeight)} pt tall but the shape is {P(shape.Height)} pt tall.{shrinking}",
                $"Set the shape height to at least {P(needed)} pt (shape set-size), shorten the text, or lower the font size{LargestFont(shape)}."));
        }

        if (textLeft + textWidth > shape.Left + shape.Width + Tolerance || textLeft < shape.Left - Tolerance)
        {
            var needed = MathF.Ceiling(textWidth + FrameMargins);
            issues.Add(Issue("text-overflow", "error", slideIndex, [shape.ShapeIndex],
                $"Text in shape {shape.ShapeIndex} ({Describe(shape)}) overflows its box horizontally: the text is {P(textWidth)} pt wide but the shape is {P(shape.Width)} pt wide.",
                $"Set the shape width to at least {P(needed)} pt (shape set-size), add line breaks, or shorten the text."));
        }
    }

    private static void CheckSlideBounds(
        int slideIndex, float slideWidth, float slideHeight, SlideShapeInfo shape, List<SlideLayoutIssue> issues)
    {
        var box = BoundsOf(shape);
        if (box.Right <= Tolerance || box.Left >= slideWidth - Tolerance ||
            box.Bottom <= Tolerance || box.Top >= slideHeight - Tolerance)
        {
            issues.Add(Issue("off-slide", "error", slideIndex, [shape.ShapeIndex],
                $"Shape {shape.ShapeIndex} ({Describe(shape)}) is entirely outside the {P(slideWidth)} x {P(slideHeight)} pt slide and will not be seen.",
                $"Move it onto the slide (shape set-position), for example left={P(Clamp(shape.Left, 0, slideWidth - shape.Width))} top={P(Clamp(shape.Top, 0, slideHeight - shape.Height))}, or delete it."));
            return;
        }

        if (IsBackdrop(shape, slideWidth, slideHeight))
            return;

        if (box.Left < -Tolerance || box.Top < -Tolerance ||
            box.Right > slideWidth + Tolerance || box.Bottom > slideHeight + Tolerance)
        {
            var fix = shape.Width > slideWidth || shape.Height > slideHeight
                ? $"Resize it to fit within {P(slideWidth)} x {P(slideHeight)} pt (shape set-size), then move it onto the slide."
                : $"Move it to left={P(Clamp(shape.Left, 0, slideWidth - shape.Width))} top={P(Clamp(shape.Top, 0, slideHeight - shape.Height))} (shape set-position).";
            issues.Add(Issue("off-slide", "warning", slideIndex, [shape.ShapeIndex],
                $"Shape {shape.ShapeIndex} ({Describe(shape)}) extends past the slide edge: it spans {P(box.Left)}..{P(box.Right)} x {P(box.Top)}..{P(box.Bottom)} pt on a {P(slideWidth)} x {P(slideHeight)} pt slide.",
                fix));
        }
    }

    private static void CheckFontSize(int slideIndex, SlideShapeInfo shape, List<SlideLayoutIssue> issues)
    {
        if (shape.HasText && shape.MinFontSize is { } size && size > 0 && size < MinReadableFontSize)
        {
            issues.Add(Issue("small-text", "warning", slideIndex, [shape.ShapeIndex],
                $"Shape {shape.ShapeIndex} ({Describe(shape)}) uses {P(size)} pt text, below the {P(MinReadableFontSize)} pt readability floor.",
                "Raise the font size to at least 12 pt (textframe set-font-size), or move the detail into speaker notes."));
        }
    }

    private static void CheckEmptyPlaceholder(int slideIndex, SlideShapeInfo shape, List<SlideLayoutIssue> issues)
    {
        if (IsEmptyPlaceholder(shape))
        {
            issues.Add(Issue("empty-placeholder", "warning", slideIndex, [shape.ShapeIndex],
                $"Placeholder {shape.ShapeIndex} ({shape.PlaceholderType}) is empty; it reserves space and shows a prompt while editing.",
                $"Fill it with shape set-placeholder-text (shape_index {shape.ShapeIndex}) or remove it with shape delete."));
        }
    }

    private static void CheckOverlap(
        int slideIndex, float slideHeight, SlideShapeInfo first, SlideShapeInfo second, List<SlideLayoutIssue> issues)
    {
        var a = BoundsOf(first);
        var b = BoundsOf(second);
        var overlapWidth = MathF.Min(a.Right, b.Right) - MathF.Max(a.Left, b.Left);
        var overlapHeight = MathF.Min(a.Bottom, b.Bottom) - MathF.Max(a.Top, b.Top);
        if (overlapWidth <= MinOverlap || overlapHeight <= MinOverlap)
            return;

        var size = $"{P(overlapWidth)} x {P(overlapHeight)} pt";
        if (IsTextual(first) && IsTextual(second))
        {
            var (upper, lower) = a.Top <= b.Top ? (first, second) : (second, first);
            var target = BoundsOf(upper).Bottom + SuggestedGap;
            var fix = target + lower.Height <= slideHeight
                ? $"Move shape {lower.ShapeIndex} down to top={P(target)} (shape set-position), or place the two side by side."
                : $"There is no room below shape {upper.ShapeIndex}: shorten or shrink one of them, place them side by side, or move part of the content to another slide.";
            issues.Add(Issue("text-overlap", "error", slideIndex, [first.ShapeIndex, second.ShapeIndex],
                $"Text in shapes {first.ShapeIndex} ({Describe(first)}) and {second.ShapeIndex} ({Describe(second)}) overlaps by {size}.",
                fix));
            return;
        }

        // One shape fully inside the other is deliberate layering (text on a card, caption on a photo).
        if (Contains(a, b) || Contains(b, a))
            return;

        issues.Add(Issue("partial-overlap", "warning", slideIndex, [first.ShapeIndex, second.ShapeIndex],
            $"Shapes {first.ShapeIndex} ({Describe(first)}) and {second.ShapeIndex} ({Describe(second)}) partially overlap by {size}.",
            "Move or resize one so they no longer overlap, or place one entirely inside the other if the layering is intended."));
    }

    private static void CheckAlignment(int slideIndex, SlideShapeInfo first, SlideShapeInfo second, List<SlideLayoutIssue> issues)
    {
        if (!IsUpright(first) || !IsUpright(second))
            return;

        AddNearMiss("left", first.Left, second.Left);
        AddNearMiss("top", first.Top, second.Top);

        void AddNearMiss(string edge, float firstValue, float secondValue)
        {
            var difference = MathF.Abs(firstValue - secondValue);
            if (difference <= 0.5f || difference > AlignmentBand)
                return;

            issues.Add(Issue("near-misaligned", "info", slideIndex, [first.ShapeIndex, second.ShapeIndex],
                $"The {edge} edges of shapes {first.ShapeIndex} and {second.ShapeIndex} differ by {P(difference)} pt, which looks accidental.",
                $"Set {edge}={P(firstValue)} on shape {second.ShapeIndex} (shape set-position) or align them with shape align."));
        }
    }

    private static bool IsContentBlock(SlideShapeInfo shape, float slideWidth, float slideHeight) =>
        shape.Kind is not ("line" or "connector") &&
        shape.Width > 0 && shape.Height > 0 &&
        !IsBackdrop(shape, slideWidth, slideHeight) &&
        !IsEmptyPlaceholder(shape);

    private static bool IsBackdrop(SlideShapeInfo shape, float slideWidth, float slideHeight)
    {
        var box = BoundsOf(shape);
        var visibleWidth = MathF.Min(box.Right, slideWidth) - MathF.Max(box.Left, 0);
        var visibleHeight = MathF.Min(box.Bottom, slideHeight) - MathF.Max(box.Top, 0);
        return !IsTextual(shape) && visibleWidth > 0 && visibleHeight > 0 &&
            visibleWidth * visibleHeight >= BackdropCoverage * slideWidth * slideHeight;
    }

    private static bool IsEmptyPlaceholder(SlideShapeInfo shape) =>
        shape.Kind == "placeholder" && !shape.HasText && shape.HasPicture != true && shape.HasChart != true &&
        shape.TableRows is null && !NonContentPlaceholders.Contains(shape.PlaceholderType ?? "");

    private static bool IsTextual(SlideShapeInfo shape) => shape.HasText || shape.TableRows is not null;

    private static bool IsUpright(SlideShapeInfo shape)
    {
        var rotation = ((shape.Rotation % 180f) + 180f) % 180f;
        return rotation < 0.5f || rotation > 179.5f;
    }

    private static Box BoundsOf(SlideShapeInfo shape)
    {
        if (IsUpright(shape))
            return new Box(shape.Left, shape.Top, shape.Left + shape.Width, shape.Top + shape.Height);

        var radians = shape.Rotation * MathF.PI / 180f;
        var cos = MathF.Abs(MathF.Cos(radians));
        var sin = MathF.Abs(MathF.Sin(radians));
        var width = (shape.Width * cos) + (shape.Height * sin);
        var height = (shape.Width * sin) + (shape.Height * cos);
        var centerX = shape.Left + (shape.Width / 2);
        var centerY = shape.Top + (shape.Height / 2);
        return new Box(centerX - (width / 2), centerY - (height / 2), centerX + (width / 2), centerY + (height / 2));
    }

    private static bool Contains(Box outer, Box inner) =>
        inner.Left >= outer.Left - Tolerance && inner.Top >= outer.Top - Tolerance &&
        inner.Right <= outer.Right + Tolerance && inner.Bottom <= outer.Bottom + Tolerance;

    private static string Describe(SlideShapeInfo shape)
    {
        var kind = shape.PlaceholderType is { } placeholder ? $"{shape.Kind} {placeholder}" : shape.Kind;
        return string.IsNullOrEmpty(shape.Name) ? kind : $"{kind} '{shape.Name}'";
    }

    private static string LargestFont(SlideShapeInfo shape) =>
        shape.MaxFontSize is { } size ? $" (largest now {P(size)} pt)" : "";

    private static float Clamp(float value, float min, float max) => max < min ? min : Math.Clamp(value, min, max);

    private static int SeverityRank(string severity) => severity switch
    {
        "error" => 0,
        "warning" => 1,
        _ => 2,
    };

    private static string P(float value) => value.ToString("0.#", CultureInfo.InvariantCulture);

    private static SlideLayoutIssue Issue(
        string code, string severity, int slideIndex, int[] shapeIndexes, string message, string suggestion) =>
        new()
        {
            Code = code,
            Severity = severity,
            SlideIndex = slideIndex,
            ShapeIndexes = shapeIndexes,
            Message = message,
            Suggestion = suggestion,
        };

    private readonly record struct Box(float Left, float Top, float Right, float Bottom);
}
