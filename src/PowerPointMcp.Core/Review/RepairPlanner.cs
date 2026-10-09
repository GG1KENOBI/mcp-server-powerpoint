using System.Globalization;
using Sbroenne.PowerPointMcp.Core.Deck;
using static Sbroenne.PowerPointMcp.Core.Deck.DeckGeometry;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Options for <see cref="RepairPlanner.Plan"/>.</summary>
public sealed class RepairOptions
{
    /// <summary>Rule codes to repair; null uses <see cref="DeckValidator.DefaultRepairCodes"/>.</summary>
    public IReadOnlySet<string>? Codes { get; init; }

    /// <summary>Only repair these finding ids (null repairs every finding of the selected codes).</summary>
    public IReadOnlySet<string>? FindingIds { get; init; }

    /// <summary>
    /// Text-overflow strategies in order: expand (grow the box into free space below),
    /// shrink (scale every run's font proportionally, not below <see cref="MinFontSize"/>).
    /// </summary>
    public IReadOnlyList<string> FitPolicy { get; init; } = ["expand", "shrink"];

    /// <summary>Font floor for shrinking, and the target for small-text.</summary>
    public float MinFontSize { get; init; } = 10f;

    /// <summary>Margin kept free at slide edges when moving or growing objects.</summary>
    public float SafeMargin { get; init; }

    /// <summary>Prefix for regenerated persistent ids.</summary>
    public string IdPrefix { get; init; } = "id";
}

/// <summary>A repair plan: actions to apply and findings left unresolved (with the reason).</summary>
public sealed class RepairPlan
{
    /// <summary>Ordered actions.</summary>
    public required IReadOnlyList<RepairAction> Actions { get; init; }

    /// <summary>Selected findings no action addresses, with why.</summary>
    public required IReadOnlyList<UnresolvedFinding> Unresolved { get; init; }
}

/// <summary>A finding repair will not fix, and why.</summary>
public sealed class UnresolvedFinding
{
    /// <summary>The finding.</summary>
    public required DeckFinding Finding { get; init; }

    /// <summary>Why it is not repaired automatically.</summary>
    public required string Reason { get; init; }
}

/// <summary>
/// Turns validation findings into geometry and font-scaling actions. Pure. Plans never delete
/// objects, change text content, flatten formatting, or rasterize; anything that would need that
/// is returned as unresolved with a reason.
/// </summary>
public static class RepairPlanner
{
    private const float Gap = 12f;

    private static readonly string[] Priority =
    [
        "off-slide", "outside-safe-area", "text-overflow", "small-text", "text-overlap", "image-distortion",
        "inconsistent-title-position", "near-misaligned", "duplicate-app-id",
    ];

    /// <summary>Plans repairs for the findings of one validation run over the same snapshot.</summary>
    public static RepairPlan Plan(
        float slideWidth,
        float slideHeight,
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck,
        IReadOnlyList<DeckFinding> findings,
        RepairOptions options)
    {
        ArgumentNullException.ThrowIfNull(deck);
        ArgumentNullException.ThrowIfNull(findings);
        ArgumentNullException.ThrowIfNull(options);

        var codes = options.Codes ?? DeckValidator.DefaultRepairCodes;
        var selected = findings
            .Where(finding => codes.Contains(finding.Code) && (options.FindingIds is null || options.FindingIds.Contains(finding.Id)))
            .OrderBy(finding => Array.IndexOf(Priority, finding.Code) is var rank && rank < 0 ? int.MaxValue : rank)
            .ThenBy(finding => finding.SlideIndex)
            .ToList();

        var objects = deck.SelectMany(entry => entry.Objects).ToDictionary(item => (item.SlideId, item.ShapeId));
        var slideObjects = deck.ToDictionary(entry => entry.Slide.SlideId, entry => entry.Objects);
        var actions = new List<RepairAction>();
        var unresolved = new List<UnresolvedFinding>();
        var touched = new HashSet<(int SlideId, int ShapeId)>();
        // Planned geometry per object, so later plans see earlier moves.
        var planned = new Dictionary<(int SlideId, int ShapeId), Box>();
        var context = new PlanContext(slideWidth, slideHeight, options, slideObjects, planned);

        foreach (var finding in selected)
        {
            if (!finding.AutoRepairable)
            {
                unresolved.Add(new UnresolvedFinding { Finding = finding, Reason = $"Not auto-repairable. {finding.Suggestion}" });
                continue;
            }

            if (finding.Code == "duplicate-app-id")
            {
                PlanDuplicateIds(deck, finding, options, actions);
                continue;
            }

            var targets = (finding.ShapeIds ?? []).Select(id => objects.GetValueOrDefault((finding.SlideId, id))).ToList();
            if (targets.Count == 0 || targets.Any(item => item is null))
            {
                unresolved.Add(new UnresolvedFinding { Finding = finding, Reason = "The objects in this finding were not found in the snapshot." });
                continue;
            }

            var resolved = targets.OfType<DeckObjectInfo>().ToList();
            var target = PickTarget(finding, resolved);
            if (touched.Contains((target.SlideId, target.ShapeId)))
            {
                unresolved.Add(new UnresolvedFinding
                {
                    Finding = finding,
                    Reason = "Another repair already changes this object in this run; validate again after repairing.",
                });
                continue;
            }

            var (planActions, reason) = finding.Code switch
            {
                "off-slide" => PlanIntoBounds(context, finding, target, 0f),
                "outside-safe-area" => PlanIntoBounds(context, finding, target, Math.Max(options.SafeMargin, ParseEvidence(finding, "safeMargin") ?? 0f)),
                "text-overflow" => PlanOverflow(context, finding, target),
                "small-text" => PlanSmallText(finding, target, options),
                "text-overlap" => PlanOverlap(context, finding, resolved),
                "image-distortion" => PlanAspect(finding, target),
                "inconsistent-title-position" => PlanMoveTo(finding, target, ParseEvidence(finding, "expectedLeft"), ParseEvidence(finding, "expectedTop"), "Align the title with the other slides of this layout."),
                "near-misaligned" => PlanAlign(finding, resolved),
                _ => ([], "No automatic repair for this rule."),
            };

            if (planActions.Count == 0)
            {
                unresolved.Add(new UnresolvedFinding { Finding = finding, Reason = reason ?? finding.Suggestion });
                continue;
            }

            foreach (var action in planActions)
            {
                touched.Add((action.SlideId, action.ShapeId));
                if (action.After is { Count: 4 } after)
                    planned[(action.SlideId, action.ShapeId)] = Box.FromSize(after[0], after[1], after[2], after[3]);
            }
            actions.AddRange(planActions);
        }

        return new RepairPlan { Actions = actions, Unresolved = unresolved };
    }

    private static DeckObjectInfo PickTarget(DeckFinding finding, List<DeckObjectInfo> targets) =>
        finding.Code is "near-misaligned" && targets.Count > 1 ? targets[1] : targets[0];

    private static (List<RepairAction>, string?) PlanIntoBounds(PlanContext context, DeckFinding finding, DeckObjectInfo item, float margin)
    {
        var box = BoundsOf(item);
        var minX = margin;
        var minY = margin;
        var maxX = context.Width - margin;
        var maxY = context.Height - margin;
        if (box.Width > maxX - minX + 0.5f || box.Height > maxY - minY + 0.5f)
            return ([], $"The object ({P(box.Width)} x {P(box.Height)} pt) is larger than the available area; resizing would change its proportions or text wrapping, so it is left for you.");

        float dx = box.Left < minX ? minX - box.Left : box.Right > maxX ? maxX - box.Right : 0f;
        float dy = box.Top < minY ? minY - box.Top : box.Bottom > maxY ? maxY - box.Bottom : 0f;
        if (Math.Abs(dx) < 0.01f && Math.Abs(dy) < 0.01f)
            return ([], "Already inside the area.");
        return ([Move(finding, item, item.Left + dx, item.Top + dy, "geometry",
            $"Shift by ({P(dx)}, {P(dy)}) pt so the whole object is inside the {(margin > 0 ? "safe area" : "slide")}.")], null);
    }

    private static (List<RepairAction>, string?) PlanOverflow(PlanContext context, DeckFinding finding, DeckObjectInfo item)
    {
        if (item.TextBounds is not { Count: 4 } bounds)
            return ([], "No measured text bounds.");

        bool horizontal = finding.Evidence?.ContainsKey("textWidth") == true;
        if (horizontal)
        {
            var neededWidth = MathF.Ceiling(bounds[2] + DeckValidator.HorizontalMargins(item));
            var limit = context.Width - context.Options.SafeMargin;
            if (context.Options.FitPolicy.Contains("expand") && item.Left + neededWidth <= limit &&
                !context.CollidesWith(item, Box.FromSize(item.Left, item.Top, neededWidth, item.Height)))
            {
                return ([Resize(finding, item, item.Left, item.Top, neededWidth, item.Height, "measured", $"Widen to {P(neededWidth)} pt, the measured text width plus margins.")], null);
            }
            return ([], "The unwrapped text needs more width than is free to the right; turn on word wrap (textframe) or shorten the line.");
        }

        var margins = DeckValidator.VerticalMargins(item);
        var needed = MathF.Ceiling(bounds[3] + margins);
        var floor = context.Height - context.Options.SafeMargin;
        var below = context.NextObstacleBelow(item);
        var limitBottom = Math.Min(floor, below - Gap);
        var actions = new List<RepairAction>();
        var available = item.Height;

        if (context.Options.FitPolicy.Contains("expand"))
        {
            var room = limitBottom - item.Top;
            if (room >= needed)
            {
                return ([Resize(finding, item, item.Left, item.Top, item.Width, needed, "measured",
                    $"Grow the box to {P(needed)} pt (measured text height {P(bounds[3])} + margins); free space below allows it.")], null);
            }
            if (room > item.Height + 1f)
            {
                actions.Add(Resize(finding, item, item.Left, item.Top, item.Width, MathF.Floor(room), "measured",
                    $"Grow the box into the {P(room - item.Height)} pt of free space below."));
                available = MathF.Floor(room);
            }
        }

        if (context.Options.FitPolicy.Contains("shrink") && item.MinFontSize is { } minFont && minFont > 0)
        {
            // Wrapped text height scales roughly with the square of the font size.
            var scale = MathF.Sqrt(Math.Max(0.01f, (available - margins) / bounds[3]));
            var estimated = MathF.Floor(minFont * scale * 2f) / 2f;
            if (estimated >= context.Options.MinFontSize)
            {
                actions.Add(new RepairAction
                {
                    FindingId = finding.Id,
                    Code = finding.Code,
                    Operation = "scale-font",
                    SlideId = item.SlideId,
                    SlideIndex = item.SlideIndex,
                    ShapeId = item.ShapeId,
                    Before = Geometry(item),
                    FontSizeBefore = minFont,
                    FontSizeAfter = estimated,
                    MinFontSize = context.Options.MinFontSize,
                    Basis = "estimated",
                    Note = $"Scale every run proportionally (smallest {P(minFont)} pt → about {P(estimated)} pt); repair re-measures and stops at the first size that fits, never below {P(context.Options.MinFontSize)} pt.",
                });
                return (actions, null);
            }
            return ([], $"Fitting needs the smallest text at about {P(estimated)} pt, below the {P(context.Options.MinFontSize)} pt minimum. Split the content across slides (compose with fit=split), shorten it, or enlarge the box by moving neighbors.");
        }

        return actions.Count > 0
            ? (actions, null)
            : ([], "No free space to grow into and font shrinking is not in fit_policy. Add \"shrink\" to fit_policy, split the content, or move the objects below.");
    }

    private static (List<RepairAction>, string?) PlanSmallText(DeckFinding finding, DeckObjectInfo item, RepairOptions options)
    {
        if (item.MinFontSize is not { } minFont || minFont <= 0)
            return ([], "Font size unknown.");
        var threshold = ParseEvidence(finding, "threshold") ?? options.MinFontSize;
        return ([new RepairAction
        {
            FindingId = finding.Id,
            Code = finding.Code,
            Operation = "scale-font",
            SlideId = item.SlideId,
            SlideIndex = item.SlideIndex,
            ShapeId = item.ShapeId,
            Before = Geometry(item),
            FontSizeBefore = minFont,
            FontSizeAfter = threshold,
            MinFontSize = threshold,
            Basis = "geometry",
            Note = $"Scale every run by {threshold / minFont:0.##} so the smallest text is {P(threshold)} pt; validate afterwards because larger text can overflow.",
        }], null);
    }

    private static (List<RepairAction>, string?) PlanOverlap(PlanContext context, DeckFinding finding, List<DeckObjectInfo> items)
    {
        var a = context.Current(items[0]);
        var b = context.Current(items[1]);
        var (upper, lower, upperBox, lowerBox) = a.Top <= b.Top ? (items[0], items[1], a, b) : (items[1], items[0], b, a);
        var target = upperBox.Bottom + Gap;
        var moved = Box.FromSize(lowerBox.Left, target, lowerBox.Width, lowerBox.Height);
        if (moved.Bottom > context.Height - context.Options.SafeMargin)
            return ([], "Moving the lower object below the upper one would push it off the slide.");
        if (context.CollidesWith(lower, moved, ignore: upper.ShapeId))
            return ([], "Moving the lower object down would create a new overlap.");
        return ([Move(finding, lower, lower.Left, target, "geometry", $"Move below {upper.Name} with a {P(Gap)} pt gap.")], null);
    }

    private static (List<RepairAction>, string?) PlanAspect(DeckFinding finding, DeckObjectInfo item)
    {
        if (ParseEvidence(finding, "sourceAspect") is not { } aspect || aspect <= 0)
            return ([], "Source aspect ratio unknown.");
        float width = item.Width;
        float height = width / aspect;
        if (height > item.Height)
        {
            height = item.Height;
            width = height * aspect;
        }
        var left = item.Left + ((item.Width - width) / 2f);
        var top = item.Top + ((item.Height - height) / 2f);
        return ([Resize(finding, item, left, top, width, height, "geometry",
            $"Fit the picture inside its current frame at the source aspect {aspect:0.###}, centered ({P(width)} x {P(height)} pt).")], null);
    }

    private static (List<RepairAction>, string?) PlanMoveTo(DeckFinding finding, DeckObjectInfo item, float? left, float? top, string note) =>
        left is null || top is null ? ([], "Target position unknown.") : ([Move(finding, item, left.Value, top.Value, "geometry", note)], null);

    private static (List<RepairAction>, string?) PlanAlign(DeckFinding finding, List<DeckObjectInfo> items)
    {
        var edge = finding.Evidence?.GetValueOrDefault("edge");
        if (ParseEvidence(finding, "first") is not { } value || items.Count < 2)
            return ([], "Alignment target unknown.");
        var item = items[1];
        return edge switch
        {
            "left" => ([Move(finding, item, value, item.Top, "geometry", $"Align the left edge with {items[0].Name}.")], null),
            "top" => ([Move(finding, item, item.Left, value, "geometry", $"Align the top edge with {items[0].Name}.")], null),
            "right" => ([Move(finding, item, value - item.Width, item.Top, "geometry", $"Align the right edge with {items[0].Name}.")], null),
            _ => ([], "Unknown edge."),
        };
    }

    private static void PlanDuplicateIds(
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck,
        DeckFinding finding,
        RepairOptions options,
        List<RepairAction> actions)
    {
        var duplicate = finding.Evidence?.GetValueOrDefault("id");
        foreach (var assignment in AppIdPlanner.Plan(deck, null, options.IdPrefix).Where(item => item.Reason == "duplicate"))
        {
            var key = assignment.ShapeId is null ? $"slide:{assignment.OldAppId}" : $"object:{assignment.OldAppId}";
            if (key != duplicate || actions.Any(action => action.Operation == "assign-id" && action.SlideId == assignment.SlideId && action.ShapeId == (assignment.ShapeId ?? 0)))
                continue;
            actions.Add(new RepairAction
            {
                FindingId = finding.Id,
                Code = finding.Code,
                Operation = "assign-id",
                SlideId = assignment.SlideId,
                SlideIndex = assignment.SlideIndex,
                ShapeId = assignment.ShapeId ?? 0,
                Before = [],
                NewAppId = assignment.NewAppId,
                Basis = "geometry",
                Note = $"Rename the later copy of '{assignment.OldAppId}' to '{assignment.NewAppId}'.",
            });
        }
    }

    private static RepairAction Move(DeckFinding finding, DeckObjectInfo item, float left, float top, string basis, string note) => new()
    {
        FindingId = finding.Id,
        Code = finding.Code,
        Operation = "move",
        SlideId = item.SlideId,
        SlideIndex = item.SlideIndex,
        ShapeId = item.ShapeId,
        Before = Geometry(item),
        After = [Round(left), Round(top), item.Width, item.Height],
        Basis = basis,
        Note = note,
    };

    private static RepairAction Resize(DeckFinding finding, DeckObjectInfo item, float left, float top, float width, float height, string basis, string note) => new()
    {
        FindingId = finding.Id,
        Code = finding.Code,
        Operation = Math.Abs(left - item.Left) > 0.01f || Math.Abs(top - item.Top) > 0.01f ? "move-resize" : "resize",
        SlideId = item.SlideId,
        SlideIndex = item.SlideIndex,
        ShapeId = item.ShapeId,
        Before = Geometry(item),
        After = [Round(left), Round(top), Round(width), Round(height)],
        Basis = basis,
        Note = note,
    };

    private static float Round(float value) => MathF.Round(value * 10f) / 10f;

    private static IReadOnlyList<float> Geometry(DeckObjectInfo item) => [item.Left, item.Top, item.Width, item.Height];

    private static float? ParseEvidence(DeckFinding finding, string key) =>
        finding.Evidence?.GetValueOrDefault(key) is { } text &&
        float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;

    private sealed class PlanContext(
        float width,
        float height,
        RepairOptions options,
        IReadOnlyDictionary<int, IReadOnlyList<DeckObjectInfo>> slideObjects,
        IReadOnlyDictionary<(int SlideId, int ShapeId), Box> planned)
    {
        public float Width { get; } = width;

        public float Height { get; } = height;

        public RepairOptions Options { get; } = options;

        public Box Current(DeckObjectInfo item) => planned.TryGetValue((item.SlideId, item.ShapeId), out var box) ? box : BoundsOf(item);

        private IEnumerable<(DeckObjectInfo Item, Box Box)> Blocks(DeckObjectInfo exclude, int? ignore) =>
            slideObjects[exclude.SlideId]
                .Where(other => other.ShapeIndex is not null && other.Visible && other.ShapeId != exclude.ShapeId && other.ShapeId != ignore &&
                    DeckValidator.IsContentBlock(Width, Height, other) && other.ShapeId != exclude.GroupShapeId)
                .Select(other => (other, Current(other)));

        public bool CollidesWith(DeckObjectInfo item, Box candidate, int? ignore = null) =>
            Blocks(item, ignore).Any(other =>
            {
                var overlap = candidate.Intersect(other.Box);
                // Containment either way is layering, not a collision.
                return overlap.Width > 2f && overlap.Height > 2f && !other.Box.Contains(candidate) && !candidate.Contains(other.Box);
            });

        public float NextObstacleBelow(DeckObjectInfo item)
        {
            var box = Current(item);
            var obstacles = Blocks(item, null)
                .Where(other => other.Box.Top >= box.Bottom - 1f && other.Box.Left < box.Right && other.Box.Right > box.Left && !other.Box.Contains(box))
                .Select(other => other.Box.Top)
                .ToList();
            return obstacles.Count > 0 ? obstacles.Min() : float.MaxValue;
        }
    }
}
