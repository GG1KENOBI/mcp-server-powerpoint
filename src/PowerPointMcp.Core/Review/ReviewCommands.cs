extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Deck;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Deck validation and repair commands.</summary>
public sealed class ReviewCommands : IReviewCommands
{
    private const int MaxPageSize = 200;
    private const int MaxScaledRuns = 400;
    private const float GeometryStaleTolerance = 0.5f;

    private static readonly string[] Severities = ["info", "warning", "error"];
    private static readonly string[] FitPolicies = ["expand", "shrink"];

    /// <inheritdoc/>
    public ReviewOperationResult Validate(
        IPresentationBatch batch,
        int? slideIndex = null,
        IReadOnlyList<string>? rules = null,
        float minFontSize = 10f,
        float? safeMargin = null,
        bool includeHeuristic = true,
        string minSeverity = "info",
        int offset = 0,
        int limit = 50)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (Array.IndexOf(Severities, minSeverity) < 0)
            return Fail("min_severity must be info, warning, or error.");
        if (offset < 0)
            return Fail("offset must be 0 or greater.");
        if (limit is < 1 or > MaxPageSize)
            return Fail($"limit must be between 1 and {MaxPageSize}.");
        if (CheckCodes(rules, "rules") is { } codeError)
            return Fail(codeError);
        if (CheckThresholds(minFontSize, safeMargin) is { } thresholdError)
            return Fail(thresholdError);

        var options = Options(rules, minFontSize, safeMargin, includeHeuristic);
        return batch.Execute((ctx, ct) =>
        {
            var snapshot = ReadSnapshot(ctx.Presentation, slideIndex, ct, out var error);
            if (error is not null)
                return Fail(error);

            var report = DeckValidator.Validate(snapshot.Width, snapshot.Height, snapshot.Deck, options);
            var minimum = Array.IndexOf(Severities, minSeverity);
            var filtered = report.Findings.Where(finding => Array.IndexOf(Severities, finding.Severity) >= minimum).ToList();
            var page = filtered.Skip(offset).Take(limit).ToList();
            return new ReviewOperationResult
            {
                Success = true,
                Revision = snapshot.Revision,
                Findings = page,
                TotalFindings = filtered.Count,
                Offset = offset,
                HasMore = offset + page.Count < filtered.Count,
                Counts = Count(report.Findings),
                RulesRun = report.RulesRun,
                NotChecked = report.NotChecked,
            };
        });
    }

    /// <inheritdoc/>
    public ReviewOperationResult PlanRepair(
        IPresentationBatch batch,
        int? slideIndex = null,
        IReadOnlyList<string>? codes = null,
        IReadOnlyList<string>? findingIds = null,
        IReadOnlyList<string>? fitPolicy = null,
        float minFontSize = 10f,
        float? safeMargin = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (CheckRepairArguments(codes, fitPolicy, minFontSize, safeMargin) is { } error)
            return Fail(error);

        return batch.Execute((ctx, ct) =>
        {
            var snapshot = ReadSnapshot(ctx.Presentation, slideIndex, ct, out var readError);
            if (readError is not null)
                return Fail(readError);
            var (report, plan) = PlanFor(snapshot, codes, findingIds, fitPolicy, minFontSize, safeMargin);
            return new ReviewOperationResult
            {
                Success = true,
                Revision = snapshot.Revision,
                Actions = plan.Actions,
                Unresolved = plan.Unresolved,
                Counts = Count(report.Findings),
                Warnings = ["Dry run: nothing was changed. Run repair with the same arguments (and expected_revision) to apply."],
            };
        });
    }

    /// <inheritdoc/>
    public ReviewOperationResult Repair(
        IPresentationBatch batch,
        int? slideIndex = null,
        IReadOnlyList<string>? codes = null,
        IReadOnlyList<string>? findingIds = null,
        IReadOnlyList<string>? fitPolicy = null,
        float minFontSize = 10f,
        float? safeMargin = null,
        string? expectedRevision = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (CheckRepairArguments(codes, fitPolicy, minFontSize, safeMargin) is { } error)
            return Fail(error);

        return batch.Execute((ctx, ct) =>
        {
            var before = ReadSnapshot(ctx.Presentation, slideIndex, ct, out var readError);
            if (readError is not null)
                return Fail(readError);
            if (expectedRevision is not null && slideIndex is null && !string.Equals(expectedRevision, before.Revision, StringComparison.Ordinal))
            {
                return Fail($"The deck changed since revision {expectedRevision} (now {before.Revision}). Validate again and repair with the new revision.");
            }

            var (_, plan) = PlanFor(before, codes, findingIds, fitPolicy, minFontSize, safeMargin);
            var outcomes = new List<RepairOutcome>();
            foreach (var action in plan.Actions)
            {
                ct.ThrowIfCancellationRequested();
                outcomes.Add(Apply(ctx.Presentation, action));
            }

            var after = ReadSnapshot(ctx.Presentation, slideIndex, ct, out _);
            var options = Options(null, minFontSize, safeMargin, includeHeuristic: true);
            var remaining = DeckValidator.Validate(after.Width, after.Height, after.Deck, options);
            var warnings = new List<string> { "The file was not saved. Preview the result, then save with presentation save." };
            if (expectedRevision is not null && slideIndex is not null)
                warnings.Add("expected_revision is only checked for whole-deck repairs and was ignored.");
            return new ReviewOperationResult
            {
                Success = true,
                Revision = before.Revision,
                RevisionAfter = after.Revision,
                Actions = plan.Actions,
                Outcomes = outcomes,
                Unresolved = plan.Unresolved,
                Findings = remaining.Findings.Take(MaxPageSize).ToList(),
                TotalFindings = remaining.Findings.Count,
                Counts = Count(remaining.Findings),
                Changes = DeckDiff.Compare(before.Deck, after.Deck),
                Warnings = warnings,
            };
        });
    }

    /// <inheritdoc/>
    public ReviewOperationResult ListRules(IPresentationBatch batch) =>
        new() { Success = true, Rules = DeckValidator.Rules };

    private static (ValidationReport Report, RepairPlan Plan) PlanFor(
        Snapshot snapshot,
        IReadOnlyList<string>? codes,
        IReadOnlyList<string>? findingIds,
        IReadOnlyList<string>? fitPolicy,
        float minFontSize,
        float? safeMargin)
    {
        var report = DeckValidator.Validate(snapshot.Width, snapshot.Height, snapshot.Deck,
            Options(null, minFontSize, safeMargin, includeHeuristic: true));
        var plan = RepairPlanner.Plan(snapshot.Width, snapshot.Height, snapshot.Deck, report.Findings, new RepairOptions
        {
            Codes = codes is { Count: > 0 } ? codes.ToHashSet(StringComparer.Ordinal) : null,
            FindingIds = findingIds is { Count: > 0 } ? findingIds.ToHashSet(StringComparer.Ordinal) : null,
            FitPolicy = fitPolicy is { Count: > 0 } ? fitPolicy : FitPolicies,
            MinFontSize = minFontSize,
            SafeMargin = safeMargin ?? 0f,
        });
        return (report, plan);
    }

    private static ValidationOptions Options(IReadOnlyList<string>? rules, float minFontSize, float? safeMargin, bool includeHeuristic) => new()
    {
        Rules = rules is { Count: > 0 } ? rules.ToHashSet(StringComparer.Ordinal) : null,
        MinFontSize = minFontSize,
        SafeMargin = safeMargin,
        IncludeHeuristic = includeHeuristic,
        InstalledFonts = FontCatalog.InstalledFamilies(),
    };

    private static RepairOutcome Apply(PowerPoint.Presentation presentation, RepairAction action)
    {
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        try
        {
            slide = DeckShapeLocator.FindSlide(presentation, action.SlideId);
            if (slide is null)
                return Outcome(action, "skipped", $"Slide {action.SlideId} no longer exists.");

            if (action.Operation == "assign-id" && action.ShapeId == 0)
            {
                SetTag(slide.Tags, DeckRoles.IdTag, action.NewAppId!);
                return Outcome(action, "applied", $"Slide id set to '{action.NewAppId}'.");
            }

            shape = DeckShapeLocator.FindShape(slide, action.ShapeId);
            if (shape is null)
                return Outcome(action, "skipped", $"Shape {action.ShapeId} no longer exists.");

            if (action.Operation == "assign-id")
            {
                SetTag(shape.Tags, DeckRoles.IdTag, action.NewAppId!);
                return Outcome(action, "applied", $"Object id set to '{action.NewAppId}'.");
            }

            if (action.Before is { Count: 4 } planned && !SameGeometry(shape, planned))
            {
                return Outcome(action, "skipped",
                    $"The object moved or was resized since planning (now {Geometry(shape)}); validate and plan again.");
            }

            return action.Operation switch
            {
                "move" or "resize" or "move-resize" => ApplyGeometry(shape, action),
                "scale-font" => ApplyFontScale(shape, action),
                _ => Outcome(action, "skipped", $"Unknown operation {action.Operation}."),
            };
        }
        finally
        {
            if (shape is not null) ComUtilities.Release(ref shape);
            if (slide is not null) ComUtilities.Release(ref slide);
        }
    }

    private static RepairOutcome ApplyGeometry(PowerPoint.Shape shape, RepairAction action)
    {
        var after = action.After!;
        var lockState = shape.LockAspectRatio;
        try
        {
            // Unlock so width and height are set independently (pictures lock by default).
            if (action.Operation != "move")
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            if (action.Operation != "move")
            {
                shape.Width = after[2];
                shape.Height = after[3];
            }
            shape.Left = after[0];
            shape.Top = after[1];
        }
        finally
        {
            shape.LockAspectRatio = lockState;
        }
        return Outcome(action, "applied", $"Now {Geometry(shape)}.");
    }

    private static RepairOutcome ApplyFontScale(PowerPoint.Shape shape, RepairAction action)
    {
        if (shape.HasTextFrame != Office.MsoTriState.msoTrue || action.FontSizeBefore is not { } before || before <= 0)
            return Outcome(action, "skipped", "The object has no text frame.");

        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            frame = shape.TextFrame;
            range = frame.TextRange;
            var sizes = ReadRunSizes(range);
            if (sizes is null)
                return Outcome(action, "skipped", $"The text has more than {MaxScaledRuns} formatting runs; scale it manually.");
            var smallest = sizes.Where(size => size > 0).DefaultIfEmpty(0).Min();
            if (smallest <= 0)
                return Outcome(action, "skipped", "Font sizes could not be read.");

            if (action.Code == "small-text")
            {
                var factor = action.FontSizeAfter!.Value / smallest;
                WriteRunSizes(range, sizes, factor);
                return Outcome(action, "applied", $"Scaled every run by {factor:0.##}; smallest text is now {Fmt(smallest * factor)} pt.");
            }

            // Overflow: step the smallest size down 0.5 pt at a time, measuring after each step.
            var floor = action.MinFontSize ?? 10f;
            var margins = frame.MarginTop + frame.MarginBottom;
            for (var target = smallest - 0.5f; target >= floor - 0.001f; target -= 0.5f)
            {
                WriteRunSizes(range, sizes, target / smallest);
                if (range.BoundHeight + margins <= shape.Height + 1f)
                    return Outcome(action, "applied", $"Scaled every run by {target / smallest:0.##}; smallest text {Fmt(smallest)} → {Fmt(target)} pt. Measured text height {Fmt(range.BoundHeight)} pt fits the {Fmt(shape.Height)} pt box.");
            }

            WriteRunSizes(range, sizes, 1f);
            return Outcome(action, "failed", $"The text still overflows at {Fmt(floor)} pt; original sizes were restored. Split the content or enlarge the box.");
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    private static List<float>? ReadRunSizes(PowerPoint.TextRange range)
    {
        PowerPoint.TextRange? runs = null;
        try
        {
            runs = range.Runs();
            if (runs.Count > MaxScaledRuns)
                return null;
            var sizes = new List<float>(runs.Count);
            for (int index = 1; index <= runs.Count; index++)
            {
                PowerPoint.TextRange? run = null;
                PowerPoint.Font? font = null;
                try
                {
                    run = range.Runs(index, 1);
                    font = run.Font;
                    sizes.Add(font.Size);
                }
                finally
                {
                    if (font is not null) ComUtilities.Release(ref font);
                    if (run is not null) ComUtilities.Release(ref run);
                }
            }
            return sizes;
        }
        finally
        {
            if (runs is not null) ComUtilities.Release(ref runs);
        }
    }

    private static void WriteRunSizes(PowerPoint.TextRange range, List<float> sizes, float factor)
    {
        for (int index = 1; index <= sizes.Count; index++)
        {
            if (sizes[index - 1] <= 0)
                continue;
            PowerPoint.TextRange? run = null;
            PowerPoint.Font? font = null;
            try
            {
                run = range.Runs(index, 1);
                font = run.Font;
                font.Size = MathF.Round(sizes[index - 1] * factor * 10f) / 10f;
            }
            finally
            {
                if (font is not null) ComUtilities.Release(ref font);
                if (run is not null) ComUtilities.Release(ref run);
            }
        }
    }

    private static void SetTag(PowerPoint.Tags tags, string name, string value)
    {
        try
        {
            tags.Add(name, value);
        }
        finally
        {
            ComUtilities.Release(ref tags!);
        }
    }

    private static bool SameGeometry(PowerPoint.Shape shape, IReadOnlyList<float> planned) =>
        Math.Abs(shape.Left - planned[0]) <= GeometryStaleTolerance && Math.Abs(shape.Top - planned[1]) <= GeometryStaleTolerance &&
        Math.Abs(shape.Width - planned[2]) <= GeometryStaleTolerance && Math.Abs(shape.Height - planned[3]) <= GeometryStaleTolerance;

    private static string Geometry(PowerPoint.Shape shape) =>
        $"left={Fmt(shape.Left)} top={Fmt(shape.Top)} width={Fmt(shape.Width)} height={Fmt(shape.Height)}";

    private static string Fmt(float value) => value.ToString("0.#", CultureInfo.InvariantCulture);

    private static RepairOutcome Outcome(RepairAction action, string status, string detail) =>
        new() { FindingId = action.FindingId, Operation = action.Operation, ShapeId = action.ShapeId, Status = status, Detail = detail };

    private sealed record Snapshot(
        float Width,
        float Height,
        List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> Deck,
        string Revision);

    private static Snapshot ReadSnapshot(PowerPoint.Presentation presentation, int? slideIndex, CancellationToken ct, out string? error)
    {
        PowerPoint.PageSetup? pageSetup = null;
        try
        {
            pageSetup = presentation.PageSetup;
            var deck = DeckCommands.ReadTargetSlides(presentation, slideIndex, null, ct, out error);
            return new Snapshot(pageSetup.SlideWidth, pageSetup.SlideHeight, deck, DeckCommands.Revision(deck));
        }
        finally
        {
            if (pageSetup is not null) ComUtilities.Release(ref pageSetup);
        }
    }

    private static Dictionary<string, int> Count(IReadOnlyList<DeckFinding> findings)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var finding in findings)
        {
            counts[finding.Severity] = counts.GetValueOrDefault(finding.Severity) + 1;
            counts[finding.Certainty] = counts.GetValueOrDefault(finding.Certainty) + 1;
        }
        return counts;
    }

    private static string? CheckCodes(IReadOnlyList<string>? codes, string name)
    {
        if (codes is null)
            return null;
        var known = DeckValidator.Rules.Select(rule => rule.Code).ToHashSet(StringComparer.Ordinal);
        var unknown = codes.Where(code => !known.Contains(code)).ToList();
        return unknown.Count == 0
            ? null
            : $"Unknown {name}: {string.Join(", ", unknown)}. Valid codes: {string.Join(", ", known)}.";
    }

    private static string? CheckThresholds(float minFontSize, float? safeMargin)
    {
        if (minFontSize is < 4f or > 72f)
            return "min_font_size must be between 4 and 72 points.";
        if (safeMargin is < 0f or > 144f)
            return "safe_margin must be between 0 and 144 points.";
        return null;
    }

    private static string? CheckRepairArguments(IReadOnlyList<string>? codes, IReadOnlyList<string>? fitPolicy, float minFontSize, float? safeMargin)
    {
        if (CheckCodes(codes, "codes") is { } codeError)
            return codeError;
        if (fitPolicy?.FirstOrDefault(policy => Array.IndexOf(FitPolicies, policy) < 0) is { } badPolicy)
            return $"Unknown fit_policy '{badPolicy}'. Use expand and/or shrink, in the order to try them.";
        return CheckThresholds(minFontSize, safeMargin);
    }

    private static ReviewOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
