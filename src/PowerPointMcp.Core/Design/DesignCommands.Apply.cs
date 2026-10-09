using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Deck;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Design;

public sealed partial class DesignCommands
{
    private static readonly HashSet<string> ShapeChangeProperties = ["font-name", "font-size", "text-color", "fill", "line", "background"];

    /// <inheritdoc/>
    public DesignOperationResult ApplyProfile(
        IPresentationBatch batch,
        string profile,
        string? fromProfile = null,
        string? selector = null,
        bool updateTheme = true,
        bool fonts = true,
        bool colors = true,
        bool sizes = true,
        bool tables = true,
        bool charts = true,
        bool dryRun = true,
        int limit = 100)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (limit is < 1 or > MaxLimit)
            return Fail($"limit must be between 1 and {MaxLimit}.");
        if (string.IsNullOrWhiteSpace(profile))
            return Fail("profile is required (a built-in or saved profile name).");
        var (target, error) = ResolveNamed(batch, profile);
        if (error is not null)
            return Fail(error);
        var context = DesignSourceContext.Create(batch, fromProfile, out error);
        if (error is not null)
            return Fail(error);

        var notes = new List<string>();
        bool writeTheme = updateTheme && string.IsNullOrWhiteSpace(selector);
        if (updateTheme && !writeTheme)
            notes.Add("The theme was left unchanged because a selector limits the scope.");

        return batch.Execute((ctx, ct) =>
        {
            var (shapes, slides, scopeError) = Scope(ctx.Presentation, selector, ct);
            if (scopeError is not null)
                return Fail(scopeError);
            var scan = DesignScanner.Scan(ctx.Presentation, shapes, slides, ct);
            if (scan.SkippedRuns > 0)
                notes.Add($"{scan.SkippedRuns} run(s) in very long text frames were not read or changed.");
            var tableIds = scan.Objects.Where(item => item.Kind == "table").Select(item => (item.SlideId, item.ShapeId)).ToHashSet();
            var runs = scan.Runs.Where(run => !(tables && run.Row is not null && tableIds.Contains((run.SlideId, run.ShapeId)))).ToList();
            var sourceBySlide = scan.Slides.ToDictionary(slide => slide.SlideId, slide => context.SourceFor(slide.ProfileTag));

            var changes = new List<DesignChange>();
            if (writeTheme)
                changes.AddRange(DesignApplier.PlanTheme(ctx.Presentation, target!));
            if (colors)
            {
                var maps = new Dictionary<string, IReadOnlyDictionary<string, (string To, string Token)>>(StringComparer.Ordinal);
                var (colorChanges, colorNotes) = StylePlanner.PlanColors(runs, scan.Fills, scan.Slides, slideId =>
                {
                    var source = sourceBySlide.GetValueOrDefault(slideId) ?? context.Theme;
                    var key = $"{source.Name}@{source.Version}";
                    if (!maps.TryGetValue(key, out var map))
                        maps[key] = map = StylePlanner.ColorMap(source, target!);
                    return map;
                }, writeTheme);
                changes.AddRange(colorChanges);
                notes.AddRange(colorNotes);
            }
            if (fonts || sizes)
            {
                var (typeChanges, typeNotes) = StylePlanner.PlanFontAndSizeMigration(runs, slideId => sourceBySlide.GetValueOrDefault(slideId) ?? context.Theme,
                    target!, fonts, sizes, context.ThemeFonts, writeTheme);
                changes.AddRange(typeChanges);
                notes.AddRange(typeNotes);
            }
            foreach (var item in scan.Objects.Where(item => (item.Kind == "table" && tables) || (item.Kind == "chart" && charts)))
                changes.Add(new DesignChange { SlideIndex = item.SlideIndex, SlideId = item.SlideId, ShapeId = item.ShapeId, ShapeName = item.ShapeName, Property = $"{item.Kind}-style", To = $"{target!.Name}@{target.Version}" });
            var tags = ComponentTagChanges(scan, target!, null);
            changes.AddRange(tags);
            foreach (var slide in scan.Slides.Where(slide => slide.ProfileTag != $"{target!.Name}@{target.Version}"))
                changes.Add(new DesignChange { SlideIndex = slide.SlideIndex, SlideId = slide.SlideId, Property = "profile-tag", From = slide.ProfileTag, To = $"{target!.Name}@{target.Version}" });

            if (dryRun)
                return PlanResult(changes, notes.Append("Dry run: nothing was changed. Pass dry_run=false to apply."), applied: false, limit, target!.Name);

            if (writeTheme)
                DesignApplier.WriteTheme(ctx.Presentation, target!);
            var (_, applyNotes) = DesignApplier.ApplyChanges(ctx.Presentation, changes.Where(change => ShapeChangeProperties.Contains(change.Property)).ToList(), ct);
            notes.AddRange(applyNotes);
            notes.AddRange(DesignApplier.RestyleObjects(ctx.Presentation, scan.Objects.Where(item => (item.Kind == "table" && tables) || (item.Kind == "chart" && charts)), target!, ct));
            DesignApplier.SetTags(ctx.Presentation, changes
                .Where(change => change.Property is "component-version" or "profile-tag")
                .Select(change => (change.SlideId!.Value, change.ShapeId, change.Property == "profile-tag" ? "PPTMCP_PROFILE" : DeckRoles.ComponentTag, change.To)));
            notes.Add("The open presentation was changed but not saved; save it under a new name to keep the original file. Run review validate to check for overflow.");
            return PlanResult(changes, notes, applied: true, limit, target!.Name);
        });
    }

    /// <inheritdoc/>
    public DesignOperationResult NormalizeTypography(
        IPresentationBatch batch,
        string? profile = null,
        string? selector = null,
        bool fonts = true,
        bool sizes = true,
        bool minSize = true,
        float tolerance = 0.25f,
        bool dryRun = true,
        int limit = 100)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (limit is < 1 or > MaxLimit)
            return Fail($"limit must be between 1 and {MaxLimit}.");
        if (tolerance is < 0 or > 1)
            return Fail("tolerance must be between 0 and 1.");
        if (!fonts && !sizes && !minSize)
            return Fail("Enable at least one of fonts, sizes, or min_size.");
        var (target, error) = ResolveNamed(batch, profile);
        if (error is not null)
            return Fail(error);

        return batch.Execute((ctx, ct) =>
        {
            var (shapes, slides, scopeError) = Scope(ctx.Presentation, selector, ct);
            if (scopeError is not null)
                return Fail(scopeError);
            var scan = DesignScanner.Scan(ctx.Presentation, shapes, slides, ct);
            var (changes, notes) = StylePlanner.PlanTypography(scan.Runs, target!, new TypographyOptions(fonts, sizes, minSize, tolerance));
            if (scan.SkippedRuns > 0)
                notes.Add($"{scan.SkippedRuns} run(s) in very long text frames were not read or changed.");
            if (dryRun)
                return PlanResult(changes, notes.Append("Dry run: nothing was changed. Pass dry_run=false to apply."), applied: false, limit, target!.Name);
            var (_, applyNotes) = DesignApplier.ApplyChanges(ctx.Presentation, changes, ct);
            notes.AddRange(applyNotes);
            notes.Add("The open presentation was changed but not saved. Run review validate to check for overflow after size changes.");
            return PlanResult(changes, notes, applied: true, limit, target!.Name);
        });
    }

    /// <inheritdoc/>
    public DesignOperationResult ListComponents(IPresentationBatch batch, string? profile = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var (target, error) = ResolveNamed(batch, profile);
        if (error is not null)
            return Fail(error);
        return batch.Execute((ctx, ct) =>
        {
            var instances = ReadInstances(ctx.Presentation, ct);
            var names = target!.Components.Keys.Union(instances.Select(item => ComponentName(item.Component!)), StringComparer.Ordinal).Order(StringComparer.Ordinal);
            var components = names.Select(name =>
            {
                var found = instances.Where(item => ComponentName(item.Component!) == name).ToList();
                var version = target.Components.TryGetValue(name, out var style) ? style.Version : null;
                return new DesignComponentInfo
                {
                    Name = name,
                    ProfileVersion = version,
                    Instances = found.Count,
                    Outdated = version is null ? 0 : found.Count(item => ComponentVersion(item.Component!) != version),
                    DeckVersions = found.Count == 0 ? null : found.GroupBy(item => ComponentVersion(item.Component!) ?? "?").ToDictionary(group => group.Key, group => group.Count()),
                    Slides = found.Count == 0 ? null : found.Select(item => item.SlideIndex).Distinct().Order().ToList(),
                };
            }).ToList();
            return new DesignOperationResult { Success = true, ProfileName = target.Name, Components = components };
        });
    }

    /// <inheritdoc/>
    public DesignOperationResult UpdateComponents(
        IPresentationBatch batch,
        string? profile = null,
        string? component = null,
        string? fromProfile = null,
        bool dryRun = true,
        int limit = 100)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (limit is < 1 or > MaxLimit)
            return Fail($"limit must be between 1 and {MaxLimit}.");
        var (target, error) = ResolveNamed(batch, profile);
        if (error is not null)
            return Fail(error);
        if (component is not null && !target!.Components.ContainsKey(component))
            return Fail($"Profile '{target.Name}' has no component '{component}'. Components: {string.Join(", ", target.Components.Keys)}.");
        var context = DesignSourceContext.Create(batch, fromProfile, out error);
        if (error is not null)
            return Fail(error);

        return batch.Execute((ctx, ct) =>
        {
            var instances = ReadInstances(ctx.Presentation, ct)
                .Where(item => target!.Components.TryGetValue(ComponentName(item.Component!), out var style)
                    && (component is null || ComponentName(item.Component!) == component)
                    && ComponentVersion(item.Component!) != style.Version)
                .ToList();
            var notes = new List<string>();
            if (instances.Count == 0)
                return PlanResult([], ["Every component instance already uses the profile's version."], applied: !dryRun, limit, target!.Name);

            var chosen = instances.Select(item => (item.SlideId, item.ShapeId)).ToHashSet();
            var scan = DesignScanner.Scan(ctx.Presentation, (slideId, shapeId) => chosen.Contains((slideId, shapeId)), slideId => instances.Any(item => item.SlideId == slideId), ct);
            var sourceBySlide = scan.Slides.ToDictionary(slide => slide.SlideId, slide => context.SourceFor(slide.ProfileTag));
            var maps = new Dictionary<string, IReadOnlyDictionary<string, (string To, string Token)>>(StringComparer.Ordinal);
            var (changes, colorNotes) = StylePlanner.PlanColors(scan.Runs, scan.Fills, [], slideId =>
            {
                var source = sourceBySlide.GetValueOrDefault(slideId) ?? context.Theme;
                var key = $"{source.Name}@{source.Version}";
                if (!maps.TryGetValue(key, out var map))
                    maps[key] = map = StylePlanner.ColorMap(source, target!);
                return map;
            }, themeUpdated: false);
            notes.AddRange(colorNotes);
            changes.AddRange(ComponentTagChanges(scan, target!, component));
            if (dryRun)
                return PlanResult(changes, notes.Append("Dry run: nothing was changed. Pass dry_run=false to apply."), applied: false, limit, target!.Name);
            var (_, applyNotes) = DesignApplier.ApplyChanges(ctx.Presentation, changes.Where(change => ShapeChangeProperties.Contains(change.Property)).ToList(), ct);
            notes.AddRange(applyNotes);
            DesignApplier.SetTags(ctx.Presentation, changes.Where(change => change.Property == "component-version")
                .Select(change => (change.SlideId!.Value, change.ShapeId, DeckRoles.ComponentTag, change.To)));
            return PlanResult(changes, notes, applied: true, limit, target!.Name);
        });
    }

    private static List<ScannedObject> ReadInstances(PowerPoint.Presentation presentation, CancellationToken cancellationToken) =>
        DesignScanner.Scan(presentation, null, null, cancellationToken).Objects.Where(item => item.Kind == "component" && item.Component is not null).ToList();

    private static List<DesignChange> ComponentTagChanges(DeckStyleScan scan, ResolvedProfile target, string? onlyComponent) =>
        scan.Objects
            .Where(item => item.Kind == "component" && item.Component is not null)
            .Where(item => onlyComponent is null || ComponentName(item.Component!) == onlyComponent)
            .Select(item => (item, style: target.Components.GetValueOrDefault(ComponentName(item.Component!))))
            .Where(pair => pair.style is not null && pair.style.Identity != pair.item.Component)
            .Select(pair => new DesignChange
            {
                SlideIndex = pair.item.SlideIndex,
                SlideId = pair.item.SlideId,
                ShapeId = pair.item.ShapeId,
                ShapeName = pair.item.ShapeName,
                Property = "component-version",
                From = pair.item.Component,
                To = pair.style!.Identity,
            })
            .ToList();

    private static string ComponentName(string identity) => identity.Split('@')[0];

    private static string? ComponentVersion(string identity) => identity.Contains('@', StringComparison.Ordinal) ? identity[(identity.IndexOf('@', StringComparison.Ordinal) + 1)..] : null;
}
