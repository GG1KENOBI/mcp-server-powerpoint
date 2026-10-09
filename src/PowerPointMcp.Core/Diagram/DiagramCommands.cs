extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Diagram;

/// <summary>Diagram commands.</summary>
public sealed class DiagramCommands : IDiagramCommands
{
    private const string SpecTagPrefix = "PPTMCP_DIAGRAM_";
    private const string AreaTagPrefix = "PPTMCP_DIAGRAM_AREA_";
    private const string DiagramTag = "PPTMCP_DIAGRAM";

    /// <inheritdoc/>
    public DiagramOperationResult Types(IPresentationBatch batch) =>
        new() { Success = true, Examples = DiagramParser.Examples };

    /// <inheritdoc/>
    public DiagramOperationResult Create(IPresentationBatch batch, string? spec = null, string? specPath = null, int? slideIndex = null, int? slideId = null, string? profile = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (spec is not null && specPath is not null)
            return Fail("Pass spec or spec_path, not both.");
        string json;
        if (specPath is not null)
        {
            if (!File.Exists(specPath))
                return Fail($"spec_path '{specPath}' does not exist.");
            json = File.ReadAllText(specPath);
        }
        else if (spec is not null)
        {
            json = spec;
        }
        else
        {
            return Fail("Pass the diagram JSON in spec, or a file path in spec_path. Call diagram types for examples.");
        }

        var (parsed, errors, warnings) = DiagramParser.Parse(json);
        if (parsed is null)
        {
            return new DiagramOperationResult
            {
                Success = false,
                ErrorMessage = $"The diagram is invalid ({errors.Count} problem(s)); nothing was created. First: {errors[0]}",
                Errors = errors,
                Warnings = warnings.Count > 0 ? warnings : null,
            };
        }

        var (resolved, profileError, profileWarnings) = ProfileResolver.Resolve(batch, profile ?? parsed.Profile);
        if (profileError is not null)
            return Fail(profileError);
        var notes = new List<string>(warnings);
        notes.AddRange(profileWarnings);

        return batch.Execute((ctx, ct) =>
        {
            if (FindDiagramSlide(ctx.Presentation, parsed.Id) is { } existing)
                return Fail($"A diagram with id '{parsed.Id}' already exists on slide {existing.Index}. Use relayout, the node and edge actions, or another id.");

            PowerPoint.Slide? slide = null;
            try
            {
                slide = TargetSlide(ctx.Presentation, slideIndex, slideId, parsed, resolved!, out var targetError);
                if (slide is null)
                    return Fail(targetError!);
                var area = parsed.Area ?? BodyArea(ctx.Presentation, slide, resolved!);
                return Render(slide, parsed, area, resolved!, notes, ct);
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    /// <inheritdoc/>
    public DiagramOperationResult Inspect(IPresentationBatch batch, string diagramId)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return batch.Execute((ctx, ct) =>
        {
            if (FindDiagramSlide(ctx.Presentation, diagramId) is not { } found)
                return Fail($"No diagram has id '{diagramId}'. Use diagram list.");
            var (spec, _, _) = DiagramParser.Parse(found.Json);
            PowerPoint.Slide? slide = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(ctx.Presentation, found.SlideId)!;
                var objects = DeckSnapshotReader.ReadObjects(slide, found.Index, detailed: true, ct)
                    .Where(item => item.Tags?.GetValueOrDefault(DiagramTag) == diagramId).ToList();
                var nodes = objects.Where(item => item.Tags!.ContainsKey(DeckRoles.NodeTag))
                    .Select(item => new DiagramNodeInfo(item.Tags![DeckRoles.NodeTag][(diagramId.Length + 1)..], item.ShapeId, item.AppId ?? "", item.Text,
                        [item.Left, item.Top, item.Width, item.Height]))
                    .ToList();
                var shapeOfNode = nodes.ToDictionary(node => node.NodeId, node => node.ShapeId, StringComparer.Ordinal);
                var labels = objects.Where(item => item.Tags!.ContainsKey("PPTMCP_EDGE_LABEL")).ToDictionary(item => item.Tags!["PPTMCP_EDGE_LABEL"], item => item.Text, StringComparer.Ordinal);
                var edges = objects.Where(item => item.Tags!.ContainsKey("PPTMCP_EDGE")).Select(item =>
                {
                    var tag = item.Tags!["PPTMCP_EDGE"];
                    var pair = tag[(diagramId.Length + 1)..].Split('>');
                    bool attached = shapeOfNode.TryGetValue(pair[0], out var from) && shapeOfNode.TryGetValue(pair[1], out var to) &&
                        item.ConnectorBeginShapeId == from && item.ConnectorEndShapeId == to;
                    return new DiagramEdgeInfo(pair[0], pair[1], item.ShapeId, attached, labels.GetValueOrDefault(tag));
                }).ToList();
                var detached = edges.Where(edge => !edge.Attached).ToList();
                return new DiagramOperationResult
                {
                    Success = true,
                    DiagramId = diagramId,
                    Type = spec?.Type,
                    SlideIndex = found.Index,
                    SlideId = found.SlideId,
                    Nodes = nodes,
                    Edges = edges,
                    Spec = found.Json,
                    Warnings = detached.Count > 0
                        ? [$"{detached.Count} connector(s) are no longer glued to their nodes ({string.Join(", ", detached.Select(edge => $"{edge.From}→{edge.To}"))}); run diagram relayout to reconnect them."]
                        : null,
                };
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    /// <inheritdoc/>
    public DiagramOperationResult List(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return batch.Execute((ctx, ct) =>
        {
            var diagrams = new List<DiagramSummary>();
            ForEachSlideTags(ctx.Presentation, (index, slideId, tags) =>
            {
                foreach (var (name, value) in tags.Where(tag => tag.Key.StartsWith(SpecTagPrefix, StringComparison.Ordinal) && !tag.Key.StartsWith(AreaTagPrefix, StringComparison.Ordinal)))
                {
                    var (spec, _, _) = DiagramParser.Parse(value);
                    diagrams.Add(new DiagramSummary(spec?.Id ?? name[SpecTagPrefix.Length..], spec?.Type ?? "unknown", index, slideId, spec?.Nodes.Count ?? 0, spec?.Edges.Count ?? 0));
                }
                return false;
            });
            return new DiagramOperationResult { Success = true, Diagrams = diagrams };
        });
    }

    /// <inheritdoc/>
    public DiagramOperationResult UpdateNode(IPresentationBatch batch, string diagramId, string nodeId, string? label = null, string? detail = null, string? tone = null, string? shape = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (label is null && detail is null && tone is null && shape is null)
            return Fail("Pass label, detail, tone, or shape.");
        if (shape is not null && Array.IndexOf(DiagramParser.Shapes, shape) < 0)
            return Fail($"shape must be one of {string.Join(", ", DiagramParser.Shapes)}.");

        if (tone is not null || shape is not null)
        {
            return Modify(batch, diagramId, null, spec =>
            {
                var node = spec.Nodes.FirstOrDefault(item => item.Id == nodeId);
                if (node is null)
                    return (null, $"Diagram '{diagramId}' has no node '{nodeId}'.");
                var updated = node with
                {
                    Label = label ?? node.Label,
                    Detail = detail is null ? node.Detail : detail.Length == 0 ? null : detail,
                    Tone = tone ?? node.Tone,
                    Shape = shape ?? node.Shape,
                };
                return (spec with { Nodes = spec.Nodes.Select(item => item.Id == nodeId ? updated : item).ToList() }, null);
            });
        }

        // Text-only change: edit in place so manual formatting survives.
        return batch.Execute((ctx, ct) =>
        {
            if (FindDiagramSlide(ctx.Presentation, diagramId) is not { } found)
                return Fail($"No diagram has id '{diagramId}'.");
            var (spec, _, _) = DiagramParser.Parse(found.Json);
            var node = spec?.Nodes.FirstOrDefault(item => item.Id == nodeId);
            if (spec is null || node is null)
                return Fail($"Diagram '{diagramId}' has no node '{nodeId}'.");

            PowerPoint.Slide? slide = null;
            PowerPoint.Shape? target = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(ctx.Presentation, found.SlideId)!;
                target = FindTagged(slide, DeckRoles.NodeTag, $"{diagramId}:{nodeId}");
                if (target is null)
                    return Fail($"The shape for node '{nodeId}' is missing from slide {found.Index}; run diagram relayout.");
                TextRuns.WithRange(target, (_, range) =>
                {
                    ReplaceParagraph(range, 1, label ?? node.Label);
                    var newDetail = detail is null ? node.Detail : detail.Length == 0 ? null : detail;
                    SetDetail(range, newDetail);
                    return true;
                });
                var updated = node with { Label = label ?? node.Label, Detail = detail is null ? node.Detail : detail.Length == 0 ? null : detail };
                var newSpec = spec with { Nodes = spec.Nodes.Select(item => item.Id == nodeId ? updated : item).ToList() };
                SetSlideTag(slide, SpecTagPrefix + diagramId.ToUpperInvariant(), DiagramParser.ToJson(newSpec));
                var needed = TextRuns.NeededHeight(target) ?? 0;
                return new DiagramOperationResult
                {
                    Success = true,
                    DiagramId = diagramId,
                    SlideIndex = found.Index,
                    SlideId = found.SlideId,
                    Nodes = [new DiagramNodeInfo(nodeId, target.Id, $"{diagramId}/node-{nodeId}", updated.Label, [target.Left, target.Top, target.Width, target.Height])],
                    Overflow = needed > target.Height + 1
                        ? [$"node-{nodeId}: text needs {needed.ToString("0.#", CultureInfo.InvariantCulture)} pt in a {target.Height.ToString("0.#", CultureInfo.InvariantCulture)} pt box; shorten the label or relayout."]
                        : null,
                };
            }
            finally
            {
                if (target is not null) ComUtilities.Release(ref target);
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    /// <inheritdoc/>
    public DiagramOperationResult AddNode(IPresentationBatch batch, string diagramId, string nodeId, string label, string? detail = null, string? shape = null,
        string? lane = null, string? layer = null, string? quadrant = null, string? connectFrom = null, string? connectTo = null) =>
        Modify(batch, diagramId, null, spec =>
        {
            if (spec.Nodes.Any(node => node.Id == nodeId))
                return (null, $"Diagram '{diagramId}' already has a node '{nodeId}'.");
            var edges = spec.Edges.ToList();
            if (connectFrom is not null)
                edges.Add(new DiagramEdge(connectFrom, nodeId));
            if (connectTo is not null)
                edges.Add(new DiagramEdge(nodeId, connectTo));
            return (spec with
            {
                Nodes = [.. spec.Nodes, new DiagramNode(nodeId, label, detail, shape ?? "process", lane, layer, quadrant)],
                Edges = edges,
            }, null);
        });

    /// <inheritdoc/>
    public DiagramOperationResult RemoveNode(IPresentationBatch batch, string diagramId, string nodeId) =>
        Modify(batch, diagramId, null, spec =>
        {
            if (!spec.Nodes.Any(node => node.Id == nodeId))
                return (null, $"Diagram '{diagramId}' has no node '{nodeId}'.");
            if (spec.Hub == nodeId)
                return (null, $"'{nodeId}' is the hub; a hub-spoke diagram needs its hub.");
            return (spec with
            {
                Nodes = spec.Nodes.Where(node => node.Id != nodeId).ToList(),
                Edges = spec.Edges.Where(edge => edge.From != nodeId && edge.To != nodeId).ToList(),
            }, null);
        });

    /// <inheritdoc/>
    public DiagramOperationResult AddEdge(IPresentationBatch batch, string diagramId, string fromNode, string toNode, string? edgeLabel = null) =>
        Modify(batch, diagramId, null, spec =>
            (spec with { Edges = [.. spec.Edges, new DiagramEdge(fromNode, toNode, edgeLabel)] }, null));

    /// <inheritdoc/>
    public DiagramOperationResult RemoveEdge(IPresentationBatch batch, string diagramId, string fromNode, string toNode) =>
        Modify(batch, diagramId, null, spec =>
            !spec.Edges.Any(edge => edge.From == fromNode && edge.To == toNode)
                ? (null, $"Diagram '{diagramId}' has no edge {fromNode}→{toNode}.")
                : (spec with { Edges = spec.Edges.Where(edge => !(edge.From == fromNode && edge.To == toNode)).ToList() }, null));

    /// <inheritdoc/>
    public DiagramOperationResult Relayout(IPresentationBatch batch, string diagramId, string? profile = null) =>
        Modify(batch, diagramId, profile, spec => (spec, null));

    /// <summary>Applies a change to the stored spec, validates it, deletes the diagram's shapes, and renders again in the same area.</summary>
    private static DiagramOperationResult Modify(IPresentationBatch batch, string diagramId, string? profile, Func<DiagramSpec, (DiagramSpec? Spec, string? Error)> change)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var stored = batch.Execute((ctx, ct) => FindDiagramSlide(ctx.Presentation, diagramId));
        if (stored is null)
            return Fail($"No diagram has id '{diagramId}'. Use diagram list.");
        var (current, _, _) = DiagramParser.Parse(stored.Json);
        if (current is null)
            return Fail($"The stored JSON of diagram '{diagramId}' is invalid; recreate the diagram.");
        var (next, error) = change(current);
        if (error is not null)
            return Fail(error);
        // Re-validate through the parser so structural edits obey the same rules as create.
        var (validated, errors, warnings) = DiagramParser.Parse(DiagramParser.ToJson(next!));
        if (validated is null)
            return new DiagramOperationResult { Success = false, ErrorMessage = $"The change would make the diagram invalid; nothing was changed. {errors[0]}", Errors = errors };

        var (resolved, profileError, profileWarnings) = ProfileResolver.Resolve(batch, profile ?? validated.Profile);
        if (profileError is not null)
            return Fail(profileError);
        var notes = new List<string>(warnings);
        notes.AddRange(profileWarnings);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slide? slide = null;
            try
            {
                slide = DeckShapeLocator.FindSlide(ctx.Presentation, stored.SlideId);
                if (slide is null)
                    return Fail($"The slide of diagram '{diagramId}' no longer exists.");
                int removed = DeleteDiagramShapes(slide, diagramId);
                notes.Add($"Relayout recreated the diagram ({removed} shapes replaced). Node ids are unchanged; shape ids are new, and manual formatting of diagram shapes is not kept.");
                var area = stored.Area ?? validated.Area ?? BodyArea(ctx.Presentation, slide, resolved!);
                return Render(slide, validated, area, resolved!, notes, ct);
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
            }
        });
    }

    private static DiagramOperationResult Render(PowerPoint.Slide slide, DiagramSpec spec, Box area, ResolvedProfile profile, List<string> notes, CancellationToken ct)
    {
        var plan = DiagramLayout.Layout(spec, area, profile);
        Dictionary<string, PowerPoint.Shape>? shapes = null;
        var warnings = new List<string>(notes);
        try
        {
            shapes = SlideRenderer.CreateElements(slide, plan.Elements, key => $"{spec.Id}/{key}", warnings, ct);
            var overflow = SlideRenderer.FitIndividually(plan.TextKeys, shapes, Math.Min(profile.MinFontSize, profile.Size("caption")), canShrink: true);
            SetSlideTag(slide, SpecTagPrefix + spec.Id.ToUpperInvariant(), DiagramParser.ToJson(spec));
            SetSlideTag(slide, AreaTagPrefix + spec.Id.ToUpperInvariant(),
                string.Join(',', new[] { area.Left, area.Top, area.Width, area.Height }.Select(value => value.ToString("0.##", CultureInfo.InvariantCulture))));

            var nodes = spec.Nodes.Where(node => shapes.ContainsKey(plan.NodeKeys[node.Id])).Select(node =>
            {
                var shape = shapes[plan.NodeKeys[node.Id]];
                return new DiagramNodeInfo(node.Id, shape.Id, $"{spec.Id}/{plan.NodeKeys[node.Id]}", node.Label, [shape.Left, shape.Top, shape.Width, shape.Height]);
            }).ToList();
            var edges = plan.Elements.Where(element => element.Type == "connector" && shapes.ContainsKey(element.Key)).Select(element =>
            {
                var tag = element.Tags!["PPTMCP_EDGE"];
                var pair = tag[(spec.Id.Length + 1)..].Split('>');
                return new DiagramEdgeInfo(pair[0], pair[1], shapes[element.Key].Id, Attached: true,
                    spec.Edges.FirstOrDefault(edge => edge.From == pair[0] && edge.To == pair[1])?.Label);
            }).ToList();
            return new DiagramOperationResult
            {
                Success = true,
                DiagramId = spec.Id,
                Type = spec.Type,
                SlideIndex = slide.SlideIndex,
                SlideId = slide.SlideID,
                Nodes = nodes,
                Edges = edges,
                Notes = plan.Notes.Count > 0 ? plan.Notes : null,
                Overflow = overflow.Count > 0 ? overflow : null,
                Warnings = warnings.Count > 0 ? warnings : null,
            };
        }
        finally
        {
            if (shapes is not null)
            {
                foreach (var shape in shapes.Values)
                {
                    var item = shape;
                    ComUtilities.Release(ref item!);
                }
            }
        }
    }

    private static PowerPoint.Slide? TargetSlide(PowerPoint.Presentation presentation, int? slideIndex, int? slideId, DiagramSpec spec, ResolvedProfile profile, out string? error)
    {
        error = null;
        if (slideId is { } id)
        {
            var slide = DeckShapeLocator.FindSlide(presentation, id);
            if (slide is null)
                error = $"No slide has slide_id {id}.";
            return slide;
        }
        if (slideIndex is { } index)
        {
            PowerPoint.Slides? slides = null;
            try
            {
                slides = presentation.Slides;
                if (index < 1 || index > slides.Count)
                {
                    error = $"Slide index {index} is out of range. The presentation has {slides.Count} slide(s) (valid range: 1-{slides.Count}).";
                    return null;
                }
                return slides[index];
            }
            finally
            {
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        }

        // New slide after the last one, with the diagram title in the title placeholder.
        var layouts = LayoutPicker.Pick(presentation);
        var created = LayoutPicker.AddSlide(presentation, int.MaxValue, layouts.TitleOnlyIndex, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
        PowerPoint.Shapes? shapes = null;
        PowerPoint.Shape? title = null;
        try
        {
            shapes = created.Shapes;
            if (shapes.HasTitle == Office.MsoTriState.msoTrue)
            {
                title = shapes.Title;
                TextRuns.WithRange(title, (_, range) =>
                {
                    range.Text = spec.Title ?? spec.Id;
                    PowerPoint.Font? font = null;
                    try
                    {
                        font = range.Font;
                        font.Name = profile.HeadingFont;
                    }
                    finally
                    {
                        if (font is not null) ComUtilities.Release(ref font);
                    }
                    return true;
                });
            }
        }
        finally
        {
            if (title is not null) ComUtilities.Release(ref title);
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
        return created;
    }

    /// <summary>The body area: below the slide's title (if any) inside the profile margins, above the footer area.</summary>
    private static Box BodyArea(PowerPoint.Presentation presentation, PowerPoint.Slide slide, ResolvedProfile profile)
    {
        PowerPoint.PageSetup? setup = null;
        PowerPoint.Shapes? shapes = null;
        PowerPoint.Shape? title = null;
        try
        {
            setup = presentation.PageSetup;
            var frame = LayoutEngine.Frame(profile, setup.SlideWidth, setup.SlideHeight, hasTitle: true);
            shapes = slide.Shapes;
            float top = frame.Safe.Top;
            if (shapes.HasTitle == Office.MsoTriState.msoTrue)
            {
                title = shapes.Title;
                top = title.Top + title.Height + frame.Gap;
            }
            return LayoutEngine.R(new Box(frame.Body.Left, Math.Min(top, frame.Body.Bottom - 100), frame.Body.Right, frame.Body.Bottom));
        }
        finally
        {
            if (title is not null) ComUtilities.Release(ref title);
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (setup is not null) ComUtilities.Release(ref setup);
        }
    }

    private sealed record StoredDiagram(int Index, int SlideId, string Json, Box? Area);

    private static StoredDiagram? FindDiagramSlide(PowerPoint.Presentation presentation, string diagramId)
    {
        StoredDiagram? found = null;
        var specTag = SpecTagPrefix + diagramId.ToUpperInvariant();
        var areaTag = AreaTagPrefix + diagramId.ToUpperInvariant();
        ForEachSlideTags(presentation, (index, slideId, tags) =>
        {
            if (!tags.TryGetValue(specTag, out var json))
                return false;
            Box? area = null;
            if (tags.TryGetValue(areaTag, out var text) && text.Split(',') is { Length: 4 } parts &&
                parts.All(part => float.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out _)))
            {
                var values = parts.Select(part => float.Parse(part, CultureInfo.InvariantCulture)).ToArray();
                area = Box.FromSize(values[0], values[1], values[2], values[3]);
            }
            found = new StoredDiagram(index, slideId, json, area);
            return true;
        });
        return found;
    }

    /// <summary>Visits each slide's tags until the visitor returns true.</summary>
    private static void ForEachSlideTags(PowerPoint.Presentation presentation, Func<int, int, Dictionary<string, string>, bool> visit)
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
                    if (visit(index, slide.SlideID, DeckSnapshotReader.ReadTags(slide.Tags)))
                        return;
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
    }

    private static PowerPoint.Shape? FindTagged(PowerPoint.Slide slide, string tagName, string value)
    {
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                if (DeckSnapshotReader.ReadTags(shape.Tags).GetValueOrDefault(tagName) == value)
                    return shape;
                ComUtilities.Release(ref shape);
            }
            return null;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    /// <summary>Deletes every shape tagged with the diagram id (only this diagram's own shapes).</summary>
    private static int DeleteDiagramShapes(PowerPoint.Slide slide, string diagramId)
    {
        PowerPoint.Shapes? shapes = null;
        int removed = 0;
        try
        {
            shapes = slide.Shapes;
            for (int index = shapes.Count; index >= 1; index--)
            {
                var shape = shapes[index];
                try
                {
                    if (DeckSnapshotReader.ReadTags(shape.Tags).GetValueOrDefault(DiagramTag) == diagramId)
                    {
                        shape.Delete();
                        removed++;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shape);
                }
            }
            return removed;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    private static void ReplaceParagraph(PowerPoint.TextRange range, int paragraphIndex, string text)
    {
        PowerPoint.TextRange? paragraph = null;
        PowerPoint.TextRange? characters = null;
        try
        {
            paragraph = range.Paragraphs(paragraphIndex, 1);
            var current = paragraph.Text.TrimEnd('\r');
            characters = paragraph.Characters(1, current.Length);
            // Replacing characters keeps the formatting of the first replaced character.
            characters.Text = text.Replace('\n', '\v');
        }
        finally
        {
            if (characters is not null) ComUtilities.Release(ref characters);
            if (paragraph is not null) ComUtilities.Release(ref paragraph);
        }
    }

    private static void SetDetail(PowerPoint.TextRange range, string? detail)
    {
        PowerPoint.TextRange? paragraphs = null;
        try
        {
            paragraphs = range.Paragraphs();
            int count = paragraphs.Count;
            if (detail is null)
            {
                if (count >= 2)
                {
                    // Remove the paragraph mark before the detail and the detail itself.
                    PowerPoint.TextRange? first = null;
                    PowerPoint.TextRange? tail = null;
                    try
                    {
                        first = range.Paragraphs(1, 1);
                        var label = first.Text.TrimEnd('\r');
                        tail = range.Characters(label.Length + 1, range.Text.Length - label.Length);
                        tail.Delete();
                    }
                    finally
                    {
                        if (tail is not null) ComUtilities.Release(ref tail);
                        if (first is not null) ComUtilities.Release(ref first);
                    }
                }
                return;
            }
            if (count >= 2)
            {
                ReplaceParagraph(range, 2, detail);
                return;
            }
            PowerPoint.TextRange? added = null;
            PowerPoint.Font? font = null;
            try
            {
                added = range.InsertAfter("\r" + detail.Replace('\n', '\v'));
                font = added.Font;
                font.Bold = Office.MsoTriState.msoFalse;
                font.Size = Math.Max(8f, font.Size * 0.8f);
            }
            finally
            {
                if (font is not null) ComUtilities.Release(ref font);
                if (added is not null) ComUtilities.Release(ref added);
            }
        }
        finally
        {
            if (paragraphs is not null) ComUtilities.Release(ref paragraphs);
        }
    }

    private static void SetSlideTag(PowerPoint.Slide slide, string name, string value)
    {
        var tags = slide.Tags;
        try
        {
            tags.Add(name, value);
        }
        finally
        {
            ComUtilities.Release(ref tags);
        }
    }

    private static DiagramOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
