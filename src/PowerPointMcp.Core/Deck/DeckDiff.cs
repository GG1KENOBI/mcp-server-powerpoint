namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Structural before/after comparison of two inspection snapshots, keyed by PowerPoint SlideID and
/// Shape.Id. Pure: used for batch change logs and agent-visible "what changed" reports.
/// </summary>
public static class DeckDiff
{
    private const float GeometryTolerance = 0.05f;

    /// <summary>Lists added, removed, moved, and modified slides and objects.</summary>
    public static IReadOnlyList<DeckChange> Compare(
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> before,
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> after)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);

        var changes = new List<DeckChange>();
        var beforeById = before.ToDictionary(entry => entry.Slide.SlideId);
        var afterById = after.ToDictionary(entry => entry.Slide.SlideId);

        foreach (var (slide, objects) in after)
        {
            if (!beforeById.TryGetValue(slide.SlideId, out var previous))
            {
                changes.Add(new DeckChange { Change = "slide-added", SlideId = slide.SlideId, SlideIndex = slide.SlideIndex });
                continue;
            }
            if (previous.Slide.SlideIndex != slide.SlideIndex)
            {
                changes.Add(new DeckChange { Change = "slide-moved", SlideId = slide.SlideId, SlideIndex = slide.SlideIndex });
            }
            CompareObjects(slide, previous.Objects, objects, changes);
        }

        foreach (var (slide, _) in before.Where(entry => !afterById.ContainsKey(entry.Slide.SlideId)))
        {
            changes.Add(new DeckChange { Change = "slide-removed", SlideId = slide.SlideId, SlideIndex = slide.SlideIndex });
        }
        return changes;
    }

    private static void CompareObjects(
        DeckSlideInfo slide,
        IReadOnlyList<DeckObjectInfo> before,
        IReadOnlyList<DeckObjectInfo> after,
        List<DeckChange> changes)
    {
        var beforeById = before.ToDictionary(item => item.ShapeId);
        var afterIds = after.Select(item => item.ShapeId).ToHashSet();
        foreach (var item in after)
        {
            if (!beforeById.TryGetValue(item.ShapeId, out var previous))
            {
                changes.Add(Object("object-added", slide, item, null));
                continue;
            }
            var fields = ChangedFields(previous, item);
            if (fields.Count > 0)
                changes.Add(Object("object-modified", slide, item, fields));
        }
        foreach (var item in before.Where(item => !afterIds.Contains(item.ShapeId)))
        {
            changes.Add(Object("object-removed", slide, item, null));
        }
    }

    /// <summary>Names of the snapshot properties that differ between two versions of an object.</summary>
    public static List<string> ChangedFields(DeckObjectInfo before, DeckObjectInfo after)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);
        var fields = new List<string>();
        Check("name", before.Name != after.Name);
        Check("left", Differs(before.Left, after.Left));
        Check("top", Differs(before.Top, after.Top));
        Check("width", Differs(before.Width, after.Width));
        Check("height", Differs(before.Height, after.Height));
        Check("rotation", Differs(before.Rotation, after.Rotation));
        Check("zOrder", before.ZOrder != after.ZOrder);
        Check("visible", before.Visible != after.Visible);
        Check("text", !string.Equals(before.Text, after.Text, StringComparison.Ordinal));
        Check("minFontSize", before.MinFontSize != after.MinFontSize);
        Check("maxFontSize", before.MaxFontSize != after.MaxFontSize);
        Check("appId", before.AppId != after.AppId);
        Check("component", before.Component != after.Component);
        Check("group", before.GroupShapeId != after.GroupShapeId);
        return fields;

        void Check(string name, bool changed)
        {
            if (changed) fields.Add(name);
        }
    }

    private static bool Differs(float a, float b) => Math.Abs(a - b) > GeometryTolerance;

    private static DeckChange Object(string change, DeckSlideInfo slide, DeckObjectInfo item, IReadOnlyList<string>? fields) => new()
    {
        Change = change,
        SlideId = slide.SlideId,
        SlideIndex = slide.SlideIndex,
        ShapeId = item.ShapeId,
        Name = item.Name,
        Fields = fields,
    };
}
