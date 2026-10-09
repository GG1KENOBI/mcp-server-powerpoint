using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Plans persistent PPTMCP_ID assignments. Pure: the plan is applied through COM afterwards.
/// Copying a slide or object copies its tags, so the same id can appear twice; the first
/// occurrence in slide/shape order keeps it and later copies get a numbered suffix.
/// </summary>
public static class AppIdPlanner
{
    /// <summary>Ids used more than once (slides and objects are separate namespaces).</summary>
    public static IReadOnlyList<string> FindDuplicates(
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck)
    {
        ArgumentNullException.ThrowIfNull(deck);
        var slideDuplicates = deck.Select(entry => entry.Slide.AppId)
            .Where(id => !string.IsNullOrEmpty(id))
            .GroupBy(id => id!, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .Select(group => $"slide:{group.Key}");
        var objectDuplicates = deck.SelectMany(entry => entry.Objects).Select(item => item.AppId)
            .Where(id => !string.IsNullOrEmpty(id))
            .GroupBy(id => id!, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .Select(group => $"object:{group.Key}");
        return slideDuplicates.Concat(objectDuplicates).ToList();
    }

    /// <summary>
    /// Plans ids for slides without one, repairs duplicate slide and object ids, and (when
    /// <paramref name="objectsToTag"/> is given) assigns ids to those objects.
    /// </summary>
    public static IReadOnlyList<AppIdAssignment> Plan(
        IReadOnlyList<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> deck,
        IReadOnlyCollection<DeckObjectInfo>? objectsToTag,
        string prefix)
    {
        ArgumentNullException.ThrowIfNull(deck);
        ArgumentException.ThrowIfNullOrWhiteSpace(prefix);

        var plan = new List<AppIdAssignment>();
        var usedSlideIds = new HashSet<string>(StringComparer.Ordinal);
        var usedObjectIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (slide, objects) in deck)
        {
            if (!string.IsNullOrEmpty(slide.AppId))
                usedSlideIds.Add(slide.AppId);
            foreach (var item in objects.Where(item => !string.IsNullOrEmpty(item.AppId)))
                usedObjectIds.Add(item.AppId!);
        }

        var seenSlides = new HashSet<string>(StringComparer.Ordinal);
        var seenObjects = new HashSet<string>(StringComparer.Ordinal);
        var tagSet = objectsToTag?.Select(item => (item.SlideId, item.ShapeId)).ToHashSet();
        foreach (var (slide, objects) in deck)
        {
            if (string.IsNullOrEmpty(slide.AppId))
            {
                var id = Unique($"{prefix}-slide-{slide.SlideId.ToString(CultureInfo.InvariantCulture)}", usedSlideIds);
                plan.Add(new AppIdAssignment { SlideId = slide.SlideId, SlideIndex = slide.SlideIndex, NewAppId = id, Reason = "missing" });
            }
            else if (!seenSlides.Add(slide.AppId))
            {
                plan.Add(new AppIdAssignment
                {
                    SlideId = slide.SlideId,
                    SlideIndex = slide.SlideIndex,
                    OldAppId = slide.AppId,
                    NewAppId = Unique(slide.AppId, usedSlideIds),
                    Reason = "duplicate",
                });
            }

            foreach (var item in objects)
            {
                if (!string.IsNullOrEmpty(item.AppId))
                {
                    if (!seenObjects.Add(item.AppId))
                    {
                        plan.Add(new AppIdAssignment
                        {
                            SlideId = slide.SlideId,
                            SlideIndex = slide.SlideIndex,
                            ShapeId = item.ShapeId,
                            OldAppId = item.AppId,
                            NewAppId = Unique(item.AppId, usedObjectIds),
                            Reason = "duplicate",
                        });
                    }
                }
                else if (tagSet is not null && tagSet.Contains((item.SlideId, item.ShapeId)))
                {
                    var id = Unique($"{prefix}-{Slug(item.Role ?? item.Kind)}-{slide.SlideId.ToString(CultureInfo.InvariantCulture)}-{item.ShapeId.ToString(CultureInfo.InvariantCulture)}", usedObjectIds);
                    plan.Add(new AppIdAssignment
                    {
                        SlideId = slide.SlideId,
                        SlideIndex = slide.SlideIndex,
                        ShapeId = item.ShapeId,
                        NewAppId = id,
                        Reason = "missing",
                    });
                }
            }
        }
        return plan;
    }

    private static string Unique(string baseId, HashSet<string> used)
    {
        if (used.Add(baseId))
            return baseId;
        for (int suffix = 2; ; suffix++)
        {
            var candidate = $"{baseId}~{suffix.ToString(CultureInfo.InvariantCulture)}";
            if (used.Add(candidate))
                return candidate;
        }
    }

    private static string Slug(string value) =>
        new string(value.ToLowerInvariant().Select(character => char.IsLetterOrDigit(character) ? character : '-').ToArray()).Trim('-');
}
