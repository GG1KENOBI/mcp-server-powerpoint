using System.Text.RegularExpressions;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Deck;

public sealed partial class DeckCommands
{
    /// <inheritdoc/>
    public DeckOperationResult FindText(
        IPresentationBatch batch,
        string findWhat,
        string? selector = null,
        bool matchCase = false,
        bool wholeWords = false,
        bool useRegex = false,
        bool includeNotes = false,
        int limit = 100)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return SearchDeckText(batch, findWhat, null, selector, matchCase, wholeWords, useRegex, includeNotes, dryRun: true, expectedCount: null, limit);
    }

    /// <inheritdoc/>
    public DeckOperationResult ReplaceText(
        IPresentationBatch batch,
        string findWhat,
        string replaceWhat,
        string? selector = null,
        bool matchCase = false,
        bool wholeWords = false,
        bool useRegex = false,
        bool includeNotes = false,
        bool dryRun = true,
        int? expectedCount = null,
        int limit = 100)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (replaceWhat is null)
            return Fail("replace_what is required (use an empty string to delete the matches).");
        if (expectedCount is < 0)
            return Fail("expected_count must be 0 or greater.");
        return SearchDeckText(batch, findWhat, replaceWhat, selector, matchCase, wholeWords, useRegex, includeNotes, dryRun, expectedCount, limit);
    }

    private static DeckOperationResult SearchDeckText(
        IPresentationBatch batch,
        string findWhat,
        string? replaceWhat,
        string? selector,
        bool matchCase,
        bool wholeWords,
        bool useRegex,
        bool includeNotes,
        bool dryRun,
        int? expectedCount,
        int limit)
    {
        if (limit is < 1 or > MaxPageSize)
            return Fail($"limit must be between 1 and {MaxPageSize}.");

        Regex matcher;
        ObjectSelector? parsed;
        try
        {
            matcher = TextSearch.Compile(findWhat, matchCase, wholeWords, useRegex);
            parsed = string.IsNullOrWhiteSpace(selector) ? null : ObjectSelector.Parse(selector);
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }

        // PowerPoint separates paragraphs with CR; accept any newline style in the replacement.
        var replacement = replaceWhat?.Replace("\r\n", "\r", StringComparison.Ordinal).Replace('\n', '\r');

        return batch.Execute((ctx, ct) =>
        {
            Func<int, int, bool>? includeShape = null;
            Func<int, bool>? includeNotesOf = null;
            if (parsed is not null)
            {
                var deck = DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: true, ct);
                var chosen = Select(deck, parsed).Select(item => (item.SlideId, item.ShapeId)).ToHashSet();
                var slidesWithMatches = chosen.Select(key => key.SlideId).ToHashSet();
                var slidesMatched = deck.Where(entry => parsed.MatchesSlide(entry.Slide)).Select(entry => entry.Slide.SlideId).ToHashSet();
                includeShape = (slideId, shapeId) => chosen.Contains((slideId, shapeId));
                includeNotesOf = slideId => slidesWithMatches.Contains(slideId) && slidesMatched.Contains(slideId);
            }

            var hits = new List<DeckTextHit>();
            int searched = 0;
            DeckTextWalker.Visit(ctx.Presentation, includeNotes, includeShape, includeNotesOf, (location, range) =>
            {
                searched++;
                var text = range.Text;
                foreach (var hit in TextSearch.Find(text, matcher, replacement, useRegex))
                {
                    var replaced = hit.Replacement?.Replace("\r\n", "\r", StringComparison.Ordinal).Replace('\n', '\r');
                    hits.Add(new DeckTextHit
                    {
                        SlideIndex = location.SlideIndex,
                        SlideId = location.SlideId,
                        ShapeId = location.ShapeId,
                        ShapeName = location.ShapeName,
                        AppId = location.AppId,
                        Row = location.Row,
                        Column = location.Column,
                        InNotes = location.InNotes ? true : null,
                        Start = hit.Start + 1,
                        Length = hit.Length,
                        Matched = hit.Matched,
                        Replacement = replaced,
                        Context = TextSearch.Context(text, hit, replacement: replaced),
                    });
                }
            }, ct);

            var warnings = new List<string>();
            bool apply = replacement is not null && !dryRun;
            if (replacement is not null && expectedCount is { } expected && expected != hits.Count)
            {
                return new DeckOperationResult
                {
                    Success = false,
                    ErrorMessage = $"Found {hits.Count} match(es) but expected_count is {expected}; nothing was changed. Check text_hits and refine find_what or selector.",
                    TextHits = hits.Take(limit).ToList(),
                    TotalCount = hits.Count,
                    TextFramesSearched = searched,
                };
            }

            int replaced = 0;
            if (apply && hits.Count > 0)
                replaced = ApplyReplacements(ctx.Presentation, hits, ct);
            if (replacement is not null && dryRun)
                warnings.Add("Dry run: nothing was changed. Pass dry_run=false to apply these replacements.");
            if (apply && replaced > 0)
                warnings.Add("The open presentation was changed but not saved; save it under a new name (file save-as) to keep the original file.");

            return new DeckOperationResult
            {
                Success = true,
                TextHits = hits.Take(limit).ToList(),
                TotalCount = hits.Count,
                Offset = 0,
                HasMore = hits.Count > limit,
                ReplacementCount = replacement is null ? null : replaced,
                TextFramesSearched = searched,
                Warnings = warnings.Count > 0 ? warnings : null,
            };
        });
    }

    /// <summary>
    /// Replaces planned hits, last to first inside each text block so earlier positions stay valid.
    /// Re-checks each hit's text before writing it so a stale plan never overwrites other text.
    /// </summary>
    private static int ApplyReplacements(PowerPoint.Presentation presentation, List<DeckTextHit> hits, CancellationToken cancellationToken)
    {
        var byBlock = hits
            .GroupBy(hit => (hit.SlideId, hit.ShapeId, hit.Row, hit.Column, hit.InNotes))
            .ToDictionary(group => group.Key, group => group.OrderByDescending(hit => hit.Start).ToList());
        int replaced = 0;
        DeckTextWalker.Visit(presentation, includeNotes: hits.Any(hit => hit.InNotes == true),
            (slideId, shapeId) => byBlock.Keys.Any(key => key.SlideId == slideId && key.ShapeId == shapeId),
            slideId => byBlock.Keys.Any(key => key.SlideId == slideId && key.InNotes == true),
            (location, range) =>
            {
                var key = (location.SlideId, location.ShapeId, location.Row, location.Column, location.InNotes ? true : (bool?)null);
                if (!byBlock.TryGetValue(key, out var planned))
                    return;
                foreach (var hit in planned)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    PowerPoint.TextRange? target = null;
                    try
                    {
                        target = range.Characters(hit.Start, hit.Length);
                        if (!string.Equals(target.Text, hit.Matched, StringComparison.Ordinal))
                            throw new InvalidOperationException($"Text on slide {hit.SlideIndex} shape {hit.ShapeId} changed while replacing; {replaced} replacement(s) were already made.");
                        target.Text = hit.Replacement!;
                        replaced++;
                    }
                    finally
                    {
                        if (target is not null) ComUtilities.Release(ref target);
                    }
                }
            }, cancellationToken);
        return replaced;
    }
}
