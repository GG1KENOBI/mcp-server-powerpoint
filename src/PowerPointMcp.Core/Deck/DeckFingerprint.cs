using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Content fingerprints over inspection snapshots. A fingerprint changes when an object is added,
/// removed, moved, resized, restyled in a measured property (font sizes), re-tagged, or its text
/// changes. It does not cover properties the snapshot does not read (for example fill colors),
/// so it is a revision check for stale edits, not a byte-level file hash.
/// </summary>
public static class DeckFingerprint
{
    /// <summary>Fingerprint of one slide from its objects (16 hex characters).</summary>
    public static string ForSlide(int slideId, IEnumerable<DeckObjectInfo> objects)
    {
        ArgumentNullException.ThrowIfNull(objects);
        var builder = new StringBuilder();
        builder.Append(slideId.ToString(CultureInfo.InvariantCulture)).Append('\n');
        foreach (var item in objects.OrderBy(item => item.ShapeId))
        {
            builder.Append(Canonical(item)).Append('\n');
        }
        return Hash(builder.ToString());
    }

    /// <summary>Fingerprint of a whole presentation from its slides' fingerprints in order.</summary>
    public static string ForDeck(IEnumerable<(int SlideId, string Fingerprint)> slides)
    {
        ArgumentNullException.ThrowIfNull(slides);
        return Hash(string.Join("\n", slides.Select(slide => $"{slide.SlideId}:{slide.Fingerprint}")));
    }

    /// <summary>The canonical text used for hashing and change detection.</summary>
    public static string Canonical(DeckObjectInfo item)
    {
        ArgumentNullException.ThrowIfNull(item);
        return string.Join('|',
            item.ShapeId.ToString(CultureInfo.InvariantCulture),
            item.Kind,
            item.Name,
            P(item.Left), P(item.Top), P(item.Width), P(item.Height), P(item.Rotation),
            item.ZOrder.ToString(CultureInfo.InvariantCulture),
            item.Visible ? "v" : "h",
            item.Text ?? "",
            item.MinFontSize is { } min ? P(min) : "",
            item.MaxFontSize is { } max ? P(max) : "",
            item.AppId ?? "",
            item.Component ?? "",
            item.GroupShapeId?.ToString(CultureInfo.InvariantCulture) ?? "");
    }

    private static string P(float value) => Math.Round(value, 1).ToString("0.0", CultureInfo.InvariantCulture);

    private static string Hash(string text) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)))[..16].ToLowerInvariant();
}
