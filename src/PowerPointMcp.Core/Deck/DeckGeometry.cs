using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>Pure slide-geometry helpers. All values are points (1/72 inch).</summary>
public static class DeckGeometry
{
    /// <summary>Aspect ratio label: 16:9, 4:3, 16:10, or the width:height ratio as N:1.</summary>
    public static string AspectRatio(float width, float height)
    {
        if (height <= 0)
            return "unknown";
        var ratio = width / height;
        if (Math.Abs(ratio - (16f / 9f)) < 0.01f) return "16:9";
        if (Math.Abs(ratio - (4f / 3f)) < 0.01f) return "4:3";
        if (Math.Abs(ratio - (16f / 10f)) < 0.01f) return "16:10";
        return $"{ratio.ToString("0.###", CultureInfo.InvariantCulture)}:1";
    }
}
