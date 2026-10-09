using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>An axis-aligned rectangle in points.</summary>
public readonly record struct Box(float Left, float Top, float Right, float Bottom)
{
    /// <summary>Width in points.</summary>
    public float Width => Right - Left;

    /// <summary>Height in points.</summary>
    public float Height => Bottom - Top;

    /// <summary>Area in square points (0 when empty).</summary>
    public float Area => Math.Max(0, Width) * Math.Max(0, Height);

    /// <summary>Creates a box from left, top, width, height.</summary>
    public static Box FromSize(float left, float top, float width, float height) => new(left, top, left + width, top + height);

    /// <summary>Intersection of two boxes (may be empty: width or height &lt;= 0).</summary>
    public Box Intersect(Box other) =>
        new(Math.Max(Left, other.Left), Math.Max(Top, other.Top), Math.Min(Right, other.Right), Math.Min(Bottom, other.Bottom));

    /// <summary>Whether <paramref name="inner"/> lies inside this box within a tolerance.</summary>
    public bool Contains(Box inner, float tolerance = 1f) =>
        inner.Left >= Left - tolerance && inner.Top >= Top - tolerance &&
        inner.Right <= Right + tolerance && inner.Bottom <= Bottom + tolerance;
}

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

    /// <summary>Whether the rotation is a multiple of 180 degrees (within half a degree).</summary>
    public static bool IsUpright(float rotation)
    {
        var normalized = ((rotation % 180f) + 180f) % 180f;
        return normalized < 0.5f || normalized > 179.5f;
    }

    /// <summary>Visual bounding box of an object, accounting for rotation.</summary>
    public static Box BoundsOf(DeckObjectInfo item)
    {
        ArgumentNullException.ThrowIfNull(item);
        return BoundsOf(item.Left, item.Top, item.Width, item.Height, item.Rotation);
    }

    /// <summary>Visual bounding box of a rectangle rotated about its center.</summary>
    public static Box BoundsOf(float left, float top, float width, float height, float rotation)
    {
        if (IsUpright(rotation))
            return Box.FromSize(left, top, width, height);

        var radians = rotation * MathF.PI / 180f;
        var cos = MathF.Abs(MathF.Cos(radians));
        var sin = MathF.Abs(MathF.Sin(radians));
        var rotatedWidth = (width * cos) + (height * sin);
        var rotatedHeight = (width * sin) + (height * cos);
        var centerX = left + (width / 2);
        var centerY = top + (height / 2);
        return new Box(centerX - (rotatedWidth / 2), centerY - (rotatedHeight / 2), centerX + (rotatedWidth / 2), centerY + (rotatedHeight / 2));
    }

    /// <summary>Formats points compactly with invariant culture ("12.5").</summary>
    public static string P(float value) => value.ToString("0.#", CultureInfo.InvariantCulture);
}
