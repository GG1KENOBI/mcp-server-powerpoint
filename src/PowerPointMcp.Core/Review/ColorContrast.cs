using System.Globalization;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>WCAG 2.x relative luminance and contrast ratio for #RRGGBB colors. Pure.</summary>
public static class ColorContrast
{
    /// <summary>Parses #RRGGBB (or RRGGBB); returns null when malformed.</summary>
    public static (byte R, byte G, byte B)? Parse(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
            return null;
        var value = hex.Trim().TrimStart('#');
        if (value.Length != 6 || !int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            return null;
        return ((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
    }

    /// <summary>WCAG contrast ratio between two #RRGGBB colors (1 to 21), or null when either is malformed.</summary>
    public static double? Ratio(string? foreground, string? background)
    {
        if (Parse(foreground) is not { } fore || Parse(background) is not { } back)
            return null;
        var a = Luminance(fore);
        var b = Luminance(back);
        var (lighter, darker) = a >= b ? (a, b) : (b, a);
        return (lighter + 0.05) / (darker + 0.05);
    }

    private static double Luminance((byte R, byte G, byte B) color) =>
        (0.2126 * Channel(color.R)) + (0.7152 * Channel(color.G)) + (0.0722 * Channel(color.B));

    private static double Channel(byte value)
    {
        var c = value / 255.0;
        return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
    }
}
