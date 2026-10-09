using System.Runtime.Versioning;

namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Installed font families on this computer (Windows GDI+ family names), cached for five minutes.</summary>
public static class FontCatalog
{
    private static readonly Lock Gate = new();
    private static IReadOnlySet<string>? _cached;
    private static DateTime _cachedAt;

    /// <summary>Installed font family names, or null when they cannot be listed (not Windows).</summary>
    public static IReadOnlySet<string>? InstalledFamilies()
    {
        if (!OperatingSystem.IsWindows())
            return null;
        lock (Gate)
        {
            if (_cached is null || DateTime.UtcNow - _cachedAt > TimeSpan.FromMinutes(5))
            {
                _cached = ReadFamilies();
                _cachedAt = DateTime.UtcNow;
            }
            return _cached;
        }
    }

    [SupportedOSPlatform("windows")]
    private static HashSet<string> ReadFamilies()
    {
        using var fonts = new System.Drawing.Text.InstalledFontCollection();
        return fonts.Families.Select(family => family.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
    }
}
