namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>
/// Builds a design profile from a deck's theme so compositions match the deck by default
/// (profile name "theme"). Pure: callers read the theme colors and fonts through COM first.
/// </summary>
public static class ThemeProfile
{
    /// <summary>Name of the theme-derived profile.</summary>
    public const string Name = "theme";

    /// <summary>
    /// Maps Office theme roles (Dark1, Light1, Dark2, Light2, Accent1-6) to profile tokens:
    /// primary=Accent1, secondary=Accent5, accent=Accent2, positive=Accent6, neutral=Accent3,
    /// text=Dark1, muted=Dark2, background=Light1, surface=Light2. Missing roles keep the
    /// default profile's values. Returns the profile and notes on what was inferred.
    /// </summary>
    public static (DesignProfile Profile, IReadOnlyList<string> Notes) FromTheme(
        IReadOnlyDictionary<string, string>? themeColors,
        string? headingFont,
        string? bodyFont)
    {
        var notes = new List<string>();
        var colors = new Dictionary<string, string>(StringComparer.Ordinal);
        void Map(string token, string role)
        {
            if (themeColors is not null && themeColors.TryGetValue(role, out var hex) && ResolvedProfile.IsHex(hex))
                colors[token] = hex.ToUpperInvariant();
            else
                notes.Add($"colors.{token}: theme role {role} not available; default kept.");
        }
        Map("primary", "Accent1");
        Map("secondary", "Accent5");
        Map("accent", "Accent2");
        Map("positive", "Accent6");
        Map("neutral", "Accent3");
        Map("text", "Dark1");
        Map("muted", "Dark2");
        Map("background", "Light1");
        Map("surface", "Light2");
        notes.Add("colors.negative and colors.border are not theme roles; the default profile's values are used.");

        var fonts = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!string.IsNullOrWhiteSpace(headingFont))
            fonts["heading"] = headingFont;
        else
            notes.Add("fonts.heading: the theme has no Latin heading font; default kept.");
        if (!string.IsNullOrWhiteSpace(bodyFont))
            fonts["body"] = bodyFont;
        else
            notes.Add("fonts.body: the theme has no Latin body font; default kept.");

        return (new DesignProfile
        {
            Name = Name,
            Version = "1.0.0",
            Description = "Derived from the presentation's theme colors and fonts.",
            Base = "default",
            Colors = colors,
            Fonts = fonts,
        }, notes);
    }
}
