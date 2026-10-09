using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Master;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>Resolves a profile name for a presentation: "theme" (from the deck), a built-in, or a saved profile.</summary>
internal static class ProfileResolver
{
    /// <summary>Call outside IPresentationBatch.Execute (it runs its own Execute calls for the theme).</summary>
    internal static (ResolvedProfile? Profile, string? Error, List<string> Warnings) Resolve(IPresentationBatch batch, string? name)
    {
        var warnings = new List<string>();
        name ??= ThemeProfile.Name;
        if (name == ThemeProfile.Name)
        {
            var masters = new MasterCommands();
            var colors = masters.GetThemeColors(batch, 1);
            var fonts = masters.GetThemeFonts(batch, 1);
            var (profile, _) = ThemeProfile.FromTheme(colors.Success ? colors.ThemeColors : null,
                fonts.Success ? fonts.MajorThemeFonts?.GetValueOrDefault("latin") : null,
                fonts.Success ? fonts.MinorThemeFonts?.GetValueOrDefault("latin") : null);
            var themed = DesignProfiles.Resolve(profile);
            return themed.IsValid
                ? (themed.Resolved, null, warnings)
                : (null, $"The deck theme could not be turned into a profile: {string.Join("; ", themed.Errors)}. Pass profile=default.", warnings);
        }

        var definition = ProfileStore.Find(name);
        if (definition is null)
            return (null, $"Unknown profile '{name}'. Available: {string.Join(", ", ProfileStore.List().Select(entry => entry.Name))}.", warnings);
        var result = DesignProfiles.Resolve(definition, ProfileStore.Find);
        if (!result.IsValid)
            return (null, $"Profile '{name}' is invalid: {string.Join("; ", result.Errors)}", warnings);
        warnings.AddRange(result.Warnings);
        return (result.Resolved, null, warnings);
    }
}
