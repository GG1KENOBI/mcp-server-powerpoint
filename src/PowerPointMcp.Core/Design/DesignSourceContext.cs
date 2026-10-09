using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>Resolves the profile a slide was built with; created outside Execute.</summary>
internal sealed class DesignSourceContext
{
    private readonly ResolvedProfile? _from;
    private readonly Dictionary<string, ResolvedProfile?> _byName = new(StringComparer.OrdinalIgnoreCase);

    private DesignSourceContext(ResolvedProfile theme, IReadOnlyList<string> themeFonts, ResolvedProfile? from)
    {
        Theme = theme;
        ThemeFonts = themeFonts;
        _from = from;
    }

    public ResolvedProfile Theme { get; }

    public IReadOnlyList<string> ThemeFonts { get; }

    public static DesignSourceContext Create(IPresentationBatch batch, string? fromProfile, out string? error)
    {
        error = null;
        var definition = DesignCommands.ThemeDefinition(batch, out var themeFonts);
        var theme = DesignProfiles.Resolve(definition).Resolved ?? DesignProfiles.ResolveBuiltIn("default");
        ResolvedProfile? from = null;
        if (!string.IsNullOrWhiteSpace(fromProfile))
        {
            (from, error) = DesignCommands.ResolveNamed(batch, fromProfile);
        }
        return new DesignSourceContext(theme, themeFonts, from);
    }

    /// <summary>from_profile, else the slide's PPTMCP_PROFILE (name@version) when it resolves, else the theme.</summary>
    public ResolvedProfile SourceFor(string? profileTag)
    {
        if (_from is not null)
            return _from;
        var name = profileTag?.Split('@')[0];
        if (string.IsNullOrEmpty(name) || name == ThemeProfile.Name)
            return Theme;
        if (!_byName.TryGetValue(name, out var resolved))
        {
            try
            {
                var definition = ProfileStore.Find(name);
                resolved = definition is null ? null : DesignProfiles.Resolve(definition, ProfileStore.Find).Resolved;
            }
            catch (System.Text.Json.JsonException)
            {
                resolved = null;
            }
            _byName[name] = resolved;
        }
        return resolved ?? Theme;
    }
}
