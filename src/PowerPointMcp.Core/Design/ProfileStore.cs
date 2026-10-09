using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>One profile known to the store.</summary>
public sealed record ProfileEntry(string Name, string Source, string? Version, string? Description, string? Path);

/// <summary>
/// User design profiles as JSON files in a local folder: PPTMCP_PROFILES_DIR when set, otherwise
/// %LOCALAPPDATA%\PowerPointMcp\profiles. Nothing is downloaded or synchronized.
/// </summary>
public static partial class ProfileStore
{
    /// <summary>The profiles folder.</summary>
    public static string Directory =>
        Environment.GetEnvironmentVariable("PPTMCP_PROFILES_DIR") is { Length: > 0 } custom
            ? custom
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PowerPointMcp", "profiles");

    /// <summary>Lists built-in and user profiles (user profiles may not shadow built-ins).</summary>
    public static IReadOnlyList<ProfileEntry> List()
    {
        var entries = DesignProfiles.BuiltInNames
            .Select(name => DesignProfiles.GetBuiltIn(name)!)
            .Select(profile => new ProfileEntry(profile.Name!, "built-in", profile.Version, profile.Description, null))
            .ToList();
        entries.Insert(0, new ProfileEntry(ThemeProfile.Name, "presentation", null, "Derived from the open presentation's theme colors and fonts (the default).", null));
        if (System.IO.Directory.Exists(Directory))
        {
            foreach (var file in System.IO.Directory.EnumerateFiles(Directory, "*.json").Order(StringComparer.OrdinalIgnoreCase))
            {
                var name = System.IO.Path.GetFileNameWithoutExtension(file);
                if (!IsValidName(name) || DesignProfiles.IsBuiltIn(name) || name == ThemeProfile.Name)
                    continue;
                try
                {
                    var profile = DesignProfiles.Parse(File.ReadAllText(file));
                    entries.Add(new ProfileEntry(name, "user", profile.Version, profile.Description, file));
                }
                catch (System.Text.Json.JsonException ex)
                {
                    entries.Add(new ProfileEntry(name, "user (invalid JSON)", null, ex.Message, file));
                }
            }
        }
        return entries;
    }

    /// <summary>Loads a user profile by name, or null when it does not exist.</summary>
    public static DesignProfile? Load(string name)
    {
        if (!IsValidName(name))
            return null;
        var path = PathFor(name);
        return File.Exists(path) ? DesignProfiles.Parse(File.ReadAllText(path)) : null;
    }

    /// <summary>Finds a profile definition: built-in first, then user profiles.</summary>
    public static DesignProfile? Find(string name) => DesignProfiles.GetBuiltIn(name) ?? Load(name);

    /// <summary>Saves a user profile; refuses built-in names and existing files unless <paramref name="overwrite"/>.</summary>
    public static string Save(DesignProfile profile, bool overwrite)
    {
        ArgumentNullException.ThrowIfNull(profile);
        var name = profile.Name ?? throw new ArgumentException("The profile has no name.");
        if (!IsValidName(name))
            throw new ArgumentException("Profile names use letters, digits, dash, and underscore (max 64).");
        if (DesignProfiles.IsBuiltIn(name) || name == ThemeProfile.Name)
            throw new ArgumentException($"'{name}' is a built-in profile name; choose another name.");
        var path = PathFor(name);
        if (File.Exists(path) && !overwrite)
            throw new ArgumentException($"A profile named '{name}' already exists at {path}. Pass overwrite=true to replace it.");
        System.IO.Directory.CreateDirectory(Directory);
        File.WriteAllText(path, DesignProfiles.ToJson(profile));
        return path;
    }

    /// <summary>The file path for a user profile name.</summary>
    public static string PathFor(string name) => System.IO.Path.Combine(Directory, name + ".json");

    /// <summary>Whether a name is safe to use as a file name.</summary>
    public static bool IsValidName(string? name) => name is not null && SafeName().IsMatch(name);

    [GeneratedRegex("^[A-Za-z0-9_-]{1,64}$", RegexOptions.CultureInvariant)]
    private static partial Regex SafeName();
}
