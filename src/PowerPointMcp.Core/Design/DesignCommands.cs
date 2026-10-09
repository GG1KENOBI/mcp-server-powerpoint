using System.Text.Json;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Master;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <inheritdoc cref="IDesignCommands"/>
public sealed partial class DesignCommands : IDesignCommands
{
    private const int MaxLimit = 500;

    /// <inheritdoc/>
    public DesignOperationResult ListProfiles(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return new DesignOperationResult { Success = true, Profiles = ProfileStore.List() };
    }

    /// <inheritdoc/>
    public DesignOperationResult GetProfile(IPresentationBatch batch, string profile)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(profile))
            return Fail("profile is required, e.g. theme, default, or a saved profile name.");
        DesignProfile? definition;
        if (profile == ThemeProfile.Name)
        {
            definition = ThemeDefinition(batch, out _);
        }
        else
        {
            try
            {
                definition = ProfileStore.Find(profile);
            }
            catch (JsonException ex)
            {
                return Fail($"Profile '{profile}' is not valid JSON: {ex.Message}");
            }
        }
        if (definition is null)
            return Fail($"Unknown profile '{profile}'. Available: {string.Join(", ", ProfileStore.List().Select(entry => entry.Name))}.");
        var resolved = DesignProfiles.Resolve(definition, ProfileStore.Find);
        return new DesignOperationResult
        {
            Success = resolved.IsValid,
            ErrorMessage = resolved.IsValid ? null : $"Profile '{profile}' is invalid: {string.Join("; ", resolved.Errors)}",
            ProfileName = profile,
            ProfileJson = DesignProfiles.ToJson(definition),
            ResolvedJson = resolved.Resolved?.ToJson(),
            Errors = resolved.Errors.Count > 0 ? resolved.Errors : null,
            Warnings = resolved.Warnings.Count > 0 ? resolved.Warnings : null,
        };
    }

    /// <inheritdoc/>
    public DesignOperationResult ValidateProfile(IPresentationBatch batch, string profileJson)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var (definition, validation, error) = ParseAndResolve(profileJson);
        if (error is not null)
            return Fail(error);
        return new DesignOperationResult
        {
            Success = validation!.IsValid,
            ErrorMessage = validation.IsValid ? null : $"The profile is invalid: {string.Join("; ", validation.Errors)}",
            ProfileName = definition!.Name,
            ResolvedJson = validation.Resolved?.ToJson(),
            Errors = validation.Errors.Count > 0 ? validation.Errors : null,
            Warnings = validation.Warnings.Count > 0 ? validation.Warnings : null,
        };
    }

    /// <inheritdoc/>
    public DesignOperationResult SaveProfile(IPresentationBatch batch, string profileJson, bool overwrite = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var (definition, validation, error) = ParseAndResolve(profileJson);
        if (error is not null)
            return Fail(error);
        return Save(definition!, validation!, overwrite, null);
    }

    /// <inheritdoc/>
    public DesignOperationResult ImportProfile(IPresentationBatch batch, string path, bool overwrite = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return Fail($"The profile file '{path}' does not exist.");
        string json;
        try
        {
            json = File.ReadAllText(path);
        }
        catch (IOException ex)
        {
            return Fail($"The profile file could not be read: {ex.Message}");
        }
        var (definition, validation, error) = ParseAndResolve(json);
        if (error is not null)
            return Fail(error);
        if (string.IsNullOrWhiteSpace(definition!.Name))
            definition.Name = System.IO.Path.GetFileNameWithoutExtension(path);
        return Save(definition, validation!, overwrite, path);
    }

    /// <inheritdoc/>
    public DesignOperationResult ExportProfile(IPresentationBatch batch, string profile, string path, bool resolved = false, bool overwrite = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (string.IsNullOrWhiteSpace(path))
            return Fail("path is required.");
        var current = GetProfile(batch, profile);
        if (!current.Success)
            return current;
        string fullPath;
        try
        {
            fullPath = System.IO.Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return Fail($"Invalid path: {ex.Message}");
        }
        if (File.Exists(fullPath) && !overwrite)
            return Fail($"{fullPath} already exists. Pass overwrite=true to replace it.");
        try
        {
            if (System.IO.Path.GetDirectoryName(fullPath) is { Length: > 0 } folder)
                Directory.CreateDirectory(folder);
            File.WriteAllText(fullPath, resolved ? current.ResolvedJson! : current.ProfileJson!);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fail($"The profile could not be written: {ex.Message}");
        }
        return new DesignOperationResult { Success = true, ProfileName = profile, Path = fullPath, Warnings = current.Warnings };
    }

    /// <inheritdoc/>
    public DesignOperationResult DeleteProfile(IPresentationBatch batch, string profile)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (profile == ThemeProfile.Name || DesignProfiles.IsBuiltIn(profile ?? ""))
            return Fail($"'{profile}' is built in and cannot be deleted.");
        if (!ProfileStore.IsValidName(profile))
            return Fail("Profile names use letters, digits, dash, and underscore.");
        var path = ProfileStore.PathFor(profile!);
        if (!File.Exists(path))
            return Fail($"No saved profile named '{profile}'.");
        File.Delete(path);
        return new DesignOperationResult { Success = true, ProfileName = profile, Path = path };
    }

    /// <inheritdoc/>
    public DesignOperationResult ExtractProfile(IPresentationBatch batch, string name = "extracted")
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (!ProfileStore.IsValidName(name) || DesignProfiles.IsBuiltIn(name) || name == ThemeProfile.Name)
            return Fail("name must use letters, digits, dash, and underscore and must not be a built-in profile name.");
        var masters = new MasterCommands();
        var colors = masters.GetThemeColors(batch, 1);
        var fonts = masters.GetThemeFonts(batch, 1);
        var sample = batch.Execute((ctx, ct) =>
        {
            PowerPoint.PageSetup? setup = null;
            try
            {
                setup = ctx.Presentation.PageSetup;
                var scan = DesignScanner.Scan(ctx.Presentation, null, null, ct);
                return new DeckStyleSample(
                    colors.Success ? colors.ThemeColors : null,
                    fonts.Success ? fonts.MajorThemeFonts?.GetValueOrDefault("latin") : null,
                    fonts.Success ? fonts.MinorThemeFonts?.GetValueOrDefault("latin") : null,
                    setup.SlideWidth,
                    setup.SlideHeight,
                    scan.Runs,
                    scan.Fills,
                    scan.TitleBoxes,
                    scan.ContentBoxes);
            }
            finally
            {
                if (setup is not null) ComUtilities.Release(ref setup);
            }
        });
        var (profile, notes) = ProfileExtractor.Extract(sample, name);
        var validation = DesignProfiles.Resolve(profile, ProfileStore.Find);
        return new DesignOperationResult
        {
            Success = true,
            ProfileName = name,
            ProfileJson = DesignProfiles.ToJson(profile),
            ResolvedJson = validation.Resolved?.ToJson(),
            Errors = validation.Errors.Count > 0 ? validation.Errors : null,
            Warnings = [.. notes, .. validation.Warnings, "Review the draft, then save it with design save-profile."],
        };
    }

    private static DesignOperationResult Save(DesignProfile definition, ProfileValidation validation, bool overwrite, string? importedFrom)
    {
        if (!validation.IsValid)
            return new DesignOperationResult { Success = false, ErrorMessage = $"The profile is invalid: {string.Join("; ", validation.Errors)}", Errors = validation.Errors };
        try
        {
            var path = ProfileStore.Save(definition, overwrite);
            var warnings = validation.Warnings.ToList();
            if (importedFrom is not null)
                warnings.Add($"Imported from {importedFrom}.");
            return new DesignOperationResult { Success = true, ProfileName = definition.Name, Path = path, Warnings = warnings.Count > 0 ? warnings : null };
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fail($"The profile could not be saved: {ex.Message}");
        }
    }

    private static (DesignProfile? Definition, ProfileValidation? Validation, string? Error) ParseAndResolve(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
            return (null, null, "profile_json is required (schema pptmcp.design-profile/1). Use design get-profile profile=default for an example.");
        DesignProfile definition;
        try
        {
            definition = DesignProfiles.Parse(json);
        }
        catch (JsonException ex)
        {
            return (null, null, $"profile_json is not valid JSON: {ex.Message}");
        }
        return (definition, DesignProfiles.Resolve(definition, ProfileStore.Find), null);
    }

    /// <summary>The theme profile definition; call outside Execute.</summary>
    internal static DesignProfile ThemeDefinition(IPresentationBatch batch, out IReadOnlyList<string> themeFonts)
    {
        var masters = new MasterCommands();
        var colors = masters.GetThemeColors(batch, 1);
        var fonts = masters.GetThemeFonts(batch, 1);
        var heading = fonts.Success ? fonts.MajorThemeFonts?.GetValueOrDefault("latin") : null;
        var body = fonts.Success ? fonts.MinorThemeFonts?.GetValueOrDefault("latin") : null;
        themeFonts = new[] { heading, body }.Where(font => !string.IsNullOrWhiteSpace(font)).Select(font => font!).ToList();
        return ThemeProfile.FromTheme(colors.Success ? colors.ThemeColors : null, heading, body).Profile;
    }

    /// <summary>Resolves a profile by name; call outside Execute.</summary>
    internal static (ResolvedProfile? Profile, string? Error) ResolveNamed(IPresentationBatch batch, string? name)
    {
        var (profile, error, _) = ProfileResolver.Resolve(batch, name);
        return (profile, error);
    }

    private static (Func<int, int, bool>? Shapes, Func<int, bool>? Slides, string? Error) Scope(PowerPoint.Presentation presentation, string? selector, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(selector))
            return (null, null, null);
        ObjectSelector parsed;
        try
        {
            parsed = ObjectSelector.Parse(selector);
        }
        catch (ArgumentException ex)
        {
            return (null, null, ex.Message);
        }
        var deck = DeckSnapshotReader.ReadAll(presentation, detailed: true, cancellationToken);
        var chosen = DeckCommands.Select(deck, parsed).Select(item => (item.SlideId, item.ShapeId)).ToHashSet();
        var slides = chosen.Select(key => key.SlideId).ToHashSet();
        var wholeSlides = deck.Where(entry => parsed.MatchesSlide(entry.Slide) && entry.Objects.All(item => chosen.Contains((item.SlideId, item.ShapeId)))).Select(entry => entry.Slide.SlideId).ToHashSet();
        return ((slideId, shapeId) => chosen.Contains((slideId, shapeId)), slideId => slides.Contains(slideId) || wholeSlides.Contains(slideId), null);
    }

    private static DesignOperationResult PlanResult(List<DesignChange> changes, IEnumerable<string> notes, bool applied, int limit, string profileName) => new()
    {
        Success = true,
        ProfileName = profileName,
        Changes = changes.Take(limit).ToList(),
        ChangeCount = changes.Count,
        ChangesByProperty = changes.GroupBy(change => change.Property).OrderBy(group => group.Key, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.Count()),
        Applied = applied,
        Warnings = notes.Distinct().ToList() is { Count: > 0 } list ? list : null,
    };

    internal static DesignOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
