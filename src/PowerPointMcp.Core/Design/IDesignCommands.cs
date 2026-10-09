using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>
/// Design profiles (schema pptmcp.design-profile/1: colors, fonts, type scale, spacing,
/// table/chart/footer/source styles, and versioned component styles), applying a profile to a
/// deck with a migration preview, typography normalization, and reusable components. Profiles
/// are local JSON files; nothing is downloaded.
/// </summary>
[ServiceCategory("design", "Design")]
[McpTool("design", Title = "Design Profiles and Components", Destructive = true, Category = "content",
    Description = "Manage design profiles (colors, fonts, type scale, spacing, table/chart styles, components): list, get, validate, save, import, export, delete, and extract one from the open deck. apply-profile and normalize-typography preview every change (dry_run=true by default) and then restyle text, fills, backgrounds, tables, charts, the theme, and component instances. list-components and update-components keep component instances on the profile's current version.")]
[McpReadOnlyActions("list-profiles", "get-profile", "validate-profile", "extract-profile", "list-components")]
public interface IDesignCommands
{
    /// <summary>Lists the deck theme profile, built-in profiles, and saved profiles with version, description, and file path.</summary>
    DesignOperationResult ListProfiles(IPresentationBatch batch);

    /// <summary>Returns a profile's definition JSON and its resolved values (theme reads the open deck).</summary>
    /// <param name="profile">Profile name: theme (the open deck's theme), a built-in (default, corporate-blue, high-contrast, minimal-mono), or a saved profile.</param>
    DesignOperationResult GetProfile(IPresentationBatch batch, string profile);

    /// <summary>Validates profile JSON (schema, colors, sizes, token references, inheritance) and returns errors, warnings, and the resolved values.</summary>
    /// <param name="profileJson">Profile JSON (schema pptmcp.design-profile/1). Omitted values inherit from base (default "default").</param>
    DesignOperationResult ValidateProfile(IPresentationBatch batch, string profileJson);

    /// <summary>Validates and saves profile JSON to the local profiles folder (PPTMCP_PROFILES_DIR or %LOCALAPPDATA%\PowerPointMcp\profiles).</summary>
    /// <param name="overwrite">Replace an existing file (default false).</param>
    DesignOperationResult SaveProfile(IPresentationBatch batch, string profileJson, bool overwrite = false);

    /// <summary>Reads a profile JSON file, validates it, and saves it to the profiles folder under its name.</summary>
    /// <param name="path">Full path of the JSON file to read or write.</param>
    DesignOperationResult ImportProfile(IPresentationBatch batch, string path, bool overwrite = false);

    /// <summary>Writes a profile to a JSON file: the definition, or with resolved=true every resolved value.</summary>
    /// <param name="resolved">Write resolved values instead of the definition (default false).</param>
    DesignOperationResult ExportProfile(IPresentationBatch batch, string profile, string path, bool resolved = false, bool overwrite = false);

    /// <summary>Deletes a saved profile (built-in profiles cannot be deleted).</summary>
    DesignOperationResult DeleteProfile(IPresentationBatch batch, string profile);

    /// <summary>
    /// Drafts a profile from the open deck: theme colors and fonts, most used title/body/caption
    /// sizes, dominant text color, most used shape fill, and margins from title and content
    /// positions. Returns the JSON and how each value was inferred; review it, then save-profile.
    /// </summary>
    /// <param name="name">Name for the new profile (letters, digits, dash, underscore).</param>
    DesignOperationResult ExtractProfile(IPresentationBatch batch, string name = "extracted");

    /// <summary>
    /// Applies a profile to the deck. Colors that equal a source-profile token become the target's
    /// value for that token (source: from_profile, else each slide's PPTMCP_PROFILE tag, else the
    /// current theme); fonts and type-scale sizes migrate by role; tables and charts are restyled;
    /// the theme colors and fonts of every master are updated; component tags move to the
    /// profile's versions. Dry run by default: returns every planned change. The deck is changed
    /// in memory only; save it under a new name to keep the original.
    /// </summary>
    /// <param name="fromProfile">Profile the deck was built with (default: each slide's PPTMCP_PROFILE tag, else the current theme).</param>
    /// <param name="selector">Only restyle objects this selector matches, e.g. slide:3 or kind:table (the theme is then left unchanged).</param>
    /// <param name="updateTheme">Write the profile colors and fonts into every master's theme (default true; ignored with a selector).</param>
    /// <param name="fonts">Change font families (default true).</param>
    /// <param name="colors">Remap text, fill, outline, and background colors (default true).</param>
    /// <param name="sizes">Change font sizes (default true).</param>
    /// <param name="tables">Restyle tables (default true).</param>
    /// <param name="charts">Restyle charts (default true).</param>
    /// <param name="dryRun">Only return the planned changes (default true).</param>
    /// <param name="limit">Changes listed in the result (1-500, default 100); the counts cover all of them.</param>
    DesignOperationResult ApplyProfile(
        IPresentationBatch batch,
        string profile,
        string? fromProfile = null,
        string? selector = null,
        bool updateTheme = true,
        bool fonts = true,
        bool colors = true,
        bool sizes = true,
        bool tables = true,
        bool charts = true,
        bool dryRun = true,
        int limit = 100);

    /// <summary>
    /// Normalizes typography against a profile (default: the deck theme): headings and titles get
    /// the heading font and other text the body font (symbol fonts are kept), titles and subtitles
    /// get their type-scale size, other sizes snap to the nearest type-scale size within
    /// tolerance, and text below min_font_size is raised. Dry run by default.
    /// </summary>
    /// <param name="minSize">Raise text below the profile's min_font_size (default true).</param>
    /// <param name="tolerance">Largest relative difference a size is snapped across (0-1, default 0.25).</param>
    DesignOperationResult NormalizeTypography(
        IPresentationBatch batch,
        string? profile = null,
        string? selector = null,
        bool fonts = true,
        bool sizes = true,
        bool minSize = true,
        float tolerance = 0.25f,
        bool dryRun = true,
        int limit = 100);

    /// <summary>Lists the profile's component styles (default: the deck theme) with the instances tagged in the deck, their versions, and how many are outdated.</summary>
    DesignOperationResult ListComponents(IPresentationBatch batch, string? profile = null);

    /// <summary>
    /// Brings component instances to the profile's current version: recolors them from the
    /// profile they were built with (from_profile, else the slide tag, else the theme) and updates
    /// their PPTMCP_COMPONENT tags. Dry run by default.
    /// </summary>
    /// <param name="component">Only this component (card, kpi, badge, callout, ...); default all.</param>
    DesignOperationResult UpdateComponents(
        IPresentationBatch batch,
        string? profile = null,
        string? component = null,
        string? fromProfile = null,
        bool dryRun = true,
        int limit = 100);
}
