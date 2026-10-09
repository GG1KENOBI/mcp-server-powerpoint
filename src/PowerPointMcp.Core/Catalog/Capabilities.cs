using System.Runtime.InteropServices;
using Microsoft.Win32;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Diagram;
using Sbroenne.PowerPointMcp.Core.Presentation;

namespace Sbroenne.PowerPointMcp.Core.Catalog;

/// <summary>Detected desktop PowerPoint installation (read from the registry; PowerPoint is not started).</summary>
public sealed record PowerPointInstall(bool Registered, string? ProgId, string? Version, string? Platform, string? Channel, string? Note);

/// <summary>One tool with its actions.</summary>
public sealed record ToolCapability(string Tool, string? Title, int Actions, IReadOnlyList<string> ReadOnlyActions, IReadOnlyList<string> AllActions);

/// <summary>A short recipe: which calls to make, in order, for a common job.</summary>
public sealed record WorkflowRecipe(string Name, string When, IReadOnlyList<string> Steps);

/// <summary>Everything a client needs to plan work without guessing.</summary>
public sealed record CapabilityReport(
    string ServerVersion,
    string Runtime,
    string OperatingSystem,
    PowerPointInstall PowerPoint,
    int OpenSessions,
    IReadOnlyList<ToolCapability>? Tools,
    IReadOnlyDictionary<string, string?> Settings,
    IReadOnlyList<string>? CompositionKinds,
    IReadOnlyList<string>? DiagramTypes,
    IReadOnlyList<string>? Profiles,
    IReadOnlyList<WorkflowRecipe>? Workflows,
    string? LocalUi,
    IReadOnlyList<string> Guarantees);

/// <summary>
/// Builds the capability report: server and PowerPoint versions, tools and their read-only
/// actions, settings from the environment, content catalogs, and workflow recipes. Needs no
/// session and never starts PowerPoint.
/// </summary>
public static class Capabilities
{
    /// <summary>Environment variables the server reads.</summary>
    public static readonly IReadOnlyList<(string Name, string Meaning)> SettingNames =
    [
        ("PPTMCP_LENIENT_ARGUMENTS", "1/true accepts common argument-name variants from small models (camelCase, aliases); default strict."),
        ("PPTMCP_PREVIEW_IMAGES", "off disables PNG image content in preview results (text output stays)."),
        ("PPTMCP_PREVIEW_DIR", "Folder for preview snapshots (default %LOCALAPPDATA%\\PowerPointMcp\\previews)."),
        ("PPTMCP_PROFILES_DIR", "Folder for saved design profiles (default %LOCALAPPDATA%\\PowerPointMcp\\profiles)."),
        ("PPTMCP_ASSETS_DIR", "Folder for the image asset catalog (default %LOCALAPPDATA%\\PowerPointMcp\\assets)."),
        ("PPTMCP_UI", "1 starts the local review UI on 127.0.0.1 (off by default)."),
        ("PPTMCP_UI_PORT", "Port for the local UI (default: a free port)."),
    ];

    /// <summary>Recipes returned by topic=workflows.</summary>
    public static readonly IReadOnlyList<WorkflowRecipe> Workflows =
    [
        new("understand-a-deck", "Before changing an existing presentation.",
            ["presentation open", "deck summary (slides, theme, layouts, revision)", "deck inspect-objects slide_index=N or selector=...", "review validate", "preview snapshot slide_index=N (if the client shows images)"]),
        new("build-slides", "Creating new slides from content.",
            ["design list-profiles (theme = the deck's own style)", "compose kinds", "compose plan spec=... (check warnings, estimated fit)", "compose create spec=...", "review validate", "preview contact-sheet"]),
        new("template-slides", "Using the deck's own layouts and placeholders.",
            ["template list-layouts", "template add-slide layout=\"Title and Content\" content={title, body, picture, notes}", "review validate"]),
        new("data-to-slides", "Tables and charts from CSV or Excel files.",
            ["data preview path=...", "data create-table / data create-chart (bound to the file)", "data refresh after the file changes"]),
        new("diagrams", "Flowcharts, swimlanes, matrices, hub-and-spoke, architecture.",
            ["diagram types", "diagram create spec=...", "diagram inspect (glue check)", "diagram update-node / add-edge / relayout"]),
        new("edit-text-everywhere", "Renames, terminology, number formats across the deck.",
            ["deck find-text find_what=...", "deck replace-text (dry run lists every change)", "deck replace-text dry_run=false expected_count=N"]),
        new("restyle", "Moving a deck to a corporate style.",
            ["design extract-profile or design get-profile", "design apply-profile profile=... (dry run)", "design apply-profile dry_run=false", "design normalize-typography", "review validate"]),
        new("change-template", "Applying another template.",
            ["template preview-template template_path=...", "presentation save-as (keep the original)", "presentation apply-template", "review validate"]),
        new("images", "Placing and checking pictures from local folders.",
            ["asset scan-folder folder=...", "asset search query=...", "asset place / asset replace (never stretched)", "asset inspect (PPI, alt text, links)"]),
        new("bulk-changes", "Many edits with references between steps.",
            ["batch validate operations=[...]", "batch run (checkpoint copy first) or batch start for long jobs", "batch status"]),
        new("repair", "Fixing overflow, overlap, off-slide, inconsistent titles.",
            ["review validate", "review plan-repair", "review repair (only the planned, safe fixes)", "review validate"]),
        new("finish", "Saving without losing the original.",
            ["presentation save-as target_path=new file", "presentation close"]),
    ];

    /// <summary>Statements about what the server does and does not do.</summary>
    public static readonly IReadOnlyList<string> Guarantees =
    [
        "Works offline: no cloud APIs, telemetry, accounts, or downloads; images are local files only.",
        "Uses the installed desktop PowerPoint through COM; text results never require reading images.",
        "Never terminates PowerPoint processes it did not start; macros and security prompts are left to the user.",
        "Bulk changes (replace-text, apply-profile, normalize-typography, repair, batch) preview by default and change the open deck in memory; nothing overwrites the original file unless you save it.",
        "No language model runs inside the server.",
    ];

    /// <summary>Builds the report for a topic: overview (default), tools, workflows, or environment.</summary>
    public static CapabilityReport Build(string topic, string serverVersion, int openSessions, string? localUi)
    {
        bool all = topic == "overview";
        var settings = SettingNames.ToDictionary(setting => setting.Name, setting => Environment.GetEnvironmentVariable(setting.Name), StringComparer.Ordinal);
        var tools = all || topic == "tools"
            ? CommandCatalog.All.Values
                .Where(category => category.McpTool is not null)
                .OrderBy(category => category.McpTool, StringComparer.Ordinal)
                .Select(category => new ToolCapability(category.McpTool!, category.Title, category.Actions.Count,
                    category.ReadOnlyActions.Order(StringComparer.Ordinal).ToList(), topic == "tools" ? category.Actions : []))
                .Prepend(PresentationTool(topic == "tools"))
                .Append(new ToolCapability("capabilities", "Capability discovery (no session)", 1, ["get"], []))
                .ToList()
            : null;
        return new CapabilityReport(
            serverVersion,
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            DetectPowerPoint(),
            openSessions,
            tools,
            settings,
            all ? CompositionKinds.All.Select(kind => kind.Name).ToList() : null,
            all ? DiagramParser.Types : null,
            all ? ProfileStore.List().Select(entry => entry.Name).ToList() : null,
            all || topic == "workflows" ? Workflows : null,
            localUi,
            Guarantees);
    }

    private static ToolCapability PresentationTool(bool withActions)
    {
        var actions = Enum.GetValues<PresentationToolAction>().Select(action => action.ToActionString()).ToList();
        return new ToolCapability("presentation", "Presentation lifecycle (create, open, save-as, close, list)", actions.Count,
            actions.Where(action => action is "list" or "test" || action.StartsWith("get-", StringComparison.Ordinal) || action == "list-tags").ToList(),
            withActions ? actions : []);
    }

    /// <summary>Reads the PowerPoint COM registration and Office version from the registry.</summary>
    public static PowerPointInstall DetectPowerPoint()
    {
        if (!OperatingSystem.IsWindows())
            return new PowerPointInstall(false, null, null, null, null, "Not Windows: desktop PowerPoint automation is unavailable.");
        try
        {
            using var curVer = Registry.ClassesRoot.OpenSubKey(@"PowerPoint.Application\CurVer");
            var progId = curVer?.GetValue(null) as string;
            using var clickToRun = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            var version = clickToRun?.GetValue("VersionToReport") as string;
            var platform = clickToRun?.GetValue("Platform") as string;
            var channel = clickToRun?.GetValue("CDNBaseUrl") is string url ? ChannelFromUrl(url) : null;
            if (version is null && progId is not null)
                version = progId.Split('.').LastOrDefault() + ".0 (MSI or volume install)";
            return new PowerPointInstall(progId is not null, progId, version, platform, channel,
                progId is null ? "PowerPoint.Application is not registered; install desktop PowerPoint (Microsoft 365 or Office 2016+)." : null);
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        {
            return new PowerPointInstall(false, null, null, null, null, $"The registry could not be read: {ex.Message}");
        }
    }

    private static string? ChannelFromUrl(string url)
    {
        var id = url.TrimEnd('/').Split('/').LastOrDefault()?.ToLowerInvariant();
        return id switch
        {
            "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" => "Current",
            "55336b82-a18d-4dd6-b5f6-9e5095c314a6" => "Monthly Enterprise",
            "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" => "Semi-Annual Enterprise",
            "64256afe-f5d9-4f86-8936-8840a6a4f5be" => "Current (Preview)",
            "5440fd1f-7ecb-4221-8110-145efaa6372f" => "Beta",
            _ => null,
        };
    }
}
