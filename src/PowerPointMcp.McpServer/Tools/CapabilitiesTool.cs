using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Service.LocalUi;
using Sbroenne.PowerPointMcp.Service;

namespace Sbroenne.PowerPointMcp.McpServer.Tools;

/// <summary>
/// Capability discovery without a session: server and PowerPoint versions, every tool with its
/// read-only actions, settings, content catalogs, and workflow recipes. Never starts PowerPoint.
/// </summary>
[McpServerToolType]
public static class CapabilitiesTool
{
    /// <summary>Returns what this server can do and how to use it.</summary>
    [McpServerTool(Name = "capabilities", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false, Title = "Capabilities and Workflows", UseStructuredContent = true, OutputSchemaType = typeof(CapabilitiesToolOutputSchema))]
    [Description("What this PowerPoint server can do, without opening a presentation: server and installed PowerPoint version, every tool with its read-only actions, settings (PPTMCP_* environment variables), composition kinds, diagram types, design profiles, and step-by-step workflow recipes. Call it first when planning. Actions: overview, tools, workflows, environment.")]
    public static Task<CallToolResult> Capabilities(
        PowerPointMcpService service,
        [Description("What to return. One of: overview (versions, tools, settings, catalogs, workflows), tools (every tool with all action names), workflows (step-by-step recipes), environment (versions and settings only).")] CapabilitiesAction action = CapabilitiesAction.Overview,
        CancellationToken cancellationToken = default) =>
        PowerPointToolsBase.ExecuteToolActionAsync("capabilities", action.ToString().ToLowerInvariant(), () =>
        {
            var chosen = action.ToString().ToLowerInvariant();
            var report = Core.Catalog.Capabilities.Build(chosen, typeof(CapabilitiesTool).Assembly.GetName().Version?.ToString() ?? "0.0.0", service.SessionCount, LocalUiServer.Current?.Url);
            return PowerPointToolsBase.Serialize(new { success = true, report });
        }, cancellationToken);
}

internal sealed class CapabilitiesToolOutputSchema
{
    public bool Success { get; set; }

    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; set; }

    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
    public Core.Catalog.CapabilityReport? Report { get; set; }
}

/// <summary>What the capabilities tool returns.</summary>
[System.Text.Json.Serialization.JsonConverter(typeof(System.Text.Json.Serialization.JsonStringEnumConverter<CapabilitiesAction>))]
public enum CapabilitiesAction
{
    /// <summary>Versions, tools, settings, content catalogs, and workflows.</summary>
    [System.Text.Json.Serialization.JsonStringEnumMemberName("overview")]
    Overview,

    /// <summary>Every tool with all action names.</summary>
    [System.Text.Json.Serialization.JsonStringEnumMemberName("tools")]
    Tools,

    /// <summary>Workflow recipes.</summary>
    [System.Text.Json.Serialization.JsonStringEnumMemberName("workflows")]
    Workflows,

    /// <summary>Versions and settings only.</summary>
    [System.Text.Json.Serialization.JsonStringEnumMemberName("environment")]
    Environment,
}
