using System.Reflection;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Catalog;

/// <summary>One command category (an MCP tool and CLI command group).</summary>
public sealed record CategoryInfo(string Category, string? McpTool, string? Title, string? Description, IReadOnlyList<string> Actions, IReadOnlySet<string> ReadOnlyActions);

/// <summary>
/// Every command category and action, read once from the Core interfaces' attributes and the
/// generated registry. Used for batch prevalidation, change tracking (which actions mutate),
/// and capability discovery.
/// </summary>
public static class CommandCatalog
{
    private static readonly Lazy<Dictionary<string, CategoryInfo>> Categories = new(Build);

    /// <summary>All categories by name.</summary>
    public static IReadOnlyDictionary<string, CategoryInfo> All => Categories.Value;

    /// <summary>Whether category.action exists.</summary>
    public static bool Exists(string category, string action) =>
        All.TryGetValue(category, out var info) && info.Actions.Contains(action);

    /// <summary>Whether category.action only reads (listed in McpReadOnlyActions).</summary>
    public static bool IsReadOnly(string category, string action) =>
        All.TryGetValue(category, out var info) && info.ReadOnlyActions.Contains(action);

    /// <summary>Validates raw camelCase service arguments for category.action; returns an error message or null.</summary>
    public static string? ValidateArguments(string category, string action, string? argsJson)
    {
        if (!All.ContainsKey(category))
            return $"Unknown category '{category}'. Categories: {string.Join(", ", All.Keys.Order(StringComparer.Ordinal))}.";
        if (!Exists(category, action))
            return $"Unknown action '{action}' for {category}. Actions: {string.Join(", ", All[category].Actions)}.";
        var registry = RegistryType(category);
        var validate = registry?.GetMethod("ValidateActionArguments", BindingFlags.Public | BindingFlags.Static);
        if (validate is null)
            return null;
        try
        {
            validate.Invoke(null, [action, argsJson]);
            return null;
        }
        catch (TargetInvocationException ex) when (ex.InnerException is ArgumentException or System.Text.Json.JsonException)
        {
            return ex.InnerException!.Message;
        }
    }

    private static Type? RegistryType(string category)
    {
        var registry = typeof(CommandCatalog).Assembly.GetType("Sbroenne.PowerPointMcp.Generated.ServiceRegistry");
        return registry?.GetNestedTypes(BindingFlags.Public).FirstOrDefault(type =>
            type.GetField("Category", BindingFlags.Public | BindingFlags.Static)?.GetValue(null) as string == category);
    }

    private static Dictionary<string, CategoryInfo> Build()
    {
        var result = new Dictionary<string, CategoryInfo>(StringComparer.Ordinal);
        foreach (var type in typeof(CommandCatalog).Assembly.GetTypes().Where(type => type.IsInterface))
        {
            var category = type.GetCustomAttribute<ServiceCategoryAttribute>();
            if (category is null)
                continue;
            var tool = type.GetCustomAttribute<McpToolAttribute>();
            var readOnly = type.GetCustomAttribute<McpReadOnlyActionsAttribute>();
            var actions = RegistryType(category.Category)?.GetField("ValidActions", BindingFlags.Public | BindingFlags.Static)?.GetValue(null) as string[] ?? [];
            result[category.Category] = new CategoryInfo(
                category.Category,
                tool?.ToolName,
                tool?.Title,
                tool?.Description,
                actions,
                (readOnly?.ActionNames ?? []).ToHashSet(StringComparer.Ordinal));
        }
        return result;
    }
}
