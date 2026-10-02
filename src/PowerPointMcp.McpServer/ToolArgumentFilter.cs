using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.PowerPointMcp.Generated;
using Sbroenne.PowerPointMcp.McpServer.Tools;

namespace Sbroenne.PowerPointMcp.McpServer;

internal static class ToolArgumentFilter
{
    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Wrap(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) =>
        (request, cancellationToken) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (request.MatchedPrimitive is not McpServerTool tool)
                return next(request, cancellationToken);

            try
            {
                var properties = tool.ProtocolTool.InputSchema.GetProperty("properties");
                var metadata = ToolCallNormalizer.GetMetadata(tool.ProtocolTool);
                // Near-miss names (add_text_box, session_id on presentation, sessionId elsewhere)
                // are mapped to the single matching canonical name; canonical calls pass through.
                var (canonicalAction, arguments) = ToolCallNormalizer.Normalize(metadata, request.Params?.Arguments);

                foreach (var (name, value) in arguments)
                {
                    ValidateValueKind(name, value, properties.GetProperty(name));
                }

                var suppliedNames = arguments.Keys
                    .Where(name => name != "action")
                    .ToArray();
                try
                {
                    if (tool.ProtocolTool.Name == "presentation")
                    {
                        PresentationTools.ValidateActionParameterNames(canonicalAction, suppliedNames);
                    }
                    else
                    {
                        ServiceRegistry.ValidateMcpActionParameters(
                            tool.ProtocolTool.Name,
                            canonicalAction,
                            suppliedNames.Where(name => name != "session_id"));
                    }
                }
                catch (ArgumentException ex)
                {
                    throw new ArgumentException(
                        ToolCallNormalizer.DescribeInapplicableParameters(metadata, canonicalAction, ex.Message));
                }

                ToolCallNormalizer.EnsureRequiredParameters(metadata, canonicalAction, arguments);

                if (!ReferenceEquals(arguments, request.Params!.Arguments))
                {
                    request.Params.Arguments = arguments;
                }
            }
            catch (ArgumentException ex)
            {
                return ValueTask.FromResult(PowerPointToolsBase.CreateToolResult(
                    PowerPointToolsBase.SerializeToolError(tool.ProtocolTool.Name, ex),
                    isError: true));
            }

            return next(request, cancellationToken);
        };

    private static void ValidateValueKind(string name, JsonElement value, JsonElement schema)
    {
        if (!schema.TryGetProperty("type", out var type))
            return;

        var types = type.ValueKind == JsonValueKind.Array
            ? type.EnumerateArray().Select(item => item.GetString()).ToArray()
            : [type.GetString()];
        var valid = types.Any(item => item switch
        {
            "null" => value.ValueKind == JsonValueKind.Null,
            "string" => value.ValueKind == JsonValueKind.String,
            "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
            "number" => value.ValueKind == JsonValueKind.Number,
            "array" => value.ValueKind == JsonValueKind.Array,
            "object" => value.ValueKind == JsonValueKind.Object,
            _ => true
        });
        if (!valid)
            throw new ArgumentException($"Parameter '{name}' must have type {string.Join(" or ", types)}.");
    }
}
