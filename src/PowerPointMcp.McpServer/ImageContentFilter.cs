using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace Sbroenne.PowerPointMcp.McpServer;

/// <summary>
/// Attaches PowerPoint-rendered PNGs from the preview tools as MCP image content, next to the
/// JSON text result (which always carries the file path, objects, findings, and wireframe).
/// Images are skipped when the call passes include_image=false, when PPTMCP_PREVIEW_IMAGES=off
/// (for local models without image input), or when the file is larger than the inline limit.
/// </summary>
internal static class ImageContentFilter
{
    private const long MaxInlineBytes = 4L * 1024 * 1024;

    internal static bool ImagesEnabled =>
        !string.Equals(Environment.GetEnvironmentVariable("PPTMCP_PREVIEW_IMAGES"), "off", StringComparison.OrdinalIgnoreCase);

    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Wrap(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) =>
        async (request, cancellationToken) =>
        {
            var result = await next(request, cancellationToken);
            var tool = request.Params?.Name;
            if (tool is not ("preview" or "preview_read") || result.IsError == true || !ImagesEnabled)
                return result;
            if (request.Params!.Arguments is { } arguments &&
                arguments.TryGetValue("include_image", out var include) && include.ValueKind == JsonValueKind.False)
            {
                return result;
            }

            var text = result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text;
            if (text is null)
                return result;
            string? path;
            try
            {
                using var document = JsonDocument.Parse(text);
                path = document.RootElement.TryGetProperty("imagePath", out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
            }
            catch (JsonException)
            {
                return result;
            }
            if (path is null || !File.Exists(path))
                return result;
            var info = new FileInfo(path);
            if (info.Length > MaxInlineBytes)
            {
                result.Content.Add(new TextContentBlock { Text = $"The image is {info.Length / 1024} KB, above the {MaxInlineBytes / 1024} KB inline limit; open it from imagePath or render at a smaller width." });
                return result;
            }
            var bytes = await File.ReadAllBytesAsync(path, cancellationToken);
            result.Content.Add(ImageContentBlock.FromBytes(bytes, "image/png"));
            return result;
        };
}
