// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Protocol-level proof that slide inspect/check-layout and table set-data are exposed through the
/// generated tools with the intended contract. No PowerPoint: calls use an unknown session, so a
/// call that passes validation ends at the service's "session not found" result.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public sealed class McpSlideReviewProtocolTests : IAsyncLifetime, IAsyncDisposable
{
    private const string MissingSession = "missing-session";
    private static readonly string[] SampleRows = ["Region|Q1", "North|12"];

    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpSlideReviewProtocolTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "SlideReviewTestClient",
            _cts.Token);
    }

    public async Task DisposeAsync() => await DisposeAsyncCore();

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        await ProgramTransportTestHost.StopAsync(
            _client,
            _clientToServerPipe,
            _serverToClientPipe,
            _serverTask,
            _output);

        _cts.Dispose();
    }

    [Fact]
    public async Task SlideTool_ExposesInspectAndCheckLayout_WithOptionalSlideIndexForCheckLayout()
    {
        var slide = await ToolAsync("slide");
        var properties = slide.JsonSchema.GetProperty("properties");
        var actions = Actions(properties);
        Assert.Contains("inspect", actions);
        Assert.Contains("check-layout", actions);

        var metadata = ToolCallNormalizer.GetMetadata(slide.ProtocolTool);
        Assert.Equal(["session_id", "slide_index"], metadata.RequiredParameters("inspect"));
        Assert.Equal(["session_id"], metadata.RequiredParameters("check-layout"));
        Assert.Equal(["session_id", "slide_index"], metadata.ValidParameters("check-layout"));

        var description = slide.Description;
        Assert.Contains("inspect", description, StringComparison.Ordinal);
        Assert.Contains("check-layout", description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task SlideTool_OutputSchema_DescribesShapesAndLayoutIssues()
    {
        var slide = await ToolAsync("slide");
        var output = Assert.IsType<JsonElement>(slide.ReturnJsonSchema).GetRawText();

        foreach (var field in new[] { "shapes", "layoutIssues", "layoutOk", "slideWidth", "slideHeight", "suggestion", "shapeIndexes", "textHeight", "minFontSize" })
        {
            Assert.Contains($"\"{field}\"", output, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task TableTool_ExposesSetData_WithRowStrings()
    {
        var table = await ToolAsync("table");
        var properties = table.JsonSchema.GetProperty("properties");
        Assert.Contains("set-data", Actions(properties));

        var data = properties.GetProperty("data");
        Assert.Contains("array", data.GetProperty("type").GetRawText(), StringComparison.Ordinal);
        Assert.Contains("string", data.GetProperty("items").GetRawText(), StringComparison.Ordinal);

        var metadata = ToolCallNormalizer.GetMetadata(table.ProtocolTool);
        Assert.Equal(["session_id", "slide_index", "shape_index", "data"], metadata.RequiredParameters("set-data"));
        Assert.Equal(
            ["session_id", "slide_index", "shape_index", "data", "separator", "start_row", "start_column"],
            metadata.ValidParameters("set-data"));
        Assert.Contains("\"cellsWritten\"", Assert.IsType<JsonElement>(table.ReturnJsonSchema).GetRawText(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task CheckLayoutWithoutSlideIndex_PassesValidation()
    {
        var message = await ErrorMessageAsync("slide", new() { ["action"] = "check-layout", ["session_id"] = MissingSession });

        Assert.Equal($"Session '{MissingSession}' not found", message);
    }

    [Fact]
    public async Task InspectWithoutSlideIndex_ListsTheMissingParameter()
    {
        var message = await ErrorMessageAsync("slide", new() { ["action"] = "inspect", ["session_id"] = MissingSession });

        Assert.Contains("slide_index is required for slide action \"inspect\"", message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task SetData_PassesValidation_AndRequiresData()
    {
        var valid = await ErrorMessageAsync("table", new()
        {
            ["action"] = "set-data",
            ["session_id"] = MissingSession,
            ["slide_index"] = 1,
            ["shape_index"] = 1,
            ["data"] = SampleRows,
            ["start_row"] = 1,
        });
        Assert.Equal($"Session '{MissingSession}' not found", valid);

        var missing = await ErrorMessageAsync("table", new()
        {
            ["action"] = "set_data",
            ["session_id"] = MissingSession,
            ["slide_index"] = 1,
            ["shape_index"] = 1,
        });
        Assert.Contains("data is required for table action \"set-data\"", missing, StringComparison.Ordinal);
    }

    [Fact]
    public void ServerInstructions_TeachTheVerifyLoop()
    {
        var instructions = string.Join(' ', _client!.ServerInstructions!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        Assert.Contains("slide(action=inspect, session_id, slide_index)", instructions, StringComparison.Ordinal);
        Assert.Contains("slide(action=check-layout, session_id)", instructions, StringComparison.Ordinal);
        Assert.Contains("until layoutOk is true", instructions, StringComparison.Ordinal);
        Assert.Contains("table(action=set-data", instructions, StringComparison.Ordinal);
    }

    private async Task<McpClientTool> ToolAsync(string name)
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        return Assert.Single(tools, tool => tool.Name == name);
    }

    private static string[] Actions(JsonElement properties) =>
        properties.GetProperty("action").GetProperty("enum").EnumerateArray().Select(value => value.GetString()!).ToArray();

    private async Task<string> ErrorMessageAsync(string tool, Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync(tool, arguments, cancellationToken: _cts.Token);
        Assert.True(result.IsError);
        var message = Assert.IsType<JsonElement>(result.StructuredContent).GetProperty("errorMessage").GetString()!;
        _output.WriteLine(message);
        return message;
    }
}
