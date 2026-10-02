// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using ModelContextProtocol.Client;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Real-PowerPoint proof that normalized near-miss calls (as small/local models write them) do the
/// same work as canonical calls: create, add a slide, add a text box, read it back, save and close.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpToolCallNormalizationRoundTripTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir = Directory.CreateTempSubdirectory("McpNormalizationTest_").FullName;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpToolCallNormalizationRoundTripTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "NormalizationRoundTripClient",
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
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch (IOException)
        {
            // Best effort: PowerPoint may still hold the file briefly.
        }
    }

    [Fact]
    public async Task NearMissCalls_CreateEditSaveAndClose()
    {
        var file = Path.Join(_tempDir, "Normalized.pptx");

        var created = await CallAsync("presentation", new() { ["action"] = "create", ["file_path"] = file });
        var sessionId = McpToolCallHelper.GetString(created, "sessionId")!;

        await CallAsync("slide", new() { ["action"] = "add_blank", ["sessionId"] = sessionId });

        var added = await CallAsync("shape", new()
        {
            ["action"] = "addTextBox",
            ["sessionId"] = sessionId,
            ["slideIndex"] = 1,
            ["left"] = 50,
            ["top"] = 50,
            ["width"] = 400,
            ["height"] = 60,
            ["text"] = "Normalized",
        });
        var shapeIndex = McpToolCallHelper.GetInt(added, "shapeIndex");
        Assert.NotNull(shapeIndex);

        var text = await CallAsync("textframe", new()
        {
            ["action"] = "get_text",
            ["session_id"] = sessionId,
            ["slide_index"] = 1,
            ["shape_index"] = shapeIndex,
        });
        Assert.Equal("Normalized", McpToolCallHelper.GetString(text, "text"));

        await CallAsync("presentation", new() { ["action"] = "close", ["session_id"] = sessionId, ["save"] = true });
        Assert.True(File.Exists(file));
    }

    private async Task<string> CallAsync(string tool, Dictionary<string, object?> arguments)
    {
        var result = await McpToolCallHelper.CallToolAsync(_client!, tool, arguments, _cts.Token);
        _output.WriteLine($"{tool} {arguments["action"]}: {result}");
        McpToolCallHelper.AssertSuccess(result, $"{tool} {arguments["action"]}");
        return result;
    }
}
