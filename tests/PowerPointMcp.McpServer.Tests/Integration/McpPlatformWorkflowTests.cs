// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit.Abstractions;
using static Sbroenne.PowerPointMcp.McpServer.Tests.Integration.McpToolCallHelper;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// End-to-end authoring-platform workflow through the MCP protocol with real PowerPoint:
/// capabilities, template slides, composition, data-bound chart, diagram, deck-wide text
/// replacement, design preview, validation, rendered preview with image content, batch
/// validation, picture placement, and save-as to a new file. One session amortizes PowerPoint
/// start-up (Rule 30: no mocking).
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "AuthoringPlatform")]
[Trait("RequiresPowerPoint", "true")]
public sealed class McpPlatformWorkflowTests : IAsyncLifetime, IAsyncDisposable
{
    private const string OnePixelPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpPlatformWorkflowTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"McpPlatformWorkflow_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public async Task InitializeAsync()
    {
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(_clientToServerPipe, _serverToClientPipe, "PlatformWorkflowTestClient", _cts.Token);
    }

    public async Task DisposeAsync() => await DisposeAsyncCore();

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        await ProgramTransportTestHost.StopAsync(_client, _clientToServerPipe, _serverToClientPipe, _serverTask, _output);
        _cts.Dispose();
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch (IOException)
            {
                // Best-effort cleanup only.
            }
        }
    }

    [Fact]
    public async Task AuthoringPlatformWorkflow_ViaMcpProtocol()
    {
        var deckPath = Path.Join(_tempDir, "Platform.pptx");
        var copyPath = Path.Join(_tempDir, "Platform-final.pptx");
        var csvPath = Path.Join(_tempDir, "выручка.csv");
        var imagePath = Path.Join(_tempDir, "logo.png");
        await File.WriteAllTextAsync(csvPath, "Квартал;Выручка\nQ1;10,5\nQ2;12,0\nQ3;14,25\n", new System.Text.UTF8Encoding(true));
        await File.WriteAllBytesAsync(imagePath, Convert.FromBase64String(OnePixelPngBase64));

        // 1. Discovery without a session.
        var capabilities = await Call("capabilities", new() { ["action"] = "workflows" });
        AssertSuccess(capabilities, "capabilities");
        Assert.Contains("build-slides", capabilities, StringComparison.Ordinal);

        var created = await Call("presentation", new() { ["action"] = "create", ["filePath"] = deckPath });
        AssertSuccess(created, "presentation.create");
        var session = GetString(created, "presentation_session_id")!;

        // 2. A slide from the deck's own layout, filled by role.
        var layouts = await Call("template_read", new() { ["action"] = "list-layouts", ["presentation_session_id"] = session });
        AssertSuccess(layouts, "template.list-layouts");
        var layoutName = FirstLayoutWith(layouts, "title", "body");
        var added = await Call("template", new()
        {
            ["action"] = "add-slide",
            ["presentation_session_id"] = session,
            ["layout"] = layoutName,
            ["content"] = """{"title":"Итоги квартала","body":["Выручка выросла","Маржа стабильна"],"notes":"Начать с выручки"}""",
        });
        AssertSuccess(added, "template.add-slide");

        // 3. A composed slide, planned first.
        const string spec = """{"kind":"cards","id":"prio","title":"Priorities","cards":[{"heading":"Growth","body":"Two new markets"},{"heading":"People","body":"Hire 40 engineers"},{"heading":"Cost","body":"Save 5% in operations"}]}""";
        AssertSuccess(await Call("compose_read", new() { ["action"] = "plan", ["presentation_session_id"] = session, ["spec"] = spec }), "compose.plan");
        AssertSuccess(await Call("compose", new() { ["action"] = "create", ["presentation_session_id"] = session, ["spec"] = spec }), "compose.create");

        // 4. Data-bound chart from a Russian CSV with decimal commas.
        var preview = await Call("data_read", new() { ["action"] = "preview", ["presentation_session_id"] = session, ["source_path"] = csvPath, ["culture"] = "ru-RU" });
        AssertSuccess(preview, "data.preview");
        AssertSuccess(await Call("data", new() { ["action"] = "create-chart", ["presentation_session_id"] = session, ["source_path"] = csvPath, ["culture"] = "ru-RU", ["title"] = "Выручка по кварталам" }), "data.create-chart");

        // 5. Diagram.
        AssertSuccess(await Call("diagram", new()
        {
            ["action"] = "create",
            ["presentation_session_id"] = session,
            ["spec"] = """{"type":"flowchart","id":"flow","title":"Approval","nodes":[{"id":"a","label":"Request"},{"id":"b","label":"Review"},{"id":"c","label":"Approve"}],"edges":[{"from":"a","to":"b"},{"from":"b","to":"c"}]}""",
        }), "diagram.create");

        // 6. Deck-wide replace: preview, then apply with a count guard.
        var dry = await Call("deck", new() { ["action"] = "replace-text", ["presentation_session_id"] = session, ["find_what"] = "Выручка", ["replace_what"] = "Доход" });
        AssertSuccess(dry, "deck.replace-text (dry run)");
        var count = GetInt(dry, "totalCount")!.Value;
        Assert.True(count >= 2, dry);
        var applied = await Call("deck", new() { ["action"] = "replace-text", ["presentation_session_id"] = session, ["find_what"] = "Выручка", ["replace_what"] = "Доход", ["dry_run"] = false, ["expected_count"] = count });
        AssertSuccess(applied, "deck.replace-text");
        Assert.Equal(count, GetInt(applied, "replacementCount"));

        // 7. Design preview, validation, rendered preview with image content.
        var design = await Call("design", new() { ["action"] = "apply-profile", ["presentation_session_id"] = session, ["profile"] = "corporate-blue" });
        AssertSuccess(design, "design.apply-profile (dry run)");
        Assert.False(GetBool(design, "applied"));
        AssertSuccess(await Call("review_read", new() { ["action"] = "validate", ["presentation_session_id"] = session }), "review.validate");
        var snapshot = await _client!.CallToolAsync("preview_read", new Dictionary<string, object?> { ["action"] = "snapshot", ["presentation_session_id"] = session, ["slide_index"] = 2 }, cancellationToken: _cts.Token);
        Assert.NotEqual(true, snapshot.IsError);
        Assert.Contains(snapshot.Content, block => block is ImageContentBlock { MimeType: "image/png" });
        Assert.Contains(snapshot.Content, block => block is TextContentBlock);

        // 8. Batch validation and picture placement.
        AssertSuccess(await Call("batch_read", new()
        {
            ["action"] = "validate",
            ["presentation_session_id"] = session,
            ["operations"] = """[{"id":"s","tool":"deck","action":"summary"},{"id":"f","tool":"deck","action":"find","args":{"selector":"kind:chart"}}]""",
        }), "batch.validate");
        AssertSuccess(await Call("asset", new() { ["action"] = "place", ["presentation_session_id"] = session, ["slide_index"] = 1, ["path"] = imagePath, ["decorative"] = true }), "asset.place");
        AssertSuccess(await Call("asset_read", new() { ["action"] = "inspect", ["presentation_session_id"] = session }), "asset.inspect");

        // 9. Keep the original: save a copy, then close without saving the working file.
        AssertSuccess(await Call("presentation", new() { ["action"] = "save-copy-as", ["presentation_session_id"] = session, ["targetPath"] = copyPath }), "presentation.save-copy-as");
        Assert.True(File.Exists(copyPath));
        AssertSuccess(await Call("presentation", new() { ["action"] = "close", ["presentation_session_id"] = session, ["save"] = false }), "presentation.close");
    }

    private static string FirstLayoutWith(string layoutsJson, params string[] roles)
    {
        using var document = JsonDocument.Parse(layoutsJson);
        foreach (var layout in document.RootElement.GetProperty("layouts").EnumerateArray())
        {
            var present = layout.GetProperty("placeholders").EnumerateArray().Select(placeholder => placeholder.GetProperty("role").GetString()).ToList();
            if (roles.All(role => present.Count(item => item == role) == 1))
                return layout.GetProperty("name").GetString()!;
        }
        throw new InvalidOperationException("No layout with a title and one body placeholder.");
    }

    private Task<string> Call(string toolName, Dictionary<string, object?> arguments) => CallToolAsync(_client!, toolName, arguments, _cts.Token);
}
