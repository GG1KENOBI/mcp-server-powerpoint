// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using ModelContextProtocol.Client;
using Sbroenne.PowerPointMcp.Tests.Shared;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// End-to-end smoke test of the PUBLISHED standalone exe over stdio: <c>presentation create</c>
/// must reach PowerPoint through built-in COM. The in-process tests in this project run the
/// development build and cannot catch a publish that disables built-in COM.
/// Set POWERPOINTMCP_PUBLISHED_EXE to test an existing exe (e.g. a release download); otherwise
/// the server is published with the release settings first.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresPowerPoint", "true")]
public sealed class PublishedServerSmokeTests(ITestOutputHelper output) : IDisposable
{
    private const string ExeVariable = "POWERPOINTMCP_PUBLISHED_EXE";
    private const string ComDisabledMessage = "Built-in COM has been disabled";

    private readonly string _tempDir = Directory.CreateTempSubdirectory("PowerPointMcp.PublishedSmoke-").FullName;

    [Fact]
    public async Task PublishedExe_CreatePresentation_ReachesPowerPointThroughBuiltInCom()
    {
        var exe = Environment.GetEnvironmentVariable(ExeVariable);
        if (string.IsNullOrWhiteSpace(exe))
        {
            exe = PublishedMcpServer.Publish(PublishedMcpServer.FindRepoRoot(), Path.Combine(_tempDir, "build"));
        }
        output.WriteLine($"Published server: {exe}");

        var stderr = new List<string>();
        using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        var transport = new StdioClientTransport(new StdioClientTransportOptions
        {
            Name = "PublishedServerSmoke",
            Command = exe,
            WorkingDirectory = _tempDir,
            StandardErrorLines = line => { lock (stderr) { stderr.Add(line); } },
        });

        var file = Path.Combine(_tempDir, "PublishedSmoke.pptx");
        await using (var client = await McpClient.CreateAsync(transport, cancellationToken: cts.Token))
        {
            var create = await McpToolCallHelper.CallToolAsync(
                client,
                "presentation",
                new Dictionary<string, object?> { ["action"] = "create", ["filePath"] = file },
                cts.Token);
            output.WriteLine(create);

            Assert.DoesNotContain(ComDisabledMessage, create, StringComparison.Ordinal);
            McpToolCallHelper.AssertSuccess(create, "presentation create");
            var sessionId = McpToolCallHelper.GetString(create, "sessionId");
            Assert.False(string.IsNullOrEmpty(sessionId), $"Expected a sessionId: {create}");
            Assert.True(File.Exists(file), $"Expected file to exist: {file}");

            var close = await McpToolCallHelper.CallToolAsync(
                client,
                "presentation",
                new Dictionary<string, object?> { ["action"] = "close", ["sessionId"] = sessionId },
                cts.Token);
            McpToolCallHelper.AssertSuccess(close, "presentation close");
        }

        lock (stderr)
        {
            Assert.DoesNotContain(stderr, line => line.Contains(ComDisabledMessage, StringComparison.Ordinal));
        }
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch (IOException)
        {
            // Best effort: PowerPoint may still hold the file briefly after close.
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
