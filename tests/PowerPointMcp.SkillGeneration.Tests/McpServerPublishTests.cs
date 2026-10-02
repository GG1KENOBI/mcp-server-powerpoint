using System.Text.Json;
using Sbroenne.PowerPointMcp.Tests.Shared;

namespace Sbroenne.PowerPointMcp.SkillGeneration.Tests;

/// <summary>
/// Guards the standalone MCP server exe against shipping with built-in COM disabled. Trimming
/// (or Native AOT) turns off System.Runtime.InteropServices.BuiltInComInterop.IsSupported, after
/// which every PowerPoint call fails with "Built-in COM has been disabled via a feature switch"
/// — invisible to development builds and to the in-process MCP tests.
/// </summary>
[Trait("RequiresPowerPoint", "false")]
public sealed class McpServerPublishTests
{
    private static readonly string RepoRoot = PublishedMcpServer.FindRepoRoot();

    [Fact]
    public void ReleasePublishProperties_KeepBuiltInComEnabled()
    {
        var properties = EvaluateReleaseProperties();

        Assert.Equal("false", properties["PublishTrimmed"]);
        Assert.Equal("false", properties["PublishAot"]);
        Assert.Equal("true", properties["BuiltInComInteropSupport"]);
        Assert.Equal("true", properties["SelfContained"]);
        Assert.Equal("win-x64", properties["RuntimeIdentifier"]);
    }

    [Theory]
    [InlineData("-p:PublishTrimmed=true")]
    [InlineData("-p:PublishAot=true")]
    [InlineData("-p:BuiltInComInteropSupport=false")]
    public void PublishGuard_RejectsComIncompatibleOverrides(string overrideArgument)
    {
        var result = PublishedMcpServer.RunDotnet(
            RepoRoot,
            TimeSpan.FromMinutes(5),
            ["msbuild", PublishedMcpServer.ProjectPath(RepoRoot), "-t:_EnsureBuiltInComInteropForPublish",
             "-p:Configuration=Release", "-p:RuntimeIdentifier=win-x64", "-p:SelfContained=true",
             "-p:PowerPointMcpSkipCleanup=true", overrideArgument]);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("requires built-in COM interop", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public void PublishedReleaseExe_IsSelfContainedUntrimmedWithBuiltInComEnabled()
    {
        var work = Path.Combine(Path.GetTempPath(), $"PowerPointMcp.PublishTest-{Guid.NewGuid():N}");
        try
        {
            var exe = PublishedMcpServer.Publish(RepoRoot, work);
            using var config = PublishedMcpServer.ReadBundledRuntimeConfig(exe);
            var options = config.RootElement.GetProperty("runtimeOptions");
            var switches = options.GetProperty("configProperties");

            Assert.True(
                switches.TryGetProperty(PublishedMcpServer.BuiltInComSwitch, out var builtInCom)
                    && builtInCom.ValueKind == JsonValueKind.True,
                $"{PublishedMcpServer.BuiltInComSwitch} must be true in the published exe:{Environment.NewLine}{options}");
            // Only emitted by the trimming targets.
            Assert.False(
                switches.TryGetProperty("Microsoft.Extensions.DependencyInjection.VerifyOpenGenericServiceTrimmability", out _),
                $"The published exe was trimmed:{Environment.NewLine}{options}");
            // Self-contained bundles list includedFrameworks; framework-dependent ones list frameworks.
            Assert.True(options.TryGetProperty("includedFrameworks", out _), $"Not self-contained:{Environment.NewLine}{options}");
        }
        finally
        {
            if (Directory.Exists(work))
            {
                Directory.Delete(work, recursive: true);
            }
        }
    }

    private static Dictionary<string, string> EvaluateReleaseProperties()
    {
        string[] names = ["PublishTrimmed", "PublishAot", "BuiltInComInteropSupport", "SelfContained", "RuntimeIdentifier"];
        var result = PublishedMcpServer.RunDotnet(
            RepoRoot,
            TimeSpan.FromMinutes(5),
            ["msbuild", PublishedMcpServer.ProjectPath(RepoRoot), "-t:_EnsureBuiltInComInteropForPublish",
             "-p:Configuration=Release", "-p:RuntimeIdentifier=win-x64", "-p:SelfContained=true",
             "-p:PublishSingleFile=true", "-p:PowerPointMcpSkipCleanup=true",
             .. names.Select(name => $"-getProperty:{name}")]);
        Assert.True(result.ExitCode == 0, result.Output);

        using var json = JsonDocument.Parse(result.Output[result.Output.IndexOf('{', StringComparison.Ordinal)..]);
        var properties = json.RootElement.GetProperty("Properties");
        return names.ToDictionary(name => name, name => properties.GetProperty(name).GetString() ?? string.Empty);
    }
}
