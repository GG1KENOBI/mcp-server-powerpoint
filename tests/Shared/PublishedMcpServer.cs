// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Buffers.Binary;
using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace Sbroenne.PowerPointMcp.Tests.Shared;

/// <summary>
/// Publishes the MCP server exactly like the standalone-exe release ("Publish MCP Server" in
/// .github/workflows/release.yml) and reads the runtimeconfig.json embedded in the resulting
/// single-file bundle — the configuration the shipped exe actually runs with, which a normal
/// development build never exercises.
/// </summary>
internal static class PublishedMcpServer
{
    internal const string ExeName = "Sbroenne.PowerPointMcp.McpServer.exe";
    internal const string BuiltInComSwitch = "System.Runtime.InteropServices.BuiltInComInterop.IsSupported";

    /// <summary>Release publish arguments (minus version stamping and the output folder).</summary>
    internal static readonly string[] ReleasePublishArguments =
    [
        "--configuration", "Release",
        "--runtime", "win-x64",
        "--self-contained", "true",
        "-p:PublishSingleFile=true",
        "-p:IncludeNativeLibrariesForSelfExtract=true",
        "-p:PublishReadyToRun=false",
        "-p:NuGetAudit=false",
    ];

    // SHA-256 of ".net core bundle": the marker the single-file host uses to find the bundle header.
    private static readonly byte[] BundleSignature =
    [
        0x8b, 0x12, 0x02, 0xb9, 0x6a, 0x61, 0x20, 0x38, 0x72, 0x7b, 0x93, 0x02, 0x14, 0xd7, 0xa0, 0x32,
        0x13, 0xf5, 0xb9, 0xe6, 0xef, 0xae, 0x33, 0x18, 0xee, 0x3b, 0x2d, 0xce, 0x24, 0xb3, 0x6a, 0xae,
    ];

    internal static string ProjectPath(string repoRoot) => Path.Combine(
        repoRoot, "src", "PowerPointMcp.McpServer", "PowerPointMcp.McpServer.csproj");

    /// <summary>
    /// Publishes into <paramref name="workDirectory"/> (with its own artifacts path so the
    /// repository's bin/obj folders are not touched) and returns the published exe path.
    /// </summary>
    internal static string Publish(string repoRoot, string workDirectory)
    {
        var output = Path.Combine(workDirectory, "publish");
        var result = RunDotnet(
            repoRoot,
            TimeSpan.FromMinutes(15),
            ["publish", ProjectPath(repoRoot), .. ReleasePublishArguments,
             "--artifacts-path", Path.Combine(workDirectory, "artifacts"),
             "--output", output,
             "-p:PowerPointMcpSkipCleanup=true"]);
        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException($"dotnet publish failed ({result.ExitCode}).{Environment.NewLine}{result.Output}");
        }

        var exe = Path.Combine(output, ExeName);
        if (!File.Exists(exe))
        {
            throw new FileNotFoundException($"Published MCP server not found.{Environment.NewLine}{result.Output}", exe);
        }

        return exe;
    }

    /// <summary>Reads runtimeconfig.json from a single-file bundle's header.</summary>
    internal static JsonDocument ReadBundledRuntimeConfig(string exePath)
    {
        var bytes = File.ReadAllBytes(exePath);
        var signature = bytes.AsSpan().IndexOf(BundleSignature);
        if (signature < sizeof(long))
        {
            throw new InvalidDataException($"Not a single-file bundle: {exePath}");
        }

        var header = checked((int)BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(signature - sizeof(long))));
        // Header: uint32 major, uint32 minor, int32 file count, 7-bit-length-prefixed bundle id,
        // then (v2+) int64 deps.json offset/size and int64 runtimeconfig.json offset/size.
        var major = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(header));
        if (major < 2)
        {
            throw new InvalidDataException($"Unsupported bundle version {major}: {exePath}");
        }

        var position = header + 12;
        int idLength = 0, shift = 0;
        byte next;
        do
        {
            next = bytes[position++];
            idLength |= (next & 0x7f) << shift;
            shift += 7;
        }
        while ((next & 0x80) != 0);
        position += idLength + (2 * sizeof(long));

        var offset = checked((int)BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(position)));
        var size = checked((int)BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(position + sizeof(long))));
        if (size <= 0)
        {
            throw new InvalidDataException($"Bundle has no runtimeconfig.json: {exePath}");
        }

        return JsonDocument.Parse(Encoding.UTF8.GetString(bytes, offset, size));
    }

    internal static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.PowerPointMcp.slnx")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName
            ?? throw new InvalidOperationException("Could not locate the repository root.");
    }

    internal static (int ExitCode, string Output) RunDotnet(
        string workingDirectory,
        TimeSpan timeout,
        params string[] arguments)
    {
        var startInfo = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        // Keep MSBuild worker nodes, the MSBuild server and the compiler server from outliving the
        // command: they inherit the redirected pipes, and reading to EOF would then never finish.
        startInfo.Environment["MSBUILDDISABLENODEREUSE"] = "1";
        startInfo.Environment["DOTNET_CLI_USE_MSBUILD_SERVER"] = "0";
        startInfo.Environment["UseSharedCompilation"] = "false";
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start dotnet.");
        var standardOutput = process.StandardOutput.ReadToEndAsync();
        var standardError = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(timeout))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"dotnet {string.Join(' ', arguments)} timed out after {timeout}.");
        }

        if (!Task.WaitAll([standardOutput, standardError], TimeSpan.FromMinutes(1)))
        {
            throw new TimeoutException($"dotnet {string.Join(' ', arguments)} exited but a child process kept its output open.");
        }

        return (
            process.ExitCode,
            $"{standardOutput.GetAwaiter().GetResult()}{Environment.NewLine}{standardError.GetAwaiter().GetResult()}");
    }
}
