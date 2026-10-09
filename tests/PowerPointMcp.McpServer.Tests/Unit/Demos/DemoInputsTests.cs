// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.RegularExpressions;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Data;
using Sbroenne.PowerPointMcp.Core.Design;
using Sbroenne.PowerPointMcp.Core.Diagram;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Demos;

/// <summary>Every input the demonstration scenarios use parses with the real validators (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Demos")]
public sealed partial class DemoInputsTests
{
    private static readonly string DemosDir = Path.Combine(FindRepoRoot(), "demos");

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.PowerPointMcp.slnx")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new InvalidOperationException("Repository root not found.");
    }

    public static TheoryData<string> CompositionFiles() => Names("*.json", diagrams: false);

    public static TheoryData<string> DiagramFiles() => Names("diagram-*.json", diagrams: true);

    private static TheoryData<string> Names(string pattern, bool diagrams)
    {
        var data = new TheoryData<string>();
        foreach (var path in Directory.GetFiles(Path.Combine(DemosDir, "specs"), pattern))
        {
            var name = Path.GetFileName(path);
            if (name.StartsWith("diagram-", StringComparison.Ordinal) == diagrams)
                data.Add(name);
        }
        return data;
    }

    [Theory]
    [MemberData(nameof(CompositionFiles))]
    public void CompositionSpecs_AreValid(string file)
    {
        var result = CompositionParser.Parse(File.ReadAllText(Path.Combine(DemosDir, "specs", file)));
        Assert.True(result.Errors.Count == 0, string.Join("; ", result.Errors));
    }

    [Theory]
    [MemberData(nameof(DiagramFiles))]
    public void DiagramSpecs_AreValid(string file)
    {
        var (spec, errors, _) = DiagramParser.Parse(File.ReadAllText(Path.Combine(DemosDir, "specs", file)));
        Assert.True(errors.Count == 0 && spec is not null, string.Join("; ", errors));
    }

    [Fact]
    public void InlineCompositionSpecsInScripts_AreValid()
    {
        var specs = Directory.GetFiles(DemosDir, "*.ps1")
            .SelectMany(path => InlineSpec().Matches(File.ReadAllText(path)).Select(match => (Path.GetFileName(path), match.Groups["json"].Value)))
            .ToList();

        Assert.True(specs.Count >= 12, $"Expected the scenarios' inline specs, found {specs.Count}.");
        foreach (var (file, json) in specs)
        {
            var result = CompositionParser.Parse(json);
            Assert.True(result.Errors.Count == 0, $"{file}: {string.Join("; ", result.Errors)} in {json}");
        }
    }

    [Fact]
    public void DemoProfile_IsValid()
    {
        var profile = DesignProfiles.Parse(File.ReadAllText(Path.Combine(DemosDir, "profiles", "acme.json")));
        var resolved = DesignProfiles.Resolve(profile);
        Assert.True(resolved.IsValid, string.Join("; ", resolved.Errors));
        Assert.Equal("2", resolved.Resolved!.Components["card"].Version);
    }

    [Fact]
    public void RussianCsv_ParsesWithDecimalCommas()
    {
        var binding = new DataBinding { Source = Path.Combine(DemosDir, "data", "quarterly-ru.csv"), Kind = "table", Culture = "ru-RU" };
        var (table, _) = DataLoader.Load(binding);
        Assert.Equal(4, table.Columns.Count);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal("number", table.Columns[1].Type);
    }

    [GeneratedRegex("""'(?<json>\{"kind":.*?\})'""", RegexOptions.CultureInvariant)]
    private static partial Regex InlineSpec();
}
