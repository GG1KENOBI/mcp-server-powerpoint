// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Table;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit;

/// <summary>Row parsing for table set-data: plain rows, pasted markdown, escapes, separators.</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Table")]
public sealed class TableDataParserTests
{
    private static string[][] Parse(string separator, params string[] rows) =>
        TableDataParser.Parse(rows, separator).Select(row => row.ToArray()).ToArray();

    [Fact]
    public void PlainRows_SplitAndTrim()
    {
        Assert.Equal(
            [["Region", "Q1", "Q2"], ["North", "12", "15"]],
            Parse("|", "Region|Q1|Q2", " North | 12 |15 "));
    }

    [Fact]
    public void MarkdownRows_DropEdgeSeparatorsAndTheDividerRow()
    {
        Assert.Equal(
            [["Region", "Revenue"], ["APAC", "$2.4M"]],
            Parse("|", "| Region | Revenue |", "|---|:---:|", "| APAC | $2.4M |"));
    }

    [Fact]
    public void EmptyCells_ArePreserved()
    {
        Assert.Equal([["a", "", "c"]], Parse("|", "a||c"));
        Assert.Equal([["", "b"]], Parse("|", "| | b |"));
    }

    [Fact]
    public void EscapedSeparator_IsLiteralText()
    {
        Assert.Equal([["either|or", "x"]], Parse("|", @"either\|or|x"));
        Assert.Equal([["ends with |"]], Parse("|", @"ends with \|"));
    }

    [Fact]
    public void CustomAndMultiCharacterSeparators_Work()
    {
        Assert.Equal([["a|b", "c"]], Parse(";", "a|b;c"));
        Assert.Equal([["a", "b", "c"]], Parse("::", "a::b::c"));
    }

    [Fact]
    public void SingleCellRows_AreKept()
    {
        Assert.Equal([["only"], [""]], Parse("|", "only", ""));
    }

    [Fact]
    public void DashesInsideRealData_AreNotADivider()
    {
        Assert.Equal([["-", "n/a"]], Parse("|", "-|n/a"));
    }

    [Fact]
    public void EmptySeparator_IsRejected()
    {
        Assert.Throws<ArgumentException>(() => TableDataParser.Parse(["a"], ""));
    }
}
