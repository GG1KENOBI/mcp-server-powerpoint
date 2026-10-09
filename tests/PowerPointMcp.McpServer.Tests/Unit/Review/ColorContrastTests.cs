// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Review;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Review;

/// <summary>WCAG contrast computations (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Review")]
public sealed class ColorContrastTests
{
    [Theory]
    [InlineData("#000000", "#FFFFFF", 21.0)]
    [InlineData("#FFFFFF", "#FFFFFF", 1.0)]
    [InlineData("#767676", "#FFFFFF", 4.54)]
    [InlineData("FFFFFF", "1F3864", 11.6)]
    public void Ratio_MatchesWcag(string fore, string back, double expected)
    {
        Assert.Equal(expected, ColorContrast.Ratio(fore, back)!.Value, 1);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("#12345")]
    [InlineData("#GGGGGG")]
    public void Ratio_IsNullForMalformedColors(string? color)
    {
        Assert.Null(ColorContrast.Ratio(color, "#FFFFFF"));
    }
}
