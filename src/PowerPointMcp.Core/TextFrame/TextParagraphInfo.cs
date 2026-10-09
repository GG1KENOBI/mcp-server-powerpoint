namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>A run of characters with uniform formatting.</summary>
public sealed class TextRunInfo
{
    /// <summary>1-based PowerPoint character position of the run in the frame.</summary>
    public required int Start { get; init; }

    /// <summary>Run length in characters.</summary>
    public required int Length { get; init; }

    /// <summary>Run text.</summary>
    public required string Text { get; init; }

    /// <summary>Typeface.</summary>
    public string? FontName { get; init; }

    /// <summary>Size in points.</summary>
    public float? FontSize { get; init; }

    /// <summary>Bold.</summary>
    public bool? Bold { get; init; }

    /// <summary>Italic.</summary>
    public bool? Italic { get; init; }

    /// <summary>Underlined.</summary>
    public bool? Underline { get; init; }

    /// <summary>Text color as #RRGGBB.</summary>
    public string? Color { get; init; }

    /// <summary>Click hyperlink address (or "slide:N" for a link to another slide).</summary>
    public string? Hyperlink { get; init; }

    /// <summary>Character spacing in points (0 is normal); omitted when 0.</summary>
    public float? CharacterSpacing { get; init; }

    /// <summary>Baseline offset (positive superscript, negative subscript); omitted when 0.</summary>
    public float? BaselineOffset { get; init; }
}

/// <summary>One paragraph with its paragraph format and runs.</summary>
public sealed class TextParagraphInfo
{
    /// <summary>1-based paragraph number.</summary>
    public required int Index { get; init; }

    /// <summary>1-based character position of the paragraph.</summary>
    public required int Start { get; init; }

    /// <summary>Length without the paragraph separator.</summary>
    public required int Length { get; init; }

    /// <summary>Paragraph text (line breaks inside it are shown as \n).</summary>
    public required string Text { get; init; }

    /// <summary>left, center, right, justify, distribute, or mixed.</summary>
    public string? Alignment { get; init; }

    /// <summary>Outline level 1-9.</summary>
    public int? IndentLevel { get; init; }

    /// <summary>Space before in points.</summary>
    public float? SpaceBefore { get; init; }

    /// <summary>Space after in points.</summary>
    public float? SpaceAfter { get; init; }

    /// <summary>Line spacing as a multiple of single spacing (when set in lines).</summary>
    public float? LineSpacing { get; init; }

    /// <summary>Exact line spacing in points (when set in points).</summary>
    public float? LineSpacingPoints { get; init; }

    /// <summary>Left indent in points.</summary>
    public float? LeftIndent { get; init; }

    /// <summary>First-line indent in points (negative for a hanging indent).</summary>
    public float? FirstLineIndent { get; init; }

    /// <summary>none, bullet, numbered, or picture.</summary>
    public string? Bullet { get; init; }

    /// <summary>Bullet glyph, for bullet.</summary>
    public string? BulletCharacter { get; init; }

    /// <summary>Numbering style, for numbered (arabic-period, alpha-lower-paren, roman-upper-period, ...).</summary>
    public string? NumberStyle { get; init; }

    /// <summary>First number, for numbered.</summary>
    public int? StartAt { get; init; }

    /// <summary>Formatting runs.</summary>
    public IReadOnlyList<TextRunInfo>? Runs { get; init; }
}

/// <summary>Text frame layout: margins, anchoring, wrapping, autofit, orientation, columns, and the shape's rotation.</summary>
public sealed class TextFrameLayoutInfo
{
    /// <summary>Left, top, right, bottom margins in points.</summary>
    public required IReadOnlyList<float> Margins { get; init; }

    /// <summary>top, middle, bottom, top-baseline, or bottom-baseline.</summary>
    public required string VerticalAnchor { get; init; }

    /// <summary>Whether text wraps inside the shape.</summary>
    public required bool WordWrap { get; init; }

    /// <summary>none, shape-to-fit-text, shrink-on-overflow, or mixed.</summary>
    public required string AutoSize { get; init; }

    /// <summary>horizontal, up, down, stacked, vertical-east-asian, or horizontal-rotated-east-asian.</summary>
    public required string Orientation { get; init; }

    /// <summary>Number of text columns.</summary>
    public required int Columns { get; init; }

    /// <summary>Space between columns in points.</summary>
    public required float ColumnSpacing { get; init; }

    /// <summary>Shape rotation in degrees.</summary>
    public required float Rotation { get; init; }
}
