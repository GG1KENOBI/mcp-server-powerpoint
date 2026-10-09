extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Deck;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

public sealed partial class TextFrameCommands
{
    private const int MaxParagraphsListed = 200;
    private const int MaxRunsListed = 400;

    private static readonly Dictionary<string, PowerPoint.PpParagraphAlignment> ShortAlignments = new(StringComparer.OrdinalIgnoreCase)
    {
        ["left"] = PowerPoint.PpParagraphAlignment.ppAlignLeft,
        ["center"] = PowerPoint.PpParagraphAlignment.ppAlignCenter,
        ["centre"] = PowerPoint.PpParagraphAlignment.ppAlignCenter,
        ["right"] = PowerPoint.PpParagraphAlignment.ppAlignRight,
        ["justify"] = PowerPoint.PpParagraphAlignment.ppAlignJustify,
        ["distribute"] = PowerPoint.PpParagraphAlignment.ppAlignDistribute,
    };

    private static readonly Dictionary<string, PowerPoint.PpAutoSize> ShortAutoSizes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["none"] = PowerPoint.PpAutoSize.ppAutoSizeNone,
        ["shape-to-fit-text"] = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText,
        ["shrink-on-overflow"] = (PowerPoint.PpAutoSize)2,
        ["shrink-text-on-overflow"] = (PowerPoint.PpAutoSize)2,
    };

    private static readonly Dictionary<string, PowerPoint.PpNumberedBulletStyle> NumberingStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["arabic-period"] = PowerPoint.PpNumberedBulletStyle.ppBulletArabicPeriod,
        ["arabic-paren"] = PowerPoint.PpNumberedBulletStyle.ppBulletArabicParenRight,
        ["alpha-lower-period"] = PowerPoint.PpNumberedBulletStyle.ppBulletAlphaLCPeriod,
        ["alpha-upper-period"] = PowerPoint.PpNumberedBulletStyle.ppBulletAlphaUCPeriod,
        ["alpha-lower-paren"] = PowerPoint.PpNumberedBulletStyle.ppBulletAlphaLCParenRight,
        ["roman-lower-period"] = PowerPoint.PpNumberedBulletStyle.ppBulletRomanLCPeriod,
        ["roman-upper-period"] = PowerPoint.PpNumberedBulletStyle.ppBulletRomanUCPeriod,
    };

    private static readonly Dictionary<string, Office.MsoTextOrientation> Orientations = new(StringComparer.OrdinalIgnoreCase)
    {
        ["horizontal"] = Office.MsoTextOrientation.msoTextOrientationHorizontal,
        ["up"] = Office.MsoTextOrientation.msoTextOrientationUpward,
        ["down"] = Office.MsoTextOrientation.msoTextOrientationDownward,
        ["stacked"] = Office.MsoTextOrientation.msoTextOrientationVertical,
        ["vertical-east-asian"] = Office.MsoTextOrientation.msoTextOrientationVerticalFarEast,
        ["horizontal-rotated-east-asian"] = Office.MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast,
    };

    private static readonly Dictionary<string, Office.MsoVerticalAnchor> Anchors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["top"] = Office.MsoVerticalAnchor.msoAnchorTop,
        ["middle"] = Office.MsoVerticalAnchor.msoAnchorMiddle,
        ["bottom"] = Office.MsoVerticalAnchor.msoAnchorBottom,
        ["top-baseline"] = Office.MsoVerticalAnchor.msoAnchorTopBaseline,
        ["bottom-baseline"] = Office.MsoVerticalAnchor.msoAnchorBottomBaseLine,
    };

    /// <summary>Parses an alignment given as left/center/... or a PpParagraphAlignment name.</summary>
    internal static bool TryParseAlignment(string value, out PowerPoint.PpParagraphAlignment alignment) =>
        ParagraphAlignments.TryGetValue(value, out alignment) || ShortAlignments.TryGetValue(value, out alignment);

    /// <summary>Parses an autofit mode given as none/shape-to-fit-text/shrink-on-overflow or a PpAutoSize name.</summary>
    internal static bool TryParseAutoSize(string value, out PowerPoint.PpAutoSize autoSize) =>
        AutoSizeModes.TryGetValue(value, out autoSize) || ShortAutoSizes.TryGetValue(value, out autoSize);

    /// <inheritdoc/>
    public TextFrameOperationResult GetParagraphs(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return WithTextShape(batch, slideIndex, shapeIndex, (presentation, shape, ct) =>
        {
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            PowerPoint.TextFrame2? frame2 = null;
            Office.TextRange2? range2 = null;
            try
            {
                frame = shape.TextFrame;
                range = frame.TextRange;
                frame2 = shape.TextFrame2;
                range2 = frame2.TextRange;
                string text = range.Text;
                var bounds = TextRanges.Paragraphs(text);
                var warnings = new List<string>();
                var paragraphs = new List<TextParagraphInfo>();
                int runsListed = 0;
                for (int index = 1; index <= Math.Min(bounds.Count, MaxParagraphsListed); index++)
                {
                    ct.ThrowIfCancellationRequested();
                    var (start, length) = bounds[index - 1];
                    var runs = new List<TextRunInfo>();
                    if (length > 0 && runsListed < MaxRunsListed)
                        runsListed += ReadRuns(presentation, range, range2, start, length, runs, MaxRunsListed - runsListed);
                    paragraphs.Add(ReadParagraph(range, range2, index, start, length, text.Substring(start - 1, length), runs));
                }
                if (bounds.Count > MaxParagraphsListed)
                    warnings.Add($"Only the first {MaxParagraphsListed} of {bounds.Count} paragraphs are listed.");
                if (runsListed >= MaxRunsListed)
                    warnings.Add($"Formatting runs are listed for the first {MaxRunsListed} runs only.");

                return new TextFrameOperationResult
                {
                    Success = true,
                    Text = text,
                    Paragraphs = paragraphs,
                    Frame = ReadLayout(shape, frame, frame2),
                    Warnings = warnings.Count > 0 ? warnings : null,
                };
            }
            finally
            {
                if (range2 is not null) ComUtilities.Release(ref range2);
                if (frame2 is not null) ComUtilities.Release(ref frame2);
                if (range is not null) ComUtilities.Release(ref range);
                if (frame is not null) ComUtilities.Release(ref frame);
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult ReplaceRange(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        string text,
        int? start = null,
        int? length = null,
        string? match = null,
        int occurrence = 1,
        bool matchCase = false,
        int? paragraph = null,
        int? paragraphCount = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (text is null)
            return new TextFrameOperationResult { ErrorMessage = "text is required (use an empty string to delete)." };
        var newText = TextRanges.ToPowerPointParagraphs(text);

        return WithTextShape(batch, slideIndex, shapeIndex, (presentation, shape, ct) =>
        {
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            PowerPoint.TextRange? target = null;
            try
            {
                frame = shape.TextFrame;
                range = frame.TextRange;
                string current = range.Text;
                int first, count;
                try
                {
                    (first, count) = TextRanges.Resolve(current, start, length, match, occurrence, matchCase, paragraph, paragraphCount);
                }
                catch (ArgumentException ex)
                {
                    return new TextFrameOperationResult { ErrorMessage = ex.Message };
                }

                if (count > 0)
                {
                    target = range.Characters(first, count);
                    target.Text = newText;
                }
                else if (newText.Length > 0)
                {
                    // Insertion: borrow the formatting of the character the text goes before (or the last one).
                    if (first <= current.Length)
                    {
                        target = range.Characters(first, 1);
                        target.InsertBefore(newText);
                    }
                    else
                    {
                        range.InsertAfter(newText);
                    }
                }

                ComUtilities.Release(ref range!);
                range = frame.TextRange;
                return new TextFrameOperationResult { Success = true, Text = range.Text, Start = first, Length = newText.Length };
            }
            finally
            {
                if (target is not null) ComUtilities.Release(ref target);
                if (range is not null) ComUtilities.Release(ref range);
                if (frame is not null) ComUtilities.Release(ref frame);
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult FormatRange(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int? start = null,
        int? length = null,
        string? match = null,
        int occurrence = 1,
        bool matchCase = false,
        int? paragraph = null,
        int? paragraphCount = null,
        string? fontName = null,
        float? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? color = null,
        string? hyperlink = null,
        float? characterSpacing = null,
        float? baselineOffset = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (fontName is null && fontSize is null && bold is null && italic is null && underline is null && color is null
            && hyperlink is null && characterSpacing is null && baselineOffset is null)
            return new TextFrameOperationResult { ErrorMessage = "Pass at least one format: font_name, font_size, bold, italic, underline, color, hyperlink, character_spacing, or baseline_offset." };
        if (fontSize is <= 0 or > 4000)
            return new TextFrameOperationResult { ErrorMessage = "font_size must be between 1 and 4000 points." };
        if (baselineOffset is < -1 or > 1)
            return new TextFrameOperationResult { ErrorMessage = "baseline_offset must be between -1 and 1 (0.3 superscript, -0.25 subscript)." };
        if (fontName is { Length: 0 })
            return new TextFrameOperationResult { ErrorMessage = "font_name must not be empty." };
        int? bgr = null;
        if (color is not null)
        {
            if (!TryParseHexColor(color, out var parsed))
                return new TextFrameOperationResult { ErrorMessage = $"color '{color}' is not #RRGGBB." };
            bgr = parsed;
        }

        return WithTextShape(batch, slideIndex, shapeIndex, (presentation, shape, ct) =>
        {
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            PowerPoint.TextRange? target = null;
            PowerPoint.Font? font = null;
            PowerPoint.ColorFormat? colorFormat = null;
            PowerPoint.TextFrame2? frame2 = null;
            Office.TextRange2? range2 = null;
            Office.TextRange2? target2 = null;
            Office.Font2? font2 = null;
            try
            {
                frame = shape.TextFrame;
                range = frame.TextRange;
                int first, count;
                try
                {
                    (first, count) = TextRanges.Resolve(range.Text, start, length, match, occurrence, matchCase, paragraph, paragraphCount);
                }
                catch (ArgumentException ex)
                {
                    return new TextFrameOperationResult { ErrorMessage = ex.Message };
                }
                if (count == 0)
                    return new TextFrameOperationResult { ErrorMessage = "The selected range is empty; pass a length of at least 1." };

                target = range.Characters(first, count);
                font = target.Font;
                if (fontName is not null) font.Name = fontName;
                if (fontSize is { } size) font.Size = size;
                if (bold is { } isBold) font.Bold = isBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (italic is { } isItalic) font.Italic = isItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (underline is { } isUnderline) font.Underline = isUnderline ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (bgr is { } value)
                {
                    colorFormat = font.Color;
                    colorFormat.RGB = value;
                }
                if (characterSpacing is not null || baselineOffset is not null)
                {
                    frame2 = shape.TextFrame2;
                    range2 = frame2.TextRange;
                    target2 = range2.Characters[first, count];
                    font2 = target2.Font;
                    if (characterSpacing is { } spacing) font2.Spacing = spacing;
                    if (baselineOffset is { } offset) font2.BaselineOffset = offset;
                }
                if (hyperlink is not null)
                {
                    var error = WriteHyperlink(presentation, target, hyperlink);
                    if (error is not null)
                        return new TextFrameOperationResult { ErrorMessage = error };
                }

                return new TextFrameOperationResult { Success = true, Start = first, Length = count, Text = target.Text };
            }
            finally
            {
                if (font2 is not null) ComUtilities.Release(ref font2);
                if (target2 is not null) ComUtilities.Release(ref target2);
                if (range2 is not null) ComUtilities.Release(ref range2);
                if (frame2 is not null) ComUtilities.Release(ref frame2);
                if (colorFormat is not null) ComUtilities.Release(ref colorFormat);
                if (font is not null) ComUtilities.Release(ref font);
                if (target is not null) ComUtilities.Release(ref target);
                if (range is not null) ComUtilities.Release(ref range);
                if (frame is not null) ComUtilities.Release(ref frame);
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetParagraphFormat(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int? paragraph = null,
        int? paragraphCount = null,
        string? alignment = null,
        float? spaceBefore = null,
        float? spaceAfter = null,
        float? lineSpacing = null,
        float? lineSpacingPoints = null,
        int? indentLevel = null,
        float? leftIndent = null,
        float? firstLineIndent = null,
        string? bulletStyle = null,
        string? character = null,
        string? numberStyle = null,
        int? startAt = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        PowerPoint.PpParagraphAlignment parsedAlignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
        if (alignment is not null && !TryParseAlignment(alignment, out parsedAlignment))
            return new TextFrameOperationResult { ErrorMessage = $"alignment '{alignment}' is not left, center, right, justify, or distribute." };
        if (lineSpacing is not null && lineSpacingPoints is not null)
            return new TextFrameOperationResult { ErrorMessage = "Pass line_spacing (in lines) or line_spacing_points, not both." };
        if (lineSpacing is <= 0 or > 10)
            return new TextFrameOperationResult { ErrorMessage = "line_spacing must be between 0.1 and 10 lines (1 is single spacing)." };
        if (lineSpacingPoints is <= 0 or > 1584)
            return new TextFrameOperationResult { ErrorMessage = "line_spacing_points must be between 0.1 and 1584 points." };
        if (spaceBefore is < 0 || spaceAfter is < 0)
            return new TextFrameOperationResult { ErrorMessage = "space_before and space_after must be 0 or greater." };
        if (indentLevel is < 1 or > 9)
            return new TextFrameOperationResult { ErrorMessage = "indent_level must be between 1 and 9." };
        if (bulletStyle is not null && bulletStyle is not ("none" or "bullet" or "numbered"))
            return new TextFrameOperationResult { ErrorMessage = $"bullet_style '{bulletStyle}' is not none, bullet, or numbered." };
        if (character is not null && character.EnumerateRunes().Count() != 1)
            return new TextFrameOperationResult { ErrorMessage = "character must be exactly one character." };
        PowerPoint.PpNumberedBulletStyle parsedNumberStyle = PowerPoint.PpNumberedBulletStyle.ppBulletArabicPeriod;
        if (numberStyle is not null && !NumberingStyles.TryGetValue(numberStyle, out parsedNumberStyle))
            return new TextFrameOperationResult { ErrorMessage = $"number_style '{numberStyle}' is not one of {string.Join(", ", NumberingStyles.Keys)}." };
        if (startAt is < 1 or > 32767)
            return new TextFrameOperationResult { ErrorMessage = "start_at must be between 1 and 32767." };
        if ((character is not null && bulletStyle is "none" or "numbered") || ((numberStyle is not null || startAt is not null) && bulletStyle is "none" or "bullet"))
            return new TextFrameOperationResult { ErrorMessage = "character goes with bullet_style=bullet; number_style and start_at go with bullet_style=numbered." };
        if (alignment is null && spaceBefore is null && spaceAfter is null && lineSpacing is null && lineSpacingPoints is null && indentLevel is null
            && leftIndent is null && firstLineIndent is null && bulletStyle is null && character is null && numberStyle is null && startAt is null)
            return new TextFrameOperationResult { ErrorMessage = "Pass at least one paragraph setting (alignment, space_before, space_after, line_spacing, indent_level, left_indent, first_line_indent, bullet_style, character, number_style, start_at)." };

        return WithTextShape(batch, slideIndex, shapeIndex, (presentation, shape, ct) =>
        {
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            PowerPoint.TextRange? target = null;
            PowerPoint.ParagraphFormat? format = null;
            PowerPoint.BulletFormat? bullet = null;
            PowerPoint.TextFrame2? frame2 = null;
            Office.TextRange2? range2 = null;
            Office.TextRange2? target2 = null;
            Office.ParagraphFormat2? format2 = null;
            try
            {
                frame = shape.TextFrame;
                range = frame.TextRange;
                int first, count;
                try
                {
                    (first, count) = TextRanges.Resolve(range.Text, null, null, null, 1, false, paragraph, paragraphCount);
                }
                catch (ArgumentException ex)
                {
                    return new TextFrameOperationResult { ErrorMessage = ex.Message };
                }

                // An empty range (an empty paragraph) still carries paragraph formatting.
                target = range.Characters(first, Math.Max(count, 0));
                format = target.ParagraphFormat;
                if (alignment is not null) format.Alignment = parsedAlignment;
                if (spaceBefore is { } before)
                {
                    format.LineRuleBefore = Office.MsoTriState.msoFalse;
                    format.SpaceBefore = before;
                }
                if (spaceAfter is { } after)
                {
                    format.LineRuleAfter = Office.MsoTriState.msoFalse;
                    format.SpaceAfter = after;
                }
                if (lineSpacing is { } lines)
                {
                    format.LineRuleWithin = Office.MsoTriState.msoTrue;
                    format.SpaceWithin = lines;
                }
                if (lineSpacingPoints is { } points)
                {
                    format.LineRuleWithin = Office.MsoTriState.msoFalse;
                    format.SpaceWithin = points;
                }
                if (indentLevel is { } level)
                    target.IndentLevel = level;

                string? effectiveBullet = bulletStyle ?? (character is not null ? "bullet" : numberStyle is not null || startAt is not null ? "numbered" : null);
                if (effectiveBullet is not null)
                {
                    bullet = format.Bullet;
                    if (effectiveBullet == "none")
                    {
                        bullet.Visible = Office.MsoTriState.msoFalse;
                    }
                    else if (effectiveBullet == "bullet")
                    {
                        bullet.Visible = Office.MsoTriState.msoTrue;
                        bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                        if (character is not null)
                            bullet.Character = char.ConvertToUtf32(character, 0);
                    }
                    else
                    {
                        bullet.Visible = Office.MsoTriState.msoTrue;
                        bullet.Type = PowerPoint.PpBulletType.ppBulletNumbered;
                        if (numberStyle is not null)
                            bullet.Style = parsedNumberStyle;
                        if (startAt is { } number)
                            bullet.StartValue = number;
                    }
                }

                if (leftIndent is not null || firstLineIndent is not null)
                {
                    // Indents in points are only on the TextFrame2 paragraph format.
                    frame2 = shape.TextFrame2;
                    range2 = frame2.TextRange;
                    target2 = range2.Characters[first, Math.Max(count, 0)];
                    format2 = target2.ParagraphFormat;
                    if (leftIndent is { } left) format2.LeftIndent = left;
                    if (firstLineIndent is { } firstLine) format2.FirstLineIndent = firstLine;
                }

                return new TextFrameOperationResult { Success = true, Start = first, Length = count };
            }
            finally
            {
                if (format2 is not null) ComUtilities.Release(ref format2);
                if (target2 is not null) ComUtilities.Release(ref target2);
                if (range2 is not null) ComUtilities.Release(ref range2);
                if (frame2 is not null) ComUtilities.Release(ref frame2);
                if (bullet is not null) ComUtilities.Release(ref bullet);
                if (format is not null) ComUtilities.Release(ref format);
                if (target is not null) ComUtilities.Release(ref target);
                if (range is not null) ComUtilities.Release(ref range);
                if (frame is not null) ComUtilities.Release(ref frame);
            }
        });
    }

    /// <inheritdoc/>
    public TextFrameOperationResult SetTextFrame(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        float? marginLeft = null,
        float? marginTop = null,
        float? marginRight = null,
        float? marginBottom = null,
        string? verticalAnchor = null,
        bool? wordWrap = null,
        string? autoSize = null,
        string? orientation = null,
        int? columnCount = null,
        float? columnSpacing = null,
        float? rotation = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (marginLeft is < 0 || marginTop is < 0 || marginRight is < 0 || marginBottom is < 0)
            return new TextFrameOperationResult { ErrorMessage = "Margins must be 0 or greater (points)." };
        Office.MsoVerticalAnchor anchor = Office.MsoVerticalAnchor.msoAnchorTop;
        if (verticalAnchor is not null && !Anchors.TryGetValue(verticalAnchor, out anchor))
            return new TextFrameOperationResult { ErrorMessage = $"vertical_anchor '{verticalAnchor}' is not top, middle, or bottom." };
        PowerPoint.PpAutoSize parsedAutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
        if (autoSize is not null && !TryParseAutoSize(autoSize, out parsedAutoSize))
            return new TextFrameOperationResult { ErrorMessage = $"auto_size '{autoSize}' is not none, shape-to-fit-text, or shrink-on-overflow." };
        Office.MsoTextOrientation parsedOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;
        if (orientation is not null && !Orientations.TryGetValue(orientation, out parsedOrientation))
            return new TextFrameOperationResult { ErrorMessage = $"orientation '{orientation}' is not horizontal, up, down, or stacked." };
        if (columnCount is < 1 or > 16)
            return new TextFrameOperationResult { ErrorMessage = "column_count must be between 1 and 16." };
        if (columnSpacing is < 0)
            return new TextFrameOperationResult { ErrorMessage = "column_spacing must be 0 or greater." };
        if (rotation is < -360 or > 360)
            return new TextFrameOperationResult { ErrorMessage = "rotation must be between -360 and 360 degrees." };

        return WithTextShape(batch, slideIndex, shapeIndex, (presentation, shape, ct) =>
        {
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextFrame2? frame2 = null;
            Office.TextColumn2? columns = null;
            try
            {
                frame = shape.TextFrame;
                frame2 = shape.TextFrame2;
                if (marginLeft is { } left) frame.MarginLeft = left;
                if (marginTop is { } top) frame.MarginTop = top;
                if (marginRight is { } right) frame.MarginRight = right;
                if (marginBottom is { } bottom) frame.MarginBottom = bottom;
                if (verticalAnchor is not null) frame.VerticalAnchor = anchor;
                if (wordWrap is { } wrap) frame.WordWrap = wrap ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (autoSize is not null) frame.AutoSize = parsedAutoSize;
                if (orientation is not null) frame.Orientation = parsedOrientation;
                if (columnCount is not null || columnSpacing is not null)
                {
                    columns = frame2.Column;
                    if (columnCount is { } number) columns.Number = number;
                    if (columnSpacing is { } spacing) columns.Spacing = spacing;
                }
                if (rotation is { } degrees)
                    shape.Rotation = ((degrees % 360) + 360) % 360;

                return new TextFrameOperationResult { Success = true, Frame = ReadLayout(shape, frame, frame2) };
            }
            finally
            {
                if (columns is not null) ComUtilities.Release(ref columns);
                if (frame2 is not null) ComUtilities.Release(ref frame2);
                if (frame is not null) ComUtilities.Release(ref frame);
            }
        });
    }

    /// <summary>Runs an action on a validated shape that has a text frame; releases the COM path.</summary>
    private static TextFrameOperationResult WithTextShape(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        Func<PowerPoint.Presentation, PowerPoint.Shape, CancellationToken, TextFrameOperationResult> action)
    {
        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            try
            {
                slides = ctx.Presentation.Slides;
                if (slideIndex < 1 || slideIndex > slides.Count)
                    return new TextFrameOperationResult { ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slides.Count} slide(s) (valid range: 1-{slides.Count})." };
                slide = slides[slideIndex];
                shapes = slide.Shapes;
                if (shapeIndex < 1 || shapeIndex > shapes.Count)
                    return new TextFrameOperationResult { ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapes.Count} shape(s) (valid range: 1-{shapes.Count})." };
                shape = shapes[shapeIndex];
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                    return new TextFrameOperationResult { ErrorMessage = $"Shape {shapeIndex} ('{shape.Name}') has no text frame. Tables are edited with the table tool; use deck find-text to locate text." };
                return action(ctx.Presentation, shape, ct);
            }
            finally
            {
                if (shape is not null) ComUtilities.Release(ref shape);
                if (shapes is not null) ComUtilities.Release(ref shapes);
                if (slide is not null) ComUtilities.Release(ref slide);
                if (slides is not null) ComUtilities.Release(ref slides);
            }
        });
    }

    private static TextParagraphInfo ReadParagraph(PowerPoint.TextRange range, Office.TextRange2 range2, int index, int start, int length, string text, List<TextRunInfo> runs)
    {
        PowerPoint.TextRange? target = null;
        PowerPoint.ParagraphFormat? format = null;
        PowerPoint.BulletFormat? bullet = null;
        Office.TextRange2? target2 = null;
        Office.ParagraphFormat2? format2 = null;
        try
        {
            target = range.Paragraphs(index, 1);
            format = target.ParagraphFormat;
            bullet = format.Bullet;
            target2 = range2.Paragraphs[index, 1];
            format2 = target2.ParagraphFormat;
            bool lines = format.LineRuleWithin == Office.MsoTriState.msoTrue;
            string bulletKind = bullet.Visible != Office.MsoTriState.msoTrue
                ? "none"
                : bullet.Type switch
                {
                    PowerPoint.PpBulletType.ppBulletNumbered => "numbered",
                    PowerPoint.PpBulletType.ppBulletPicture => "picture",
                    PowerPoint.PpBulletType.ppBulletNone => "none",
                    _ => "bullet",
                };
            return new TextParagraphInfo
            {
                Index = index,
                Start = start,
                Length = length,
                Text = text.Replace('\v', '\n'),
                Alignment = format.Alignment switch
                {
                    PowerPoint.PpParagraphAlignment.ppAlignLeft => "left",
                    PowerPoint.PpParagraphAlignment.ppAlignCenter => "center",
                    PowerPoint.PpParagraphAlignment.ppAlignRight => "right",
                    PowerPoint.PpParagraphAlignment.ppAlignJustify => "justify",
                    PowerPoint.PpParagraphAlignment.ppAlignDistribute => "distribute",
                    _ => "mixed",
                },
                IndentLevel = target.IndentLevel,
                SpaceBefore = Round(format.SpaceBefore),
                SpaceAfter = Round(format.SpaceAfter),
                LineSpacing = lines ? Round(format.SpaceWithin) : null,
                LineSpacingPoints = lines ? null : Round(format.SpaceWithin),
                LeftIndent = Round(format2.LeftIndent),
                FirstLineIndent = Round(format2.FirstLineIndent),
                Bullet = bulletKind,
                BulletCharacter = bulletKind == "bullet" && bullet.Character > 0 ? char.ConvertFromUtf32(bullet.Character) : null,
                NumberStyle = bulletKind == "numbered" ? NumberingStyles.FirstOrDefault(pair => pair.Value == bullet.Style).Key ?? bullet.Style.ToString() : null,
                StartAt = bulletKind == "numbered" ? bullet.StartValue : null,
                Runs = runs,
            };
        }
        finally
        {
            if (format2 is not null) ComUtilities.Release(ref format2);
            if (target2 is not null) ComUtilities.Release(ref target2);
            if (bullet is not null) ComUtilities.Release(ref bullet);
            if (format is not null) ComUtilities.Release(ref format);
            if (target is not null) ComUtilities.Release(ref target);
        }
    }

    /// <summary>Appends the runs of one paragraph; returns how many were read.</summary>
    private static int ReadRuns(PowerPoint.Presentation presentation, PowerPoint.TextRange range, Office.TextRange2 range2, int start, int length, List<TextRunInfo> runs, int budget)
    {
        PowerPoint.TextRange? paragraphRange = null;
        PowerPoint.TextRange? allRuns = null;
        try
        {
            paragraphRange = range.Characters(start, length);
            allRuns = paragraphRange.Runs();
            int count = Math.Min(allRuns.Count, budget);
            for (int index = 1; index <= count; index++)
            {
                PowerPoint.TextRange? run = null;
                PowerPoint.Font? font = null;
                PowerPoint.ColorFormat? color = null;
                try
                {
                    run = paragraphRange.Runs(index, 1);
                    font = run.Font;
                    color = font.Color;
                    runs.Add(new TextRunInfo
                    {
                        Start = run.Start,
                        Length = run.Length,
                        Text = run.Text.Replace('\v', '\n'),
                        FontName = font.Name,
                        FontSize = font.Size,
                        Bold = font.Bold == Office.MsoTriState.msoTrue,
                        Italic = font.Italic == Office.MsoTriState.msoTrue,
                        Underline = font.Underline == Office.MsoTriState.msoTrue,
                        Color = DeckSnapshotReader.Hex(color.RGB),
                        Hyperlink = ReadHyperlink(presentation, run),
                    });
                }
                finally
                {
                    if (color is not null) ComUtilities.Release(ref color);
                    if (font is not null) ComUtilities.Release(ref font);
                    if (run is not null) ComUtilities.Release(ref run);
                }
            }
            AddSpacing(range2, runs, runs.Count - count);
            return count;
        }
        finally
        {
            if (allRuns is not null) ComUtilities.Release(ref allRuns);
            if (paragraphRange is not null) ComUtilities.Release(ref paragraphRange);
        }
    }

    /// <summary>Fills character spacing and baseline offset (TextRange2 only) for the runs just read.</summary>
    private static void AddSpacing(Office.TextRange2 all2, List<TextRunInfo> runs, int from)
    {
        for (int index = from; index < runs.Count; index++)
        {
            Office.TextRange2? run2 = null;
            Office.Font2? font2 = null;
            try
            {
                var run = runs[index];
                run2 = all2.Characters[run.Start, run.Length];
                font2 = run2.Font;
                float spacing = Round(font2.Spacing);
                float offset = Round(font2.BaselineOffset);
                if (spacing != 0 || offset != 0)
                {
                    runs[index] = new TextRunInfo
                    {
                        Start = run.Start,
                        Length = run.Length,
                        Text = run.Text,
                        FontName = run.FontName,
                        FontSize = run.FontSize,
                        Bold = run.Bold,
                        Italic = run.Italic,
                        Underline = run.Underline,
                        Color = run.Color,
                        Hyperlink = run.Hyperlink,
                        CharacterSpacing = spacing != 0 ? spacing : null,
                        BaselineOffset = offset != 0 ? offset : null,
                    };
                }
            }
            finally
            {
                if (font2 is not null) ComUtilities.Release(ref font2);
                if (run2 is not null) ComUtilities.Release(ref run2);
            }
        }
    }

    private static string? ReadHyperlink(PowerPoint.Presentation presentation, PowerPoint.TextRange run)
    {
        PowerPoint.ActionSettings? settings = null;
        PowerPoint.ActionSetting? click = null;
        PowerPoint.Hyperlink? link = null;
        try
        {
            settings = run.ActionSettings;
            click = settings[PowerPoint.PpMouseActivation.ppMouseClick];
            if (click.Action != PowerPoint.PpActionType.ppActionHyperlink)
                return null;
            link = click.Hyperlink;
            if (!string.IsNullOrEmpty(link.Address))
                return link.Address;
            return SlideLinkText(presentation, link.SubAddress);
        }
        finally
        {
            if (link is not null) ComUtilities.Release(ref link);
            if (click is not null) ComUtilities.Release(ref click);
            if (settings is not null) ComUtilities.Release(ref settings);
        }
    }

    /// <summary>"slide:N" for a SubAddress "SlideID,SlideIndex,Title" (resolved by SlideID, which survives reordering).</summary>
    private static string? SlideLinkText(PowerPoint.Presentation presentation, string? subAddress)
    {
        if (string.IsNullOrEmpty(subAddress))
            return null;
        var parts = subAddress.Split(',');
        if (int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var slideId))
        {
            var slide = DeckShapeLocator.FindSlide(presentation, slideId);
            if (slide is not null)
            {
                try
                {
                    return $"slide:{slide.SlideIndex}";
                }
                finally
                {
                    ComUtilities.Release(ref slide!);
                }
            }
        }
        return "#" + subAddress;
    }

    private static string? WriteHyperlink(PowerPoint.Presentation presentation, PowerPoint.TextRange target, string hyperlink)
    {
        PowerPoint.ActionSettings? settings = null;
        PowerPoint.ActionSetting? click = null;
        PowerPoint.Hyperlink? link = null;
        PowerPoint.Slides? slides = null;
        PowerPoint.Slide? slide = null;
        try
        {
            settings = target.ActionSettings;
            click = settings[PowerPoint.PpMouseActivation.ppMouseClick];
            if (hyperlink.Length == 0)
            {
                if (click.Action == PowerPoint.PpActionType.ppActionHyperlink)
                {
                    link = click.Hyperlink;
                    link.Delete();
                }
                return null;
            }

            link = click.Hyperlink;
            if (hyperlink.StartsWith("slide:", StringComparison.OrdinalIgnoreCase))
            {
                slides = presentation.Slides;
                if (!int.TryParse(hyperlink.AsSpan(6), NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) || index < 1 || index > slides.Count)
                    return $"hyperlink '{hyperlink}' names no slide; use slide:1 to slide:{slides.Count}.";
                slide = slides[index];
                link.Address = "";
                link.SubAddress = $"{slide.SlideID},{index},";
                return null;
            }
            if (hyperlink.Any(char.IsControl))
                return "hyperlink must not contain control characters.";
            link.Address = hyperlink;
            return null;
        }
        finally
        {
            if (slide is not null) ComUtilities.Release(ref slide);
            if (slides is not null) ComUtilities.Release(ref slides);
            if (link is not null) ComUtilities.Release(ref link);
            if (click is not null) ComUtilities.Release(ref click);
            if (settings is not null) ComUtilities.Release(ref settings);
        }
    }

    private static TextFrameLayoutInfo ReadLayout(PowerPoint.Shape shape, PowerPoint.TextFrame frame, PowerPoint.TextFrame2 frame2)
    {
        Office.TextColumn2? columns = null;
        try
        {
            columns = frame2.Column;
            return new TextFrameLayoutInfo
            {
                Margins = [Round(frame.MarginLeft), Round(frame.MarginTop), Round(frame.MarginRight), Round(frame.MarginBottom)],
                VerticalAnchor = Anchors.FirstOrDefault(pair => pair.Value == frame.VerticalAnchor).Key ?? "mixed",
                WordWrap = frame.WordWrap == Office.MsoTriState.msoTrue,
                AutoSize = frame2.AutoSize switch
                {
                    Office.MsoAutoSize.msoAutoSizeNone => "none",
                    Office.MsoAutoSize.msoAutoSizeShapeToFitText => "shape-to-fit-text",
                    Office.MsoAutoSize.msoAutoSizeTextToFitShape => "shrink-on-overflow",
                    _ => "mixed",
                },
                Orientation = Orientations.FirstOrDefault(pair => pair.Value == frame.Orientation).Key ?? "mixed",
                Columns = columns.Number,
                ColumnSpacing = Round(columns.Spacing),
                Rotation = Round(shape.Rotation),
            };
        }
        finally
        {
            if (columns is not null) ComUtilities.Release(ref columns);
        }
    }

    /// <summary>Parses #RRGGBB (or RRGGBB) into PowerPoint's BGR integer.</summary>
    internal static bool TryParseHexColor(string value, out int bgr)
    {
        bgr = 0;
        var hex = value.Trim().TrimStart('#');
        if (hex.Length != 6 || !int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            return false;
        bgr = ((rgb >> 16) & 0xFF) | (rgb & 0xFF00) | ((rgb & 0xFF) << 16);
        return true;
    }

    private static float Round(float value) => MathF.Round(value * 10f) / 10f;
}
