using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

/// <summary>
/// Text frame commands: set/get/find/replace text and basic font formatting (size, bold, italic, underline,
/// font name, color, alignment, bullets) for a shape's text range. Operates within an
/// already-open IPresentationBatch, targeting a specific shape by its 1-based slide and shape
/// index.
/// </summary>
[ServiceCategory("textframe", "TextFrame")]
[McpTool("textframe", Title = "Text Frame Operations", Destructive = true, Category = "content",
    Description = "Set, get, find, or replace text and font/paragraph formatting in one shape's text frame. get-paragraphs shows paragraphs and formatting runs with 1-based positions; replace-range and format-range edit part of the text (by start/length, the Nth match, or paragraphs) while other runs keep their formatting; set-paragraph-format sets alignment, spacing, indents, bullets, and numbering; set-text-frame sets margins, anchoring, wrap, autofit, orientation, columns, and rotation. For text across many shapes or slides use deck find-text / replace-text.")]
[McpReadOnlyActions("get-text", "find-text", "get-font-size", "get-bold", "get-font-color",
    "get-italic", "get-underline", "get-font-name", "get-alignment", "get-bullet", "get-auto-size", "get-paragraphs")]
public interface ITextFrameCommands
{
    /// <summary>Sets the text content of a shape's text frame.</summary>
    TextFrameOperationResult SetText(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        [AllowEmptyString] string text);

    /// <summary>Gets the text content of a shape's text frame.</summary>
    TextFrameOperationResult GetText(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Finds all non-overlapping literal matches in one shape's text frame without mutation.
    /// Returns ascending 1-based PowerPoint character positions, lengths, and original matched text.
    /// Empty search text is invalid. Matching uses PowerPoint case and whole-word rules, not regex.
    /// </summary>
    TextFrameOperationResult FindText(IPresentationBatch batch, int slideIndex, int shapeIndex,
        [AllowEmptyString] string findWhat, bool matchCase = false, bool wholeWords = false);

    /// <summary>
    /// Replaces all non-overlapping literal matches in one shape's text frame. Empty replacement
    /// deletes matches; empty search text is invalid. Uses PowerPoint case/whole-word rules.
    /// Does not revisit inserted text or rewrite the whole frame. Unexpected failures are not transactional.
    /// </summary>
    TextFrameOperationResult ReplaceText(IPresentationBatch batch, int slideIndex, int shapeIndex,
        [AllowEmptyString] string findWhat, [AllowEmptyString] string replaceWhat, bool matchCase = false, bool wholeWords = false);

    /// <summary>Sets the font size (in points) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontSize(IPresentationBatch batch, int slideIndex, int shapeIndex, float fontSize);

    /// <summary>Gets the font size (in points) of a shape's text range.</summary>
    TextFrameOperationResult GetFontSize(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets whether a shape's entire text range is bold.</summary>
    TextFrameOperationResult SetBold(IPresentationBatch batch, int slideIndex, int shapeIndex, bool bold);

    /// <summary>Gets whether a shape's text range is bold.</summary>
    TextFrameOperationResult GetBold(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the font color (RGB) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontColor(IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue);

    /// <summary>Gets the font color of a shape's text range as an RGB integer (0xBBGGRR, PowerPoint's native color order).</summary>
    TextFrameOperationResult GetFontColor(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets whether a shape's entire text range is italic.</summary>
    TextFrameOperationResult SetItalic(IPresentationBatch batch, int slideIndex, int shapeIndex, bool italic);

    /// <summary>Gets whether a shape's text range is italic.</summary>
    TextFrameOperationResult GetItalic(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets whether a shape's entire text range is underlined.</summary>
    TextFrameOperationResult SetUnderline(IPresentationBatch batch, int slideIndex, int shapeIndex, bool underline);

    /// <summary>Gets whether a shape's text range is underlined.</summary>
    TextFrameOperationResult GetUnderline(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>Sets the font name (typeface) of a shape's entire text range.</summary>
    TextFrameOperationResult SetFontName(IPresentationBatch batch, int slideIndex, int shapeIndex, string fontName);

    /// <summary>Gets the font name (typeface) of a shape's text range.</summary>
    TextFrameOperationResult GetFontName(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets the paragraph alignment of a shape's entire text range, identified by its
    /// <c>PpParagraphAlignment</c> enum member name (<c>"ppAlignLeft"</c>, <c>"ppAlignCenter"</c>,
    /// <c>"ppAlignRight"</c>, <c>"ppAlignJustify"</c>, or <c>"ppAlignDistribute"</c>).
    /// </summary>
    TextFrameOperationResult SetAlignment(IPresentationBatch batch, int slideIndex, int shapeIndex, string alignment);

    /// <summary>Gets the paragraph alignment of a shape's text range as a <c>PpParagraphAlignment</c> enum member name.</summary>
    TextFrameOperationResult GetAlignment(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Turns bullets on or off for a shape's entire text range. When <paramref name="enabled"/> is
    /// <c>true</c>, an optional single-character <paramref name="character"/> sets the bullet
    /// glyph (defaults to PowerPoint's theme bullet character if omitted).
    /// </summary>
    TextFrameOperationResult SetBullet(IPresentationBatch batch, int slideIndex, int shapeIndex, bool enabled, string? character = null);

    /// <summary>Gets whether a shape's text range has bullets enabled, and the bullet character if so.</summary>
    TextFrameOperationResult GetBullet(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Sets how a shape's text frame automatically resizes text/shape, identified by its
    /// <c>PpAutoSize</c> enum member name (<c>"ppAutoSizeNone"</c>, <c>"ppAutoSizeShapeToFitText"</c>,
    /// or <c>"ppAutoSizeTextToFitShape"</c>).
    /// </summary>
    TextFrameOperationResult SetAutoSize(IPresentationBatch batch, int slideIndex, int shapeIndex, string autoSize);

    /// <summary>Gets a shape's text frame auto-size mode as a <c>PpAutoSize</c> enum member name.</summary>
    TextFrameOperationResult GetAutoSize(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Lists the frame's paragraphs (1-based start, length, text, alignment, indent level,
    /// spacing, indents, bullet or numbering) with their formatting runs (font, size, bold,
    /// italic, underline, #RRGGBB color, hyperlink, character spacing, baseline offset), plus the
    /// frame's margins, anchoring, wrap, autofit, orientation, columns, and rotation. Positions
    /// are what replace-range and format-range take.
    /// </summary>
    TextFrameOperationResult GetParagraphs(IPresentationBatch batch, int slideIndex, int shapeIndex);

    /// <summary>
    /// Replaces part of the text, chosen by start and length, by match and occurrence, or by
    /// paragraph, without touching the rest. The new text takes the formatting of the first
    /// replaced character (inserted text, with length 0, takes the formatting of the character it
    /// is inserted before). A newline in the text starts a new paragraph.
    /// </summary>
    /// <param name="text">Text to write; for replace-range an empty text deletes the selected characters and a newline starts a new paragraph.</param>
    /// <param name="start">1-based first character (from get-paragraphs or find-text); text length + 1 appends.</param>
    /// <param name="length">Number of characters from start (default: to the end of the text).</param>
    /// <param name="match">Select an occurrence of this literal text instead of start/length.</param>
    /// <param name="occurrence">Which occurrence of match (1-based, default 1; -1 for the last).</param>
    /// <param name="matchCase">Match upper/lower case exactly (default false).</param>
    /// <param name="paragraph">Select whole paragraphs starting at this 1-based paragraph (default: the whole frame).</param>
    /// <param name="paragraphCount">Number of paragraphs from paragraph (default 1).</param>
    TextFrameOperationResult ReplaceRange(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        [AllowEmptyString] string text,
        int? start = null,
        int? length = null,
        string? match = null,
        int occurrence = 1,
        bool matchCase = false,
        int? paragraph = null,
        int? paragraphCount = null);

    /// <summary>
    /// Formats part of the text (by start/length, match/occurrence, or paragraph; default the
    /// whole frame): font, size, bold, italic, underline, #RRGGBB color, hyperlink, character
    /// spacing, and superscript/subscript offset. Only the options passed change; other runs keep
    /// their formatting.
    /// </summary>
    /// <param name="fontName">Typeface name, e.g. Segoe UI.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="bold">Bold on or off.</param>
    /// <param name="italic">Italic on or off.</param>
    /// <param name="underline">Underline on or off.</param>
    /// <param name="color">Text color as #RRGGBB.</param>
    /// <param name="hyperlink">Link address (https://..., mailto:..., a file path) or slide:N for another slide; empty removes the link.</param>
    /// <param name="characterSpacing">Character spacing in points (0 normal, negative condensed, positive expanded).</param>
    /// <param name="baselineOffset">Baseline offset from -1 to 1 (0.3 superscript, -0.25 subscript, 0 normal).</param>
    TextFrameOperationResult FormatRange(
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
        float? baselineOffset = null);

    /// <summary>
    /// Sets paragraph formatting for one or more paragraphs (default all): alignment, space
    /// before/after, line spacing (in lines or points), outline level, left and first-line
    /// indents, and bullets or numbering. Only the options passed change.
    /// </summary>
    /// <param name="alignment">left, center, right, justify, or distribute (ppAlign* names are also accepted).</param>
    /// <param name="spaceBefore">Space before the paragraph in points.</param>
    /// <param name="spaceAfter">Space after the paragraph in points.</param>
    /// <param name="lineSpacing">Line spacing as a multiple of single spacing, e.g. 1.15.</param>
    /// <param name="lineSpacingPoints">Exact line spacing in points (instead of line_spacing).</param>
    /// <param name="indentLevel">Outline level 1-9.</param>
    /// <param name="leftIndent">Left indent in points.</param>
    /// <param name="firstLineIndent">First-line indent in points; negative makes a hanging indent for bullets.</param>
    /// <param name="bulletStyle">none, bullet, or numbered.</param>
    /// <param name="character">Bullet glyph (one character), e.g. • or –.</param>
    /// <param name="numberStyle">Numbering: arabic-period (1.), arabic-paren (1)), alpha-lower-period (a.), alpha-upper-period (A.), alpha-lower-paren (a)), roman-lower-period (i.), roman-upper-period (I.).</param>
    /// <param name="startAt">First number for numbered paragraphs.</param>
    TextFrameOperationResult SetParagraphFormat(
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
        int? startAt = null);

    /// <summary>
    /// Sets text frame layout: margins, vertical anchoring, word wrap, autofit, text orientation,
    /// columns, and the shape's rotation. Only the options passed change. Returns the resulting
    /// layout.
    /// </summary>
    /// <param name="marginLeft">Left margin in points.</param>
    /// <param name="marginTop">Top margin in points.</param>
    /// <param name="marginRight">Right margin in points.</param>
    /// <param name="marginBottom">Bottom margin in points.</param>
    /// <param name="verticalAnchor">top, middle, or bottom.</param>
    /// <param name="wordWrap">Wrap text inside the shape.</param>
    /// <param name="autoSize">Autofit: none, shape-to-fit-text, or shrink-on-overflow (ppAutoSize* names are also accepted).</param>
    /// <param name="orientation">horizontal, up (rotated 270°), down (rotated 90°), or stacked.</param>
    /// <param name="columnCount">Number of text columns (1-16).</param>
    /// <param name="columnSpacing">Space between columns in points.</param>
    /// <param name="rotation">Shape rotation in degrees (0-360).</param>
    TextFrameOperationResult SetTextFrame(
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
        float? rotation = null);
}
