using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Table;

/// <summary>
/// Table commands: add a table shape, read/write cell text, fill a whole table in one call,
/// insert/delete rows and columns, format cell fill and borders, and merge cells. Operates within
/// an already-open IPresentationBatch, targeting a specific slide and table shape by their 1-based
/// indices.
/// </summary>
[ServiceCategory("table", "Table")]
[McpTool("table", Title = "Table Operations", Destructive = true, Category = "content",
    Description = "Add a table shape, read/write cell text, fill a whole table from rows in one call (set-data), edit rows/columns, and format cells in an open presentation session.")]
[McpReadOnlyActions("get-cell-text", "get-cell-fill", "get-cell-border")]
public interface ITableCommands
{
    /// <summary>Adds a new table shape with the given number of rows/columns to a slide.</summary>
    TableOperationResult AddTable(IPresentationBatch batch, int slideIndex, int rows, int columns, float left, float top, float width, float height);

    /// <summary>Sets the text of a table cell (1-based row/column).</summary>
    TableOperationResult SetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string text);

    /// <summary>Gets the text of a table cell (1-based row/column).</summary>
    TableOperationResult GetCellText(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column);

    /// <summary>
    /// Fills table cells from rows of separated cells in one call, e.g. <c>["Region|Q1|Q2",
    /// "North|12|15"]</c>; pasted markdown rows (<c>| a | b |</c>, <c>|---|---|</c>) also work.
    /// Writing starts at <paramref name="startRow"/>/<paramref name="startColumn"/>. Everything is
    /// validated before any cell is written; a table too small for the data is rejected unchanged.
    /// </summary>
    /// <param name="data">One entry per table row with its cells separated by the separator (a vertical bar by default). Pasted markdown table rows, including the header divider row, are accepted; put a backslash before the separator to use it literally inside a cell.</param>
    /// <param name="separator">Cell separator. Defaults to the vertical bar character.</param>
    /// <param name="startRow">1-based table row that receives the first data row. Defaults to 1.</param>
    /// <param name="startColumn">1-based table column that receives the first cell of each row. Defaults to 1.</param>
    TableOperationResult SetData(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        IReadOnlyList<string> data,
        string separator = "|",
        int startRow = 1,
        int startColumn = 1);

    /// <summary>
    /// Inserts a new row into the table. If <paramref name="beforeRow"/> is omitted, the row is
    /// appended as the last row. Returns the new total <c>rowCount</c>.
    /// </summary>
    TableOperationResult InsertRow(IPresentationBatch batch, int slideIndex, int shapeIndex, int? beforeRow = null);

    /// <summary>Deletes the row at the given 1-based index. Returns the new total <c>rowCount</c>.</summary>
    TableOperationResult DeleteRow(IPresentationBatch batch, int slideIndex, int shapeIndex, int row);

    /// <summary>
    /// Inserts a new column into the table. If <paramref name="beforeColumn"/> is omitted, the
    /// column is appended as the last column. Returns the new total <c>columnCount</c>.
    /// </summary>
    TableOperationResult InsertColumn(IPresentationBatch batch, int slideIndex, int shapeIndex, int? beforeColumn = null);

    /// <summary>Deletes the column at the given 1-based index. Returns the new total <c>columnCount</c>.</summary>
    TableOperationResult DeleteColumn(IPresentationBatch batch, int slideIndex, int shapeIndex, int column);

    /// <summary>Sets a table cell's fill to a solid RGB color.</summary>
    TableOperationResult SetCellFill(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, byte red, byte green, byte blue);

    /// <summary>Gets a table cell's solid fill color.</summary>
    TableOperationResult GetCellFill(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column);

    /// <summary>
    /// Sets one or more properties of a single border of a table cell. <paramref name="borderType"/>
    /// is a <c>PpBorderType</c> enum member name (<c>"ppBorderTop"</c>, <c>"ppBorderBottom"</c>,
    /// <c>"ppBorderLeft"</c>, <c>"ppBorderRight"</c>, <c>"ppBorderDiagonalDown"</c>, or
    /// <c>"ppBorderDiagonalUp"</c>). Any other parameter left null is unchanged; passing
    /// <paramref name="red"/>/<paramref name="green"/>/<paramref name="blue"/> together sets the
    /// border color; <paramref name="dashStyle"/> is an <c>MsoLineDashStyle</c> enum member name.
    /// </summary>
    TableOperationResult SetCellBorder(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        int row,
        int column,
        string borderType,
        byte? red = null,
        byte? green = null,
        byte? blue = null,
        float? weight = null,
        string? dashStyle = null,
        bool? visible = null);

    /// <summary>Gets a single border's color, weight, dash style, and visibility for a table cell.</summary>
    TableOperationResult GetCellBorder(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, string borderType);

    /// <summary>
    /// Merges the cell at (<paramref name="row"/>, <paramref name="column"/>) with the cell at
    /// (<paramref name="mergeToRow"/>, <paramref name="mergeToColumn"/>), producing a single
    /// merged cell. The two cells must be adjacent (in the same row or column).
    /// </summary>
    TableOperationResult MergeCells(IPresentationBatch batch, int slideIndex, int shapeIndex, int row, int column, int mergeToRow, int mergeToColumn);

    /// <summary>
    /// Styles an existing table with a design profile's table style (header fill and text, banded
    /// rows, total row, font, padding, borders) without changing any cell text.
    /// </summary>
    /// <param name="profile">Design profile (default theme).</param>
    /// <param name="headerRow">Style row 1 as a header (default true).</param>
    /// <param name="totalRow">Style the last row as a total (default false).</param>
    /// <param name="banded">Alternate body row fills (default true).</param>
    /// <param name="fontSize">Text size in points (apply-style default: the profile's table size).</param>
    TableOperationResult ApplyStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, string? profile = null, bool headerRow = true, bool totalRow = false,
        bool banded = true, float? fontSize = null);

    /// <summary>Sets column widths in points, or as relative weights that fill the table's current width.</summary>
    /// <param name="widths">One value per column.</param>
    /// <param name="mode">points (default) or weights.</param>
    TableOperationResult SetColumnWidths(IPresentationBatch batch, int slideIndex, int shapeIndex, IReadOnlyList<double> widths, string mode = "points");

    /// <summary>
    /// Re-formats the numbers in one column with an Excel number format (e.g. #,##0.0, 0%, "$"0.0"M").
    /// Cell text is parsed with the culture; cells that are not numbers are left unchanged and
    /// listed. Formatting of the text runs is kept.
    /// </summary>
    /// <param name="format">Excel number format code.</param>
    /// <param name="culture">Culture for parsing and output separators: invariant (default), ru-RU, en-US, ...</param>
    /// <param name="firstRow">First row (1-based); format-numbers and conditional-format default to 2, below the header.</param>
    TableOperationResult FormatNumbers(IPresentationBatch batch, int slideIndex, int shapeIndex, int column, string format, string culture = "invariant", int firstRow = 2);

    /// <summary>
    /// Highlights cells in one column by rule: negative, positive, above, below, equals (with
    /// threshold), top or bottom (with count). Non-numeric cells are skipped and listed.
    /// </summary>
    /// <param name="rule">negative, positive, above, below, equals, top, or bottom.</param>
    /// <param name="threshold">Comparison value for above, below, equals.</param>
    /// <param name="count">How many cells for top and bottom (default 3).</param>
    /// <param name="textColor">#RRGGBB text color.</param>
    /// <param name="fillColor">#RRGGBB cell fill.</param>
    /// <param name="bold">Bold text.</param>
    TableOperationResult ConditionalFormat(IPresentationBatch batch, int slideIndex, int shapeIndex, int column, string rule, double? threshold = null, int count = 3,
        string? textColor = null, string? fillColor = null, bool? bold = null, string culture = "invariant", int firstRow = 2);

    /// <summary>Styles a block of cells: fill, text color, bold, size, horizontal and vertical alignment, and padding.</summary>
    /// <param name="lastRow">Last row of the block (inclusive).</param>
    /// <param name="firstColumn">First column of the block.</param>
    /// <param name="lastColumn">Last column of the block (inclusive).</param>
    /// <param name="align">left, center, or right.</param>
    /// <param name="verticalAlign">top, middle, or bottom.</param>
    /// <param name="padding">Cell padding in points on all sides.</param>
    TableOperationResult SetRangeStyle(IPresentationBatch batch, int slideIndex, int shapeIndex, int firstRow, int lastRow, int firstColumn, int lastColumn,
        string? fillColor = null, string? textColor = null, bool? bold = null, float? fontSize = null, string? align = null, string? verticalAlign = null, float? padding = null);
}
