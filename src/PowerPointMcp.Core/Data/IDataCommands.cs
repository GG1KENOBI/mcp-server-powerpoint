using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>
/// Local data files (CSV and XLSX, read without Excel) to native tables and charts that remember
/// their source. Column types are inferred (number, percent, date, bool, text) with the culture's
/// separators; mixed columns stay text and are reported, never coerced. Refresh re-reads the source
/// and updates cell text and chart data in place, keeping formatting.
/// </summary>
[ServiceCategory("data", "Data")]
[McpTool("data", Title = "Data Import and Refresh", Destructive = true, Category = "content",
    Description = "Read local CSV/XLSX files (no Excel needed) with typed columns; create styled native tables (split onto continuation slides with repeated headers) and charts bound to the file; refresh bound tables and charts in place; list bindings and stale sources.")]
[McpReadOnlyActions("preview", "bindings")]
public interface IDataCommands
{
    /// <summary>Reads a data file and returns typed columns, notes, and the first rows.</summary>
    /// <param name="sourcePath">Full path of a .csv, .tsv, .txt, or .xlsx file.</param>
    /// <param name="sheet">XLSX worksheet name (default: first sheet).</param>
    /// <param name="range">XLSX cell range such as A1:D20 (default: used range).</param>
    /// <param name="header">First row is a header (default true).</param>
    /// <param name="delimiter">CSV delimiter (default: detected): one character or tab.</param>
    /// <param name="encoding">CSV encoding (default: detected), e.g. utf-8, windows-1251.</param>
    /// <param name="culture">Number and date culture: invariant (default), ru-RU, en-US, de-DE, ...</param>
    /// <param name="maxRows">Rows to return (1-200, default 20).</param>
    DataOperationResult Preview(IPresentationBatch batch, string sourcePath, string? sheet = null, string? range = null, bool header = true,
        string? delimiter = null, string? encoding = null, string culture = "invariant", int maxRows = 20);

    /// <summary>
    /// Creates a styled table slide from the file (a table composition), bound to the source.
    /// Long tables continue on new slides with the header repeated.
    /// </summary>
    /// <param name="title">Slide title.</param>
    /// <param name="slideAppId">Persistent id for the slide (default generated).</param>
    /// <param name="takeaway">Key message under the table.</param>
    /// <param name="sourceNote">Source line shown on the slide (nothing is shown when omitted).</param>
    /// <param name="totalRow">Style the last row as a total (default false).</param>
    /// <param name="insertAt">1-based position (default after the last slide).</param>
    /// <param name="profile">Design profile (default theme).</param>
    DataOperationResult CreateTable(IPresentationBatch batch, string sourcePath, string title, string? sheet = null, string? range = null, bool header = true,
        string? delimiter = null, string? encoding = null, string culture = "invariant", string? slideAppId = null, string? takeaway = null,
        string? sourceNote = null, bool totalRow = false, int? insertAt = null, string? profile = null);

    /// <summary>
    /// Creates a chart slide from the file, bound to the source: one category column and one or
    /// more numeric value columns (default: every numeric column). Empty cells are missing values;
    /// text in a value column is an error listing the cells.
    /// </summary>
    /// <param name="chartType">column, bar, line, pie, doughnut, area, stacked-column, stacked-bar.</param>
    /// <param name="categoryColumn">Category column name or 1-based number (default: first column).</param>
    /// <param name="valueColumns">Value column names or numbers (default: all numeric columns).</param>
    /// <param name="numberFormat">Excel number format for values and labels, e.g. 0.0 or 0%.</param>
    DataOperationResult CreateChart(IPresentationBatch batch, string sourcePath, string title, string chartType = "column", string? categoryColumn = null,
        IReadOnlyList<string>? valueColumns = null, string? sheet = null, string? range = null, bool header = true, string? delimiter = null,
        string? encoding = null, string culture = "invariant", string? numberFormat = null, string? slideAppId = null, string? takeaway = null,
        string? sourceNote = null, int? insertAt = null, string? profile = null);

    /// <summary>Binds an existing table or chart (by PPTMCP_ID) to a data file so refresh can update it.</summary>
    /// <param name="appId">PPTMCP_ID of the table or chart (deck find or deck assign-ids).</param>
    DataOperationResult Bind(IPresentationBatch batch, string appId, string sourcePath, string? sheet = null, string? range = null, bool header = true,
        string? delimiter = null, string? encoding = null, string culture = "invariant", string? categoryColumn = null, IReadOnlyList<string>? valueColumns = null,
        string? numberFormat = null);

    /// <summary>
    /// Re-reads sources and updates bound tables (changed cell text only; rows added or removed at
    /// the end of single-slide tables) and charts (embedded data rewritten, chart kept). Skips
    /// unchanged sources unless force=true. dry_run reports without writing.
    /// </summary>
    /// <param name="appId">Refresh only this object (or the parts of a split table sharing this slide id).</param>
    /// <param name="dryRun">Report what would change without writing (default false).</param>
    /// <param name="force">Rewrite even if the source file is unchanged (default false).</param>
    DataOperationResult Refresh(IPresentationBatch batch, string? appId = null, bool dryRun = false, bool force = false);

    /// <summary>Lists bound tables and charts with source status (missing, stale).</summary>
    DataOperationResult Bindings(IPresentationBatch batch);
}
