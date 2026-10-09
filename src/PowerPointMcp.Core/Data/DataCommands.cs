extern alias OfficeInterop;

using System.Globalization;
using System.Text.Json.Nodes;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>Data import, binding, and refresh commands.</summary>
public sealed class DataCommands : IDataCommands
{
    private readonly ComposeCommands _compose = new();

    /// <inheritdoc/>
    public DataOperationResult Preview(IPresentationBatch batch, string sourcePath, string? sheet = null, string? range = null, bool header = true,
        string? delimiter = null, string? encoding = null, string culture = "invariant", int maxRows = 20)
    {
        if (maxRows is < 1 or > 200)
            return Fail("max_rows must be between 1 and 200.");
        var binding = new DataBinding { Source = sourcePath, Sheet = sheet, Range = range, Header = header, Delimiter = delimiter, Encoding = encoding, Culture = culture, Kind = "table" };
        if (TryLoad(binding, out var table, out var hash) is { } error)
            return Fail(error);
        return new DataOperationResult
        {
            Success = true,
            Columns = table!.Columns,
            Rows = table.Rows.Take(maxRows).Select(row => (IReadOnlyList<string>)row.Select(value => value.Display).ToList()).ToList(),
            TotalRows = table.Rows.Count,
            Hash = hash,
            Warnings = table.Notes.Count > 0 ? table.Notes : null,
        };
    }

    /// <inheritdoc/>
    public DataOperationResult CreateTable(IPresentationBatch batch, string sourcePath, string title, string? sheet = null, string? range = null, bool header = true,
        string? delimiter = null, string? encoding = null, string culture = "invariant", string? slideAppId = null, string? takeaway = null,
        string? sourceNote = null, bool totalRow = false, int? insertAt = null, string? profile = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var binding = new DataBinding { Source = sourcePath, Sheet = sheet, Range = range, Header = header, Delimiter = delimiter, Encoding = encoding, Culture = culture, Kind = "table" };
        if (TryLoad(binding, out var table, out var hash) is { } error)
            return Fail(error);
        if (table!.Rows.Count == 0)
            return Fail("The data has no rows below the header.");

        var spec = new JsonObject
        {
            ["kind"] = "table",
            ["title"] = title,
            ["table"] = new JsonObject
            {
                ["header"] = new JsonArray(table.Columns.Select(column => (JsonNode)column.Name).ToArray()),
                ["rows"] = new JsonArray(table.Rows.Select(row => (JsonNode)new JsonArray(row.Select(value => (JsonNode)value.Display).ToArray())).ToArray()),
                ["align"] = new JsonArray(table.Columns.Select(column => (JsonNode)(column.Type is "number" or "percent" ? "right" : "left")).ToArray()),
                ["total_row"] = totalRow,
            },
            ["metadata"] = new JsonObject { ["source_file"] = Path.GetFileName(sourcePath), ["source_hash"] = hash },
        };
        if (slideAppId is not null) spec["id"] = slideAppId;
        if (takeaway is not null) spec["takeaway"] = takeaway;
        if (sourceNote is not null) spec["source"] = sourceNote;

        var created = _compose.Create(batch, spec: spec.ToJsonString(), profile: profile, insertAt: insertAt);
        if (!created.Success)
            return Fail(created.ErrorMessage ?? "The table could not be created.", created.Errors);

        // Bind every part of the table; parts of a split table record their row range.
        var refreshedAt = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
        batch.Execute((ctx, ct) =>
        {
            int offset = 0;
            foreach (var slide in created.Slides!)
            {
                var element = slide.Elements.FirstOrDefault(item => item.Key == "table");
                if (element is null)
                    continue;
                WithShape(ctx.Presentation, slide.SlideId, element.ShapeId, shape =>
                {
                    int bodyRows = BodyRowCount(shape);
                    var partBinding = binding with
                    {
                        Hash = hash,
                        RefreshedAt = refreshedAt,
                        RowOffset = offset,
                        RowCount = created.Slides!.Count > 1 ? bodyRows : null,
                    };
                    SetTag(shape, DeckRoles.BindingTag, partBinding.ToJson());
                    offset += bodyRows;
                    return true;
                });
            }
        });

        return new DataOperationResult
        {
            Success = true,
            Columns = table.Columns,
            TotalRows = table.Rows.Count,
            Hash = hash,
            Slides = created.Slides,
            Overflow = created.Overflow,
            Warnings = Merge(table.Notes, created.Warnings),
        };
    }

    /// <inheritdoc/>
    public DataOperationResult CreateChart(IPresentationBatch batch, string sourcePath, string title, string chartType = "column", string? categoryColumn = null,
        IReadOnlyList<string>? valueColumns = null, string? sheet = null, string? range = null, bool header = true, string? delimiter = null,
        string? encoding = null, string culture = "invariant", string? numberFormat = null, string? slideAppId = null, string? takeaway = null,
        string? sourceNote = null, int? insertAt = null, string? profile = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (!ChartDataWriter.ChartTypes.ContainsKey(chartType))
            return Fail($"chart_type must be one of {string.Join(", ", ChartDataWriter.ChartTypes.Keys)}.");
        var binding = new DataBinding
        {
            Source = sourcePath,
            Sheet = sheet,
            Range = range,
            Header = header,
            Delimiter = delimiter,
            Encoding = encoding,
            Culture = culture,
            Kind = "chart",
            CategoryColumn = categoryColumn,
            ValueColumns = valueColumns,
            NumberFormat = numberFormat,
        };
        if (TryLoad(binding, out var table, out var hash) is { } error)
            return Fail(error);
        if (ChartSeries(table!, binding, out var categories, out var series, out var format) is { } seriesError)
            return Fail(seriesError);

        var spec = new JsonObject
        {
            ["kind"] = "chart",
            ["title"] = title,
            ["chart"] = new JsonObject
            {
                ["type"] = chartType,
                ["categories"] = new JsonArray(categories!.Select(category => (JsonNode)category).ToArray()),
                ["series"] = new JsonArray(series!.Select(item => (JsonNode)new JsonObject
                {
                    ["name"] = item.Name,
                    ["values"] = new JsonArray(item.Values.Select(value => double.IsNaN(value) ? null : (JsonNode)value).ToArray()),
                }).ToArray()),
            },
            ["metadata"] = new JsonObject { ["source_file"] = Path.GetFileName(sourcePath), ["source_hash"] = hash },
        };
        if (format is not null) ((JsonObject)spec["chart"]!)["number_format"] = format;
        if (slideAppId is not null) spec["id"] = slideAppId;
        if (takeaway is not null) spec["takeaway"] = takeaway;
        if (sourceNote is not null) spec["source"] = sourceNote;

        var created = _compose.Create(batch, spec: spec.ToJsonString(), profile: profile, insertAt: insertAt);
        if (!created.Success)
            return Fail(created.ErrorMessage ?? "The chart could not be created.", created.Errors);

        var slide = created.Slides![0];
        var chart = slide.Elements.First(item => item.Key == "chart");
        batch.Execute((ctx, ct) => WithShape(ctx.Presentation, slide.SlideId, chart.ShapeId, shape =>
        {
            SetTag(shape, DeckRoles.BindingTag, (binding with { Hash = hash, NumberFormat = format, RefreshedAt = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) }).ToJson());
            return true;
        }));

        return new DataOperationResult
        {
            Success = true,
            Columns = table!.Columns,
            TotalRows = table.Rows.Count,
            Hash = hash,
            Slides = created.Slides,
            Overflow = created.Overflow,
            Warnings = Merge(table.Notes, created.Warnings),
        };
    }

    /// <inheritdoc/>
    public DataOperationResult Bind(IPresentationBatch batch, string appId, string sourcePath, string? sheet = null, string? range = null, bool header = true,
        string? delimiter = null, string? encoding = null, string culture = "invariant", string? categoryColumn = null, IReadOnlyList<string>? valueColumns = null,
        string? numberFormat = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var found = new DeckCommands().Find(batch, $"appid:\"{appId}\"", requireSingle: true);
        if (!found.Success)
            return Fail(found.ErrorMessage ?? $"No object has app id '{appId}'.");
        var target = found.Objects![0];
        if (target.Kind is not ("table" or "chart"))
            return Fail($"'{appId}' is a {target.Kind}; only tables and charts can be bound.");
        var binding = new DataBinding
        {
            Source = sourcePath,
            Sheet = sheet,
            Range = range,
            Header = header,
            Delimiter = delimiter,
            Encoding = encoding,
            Culture = culture,
            Kind = target.Kind,
            CategoryColumn = categoryColumn,
            ValueColumns = valueColumns,
            NumberFormat = numberFormat,
        };
        if (TryLoad(binding, out var table, out var hash) is { } error)
            return Fail(error);
        if (target.Kind == "table" && target.TableColumns != table!.Columns.Count)
            return Fail($"The table has {target.TableColumns} columns but the data has {table.Columns.Count}; bind a matching range or create a new table with data create-table.");
        if (target.Kind == "chart" && ChartSeries(table!, binding, out _, out _, out _) is { } seriesError)
            return Fail(seriesError);

        batch.Execute((ctx, ct) => WithShape(ctx.Presentation, target.SlideId, target.ShapeId, shape =>
        {
            SetTag(shape, DeckRoles.BindingTag, (binding with { Hash = null }).ToJson());
            return true;
        }));
        return new DataOperationResult
        {
            Success = true,
            Columns = table!.Columns,
            Hash = hash,
            Bound = [new BoundObject(appId, target.SlideIndex, target.SlideId, target.ShapeId, target.Kind, sourcePath, true, true, null)],
            Warnings = ["Bound. The object's current content was not changed; run data refresh to load the file into it."],
        };
    }

    /// <inheritdoc/>
    public DataOperationResult Refresh(IPresentationBatch batch, string? appId = null, bool dryRun = false, bool force = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var targets = ListBound(batch).Where(item => appId is null || item.AppId == appId || Continues(item.AppId, appId)).ToList();
        if (targets.Count == 0)
            return Fail(appId is null ? "No bound tables or charts in this presentation." : $"No bound table or chart has app id '{appId}'.");

        var outcomes = new List<RefreshOutcome>();
        var warnings = new List<string>();
        foreach (var group in targets.GroupBy(item => item.Binding.Source + "|" + item.Binding.Sheet + "|" + item.Binding.Range + "|" + item.Binding.Culture))
        {
            var first = group.First().Binding;
            if (TryLoad(first, out var table, out var hash) is { } loadError)
            {
                outcomes.AddRange(group.Select(item => new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, item.Binding.Kind, "failed", 0, 0, 0, [], loadError)));
                continue;
            }
            warnings.AddRange(table!.Notes);
            var parts = group.Where(item => item.Binding.Kind == "table").ToList();
            bool split = parts.Count > 1 || parts.Any(item => item.Binding.RowCount is not null);
            foreach (var item in group)
            {
                if (!force && item.Binding.Hash == hash)
                {
                    outcomes.Add(new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, item.Binding.Kind, "unchanged", 0, 0, 0, [], "The source file has not changed since the last refresh (force=true rewrites anyway)."));
                    continue;
                }
                var outcome = item.Binding.Kind == "chart"
                    ? RefreshChart(batch, item, table, hash!, dryRun)
                    : RefreshTable(batch, item, table, hash!, dryRun, split);
                outcomes.Add(outcome);
            }
        }

        var failed = outcomes.Where(outcome => outcome.Status == "failed").ToList();
        return new DataOperationResult
        {
            Success = failed.Count == 0,
            ErrorMessage = failed.Count == 0 ? null : $"{failed.Count} of {outcomes.Count} bound object(s) could not be refreshed; first: {failed[0].Detail}",
            Refreshed = outcomes,
            Warnings = (dryRun ? warnings.Prepend("Dry run: nothing was written.") : warnings).Distinct().ToList() is { Count: > 0 } list ? list : null,
        };
    }

    /// <inheritdoc/>
    public DataOperationResult Bindings(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);
        var bound = ListBound(batch).Select(item =>
        {
            bool exists = File.Exists(item.Binding.Source);
            bool? stale = exists && item.Binding.Hash is not null ? DataLoader.HashOf(item.Binding.Source) != item.Binding.Hash : null;
            return new BoundObject(item.AppId, item.SlideIndex, item.SlideId, item.ShapeId, item.Binding.Kind, item.Binding.Source, exists, stale, item.Binding.RefreshedAt);
        }).ToList();
        return new DataOperationResult { Success = true, Bound = bound };
    }

    private sealed record BoundItem(string? AppId, int SlideIndex, int SlideId, int ShapeId, DataBinding Binding);

    private static List<BoundItem> ListBound(IPresentationBatch batch) =>
        batch.Execute((ctx, ct) => DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: true, ct)
            .SelectMany(entry => entry.Objects)
            .Select(item => (item, binding: DataBinding.Parse(item.Tags?.GetValueOrDefault(DeckRoles.BindingTag))))
            .Where(entry => entry.binding is not null)
            .Select(entry => new BoundItem(entry.item.AppId, entry.item.SlideIndex, entry.item.SlideId, entry.item.ShapeId, entry.binding!))
            .ToList());

    private static bool Continues(string? candidate, string appId)
    {
        // Split tables: "sales/table" continues as "sales~2/table".
        if (candidate is null)
            return false;
        var slash = appId.IndexOf('/', StringComparison.Ordinal);
        if (slash < 0)
            return false;
        var (slidePart, key) = (appId[..slash], appId[slash..]);
        return candidate.StartsWith(slidePart + "~", StringComparison.Ordinal) && candidate.EndsWith(key, StringComparison.Ordinal);
    }

    private static RefreshOutcome RefreshTable(IPresentationBatch batch, BoundItem item, DataTableContent table, string hash, bool dryRun, bool split)
    {
        var start = item.Binding.RowOffset;
        var count = item.Binding.RowCount ?? Math.Max(0, table.Rows.Count - start);
        var rows = table.Rows.Skip(start).Take(count).ToList();
        if (split && item.Binding.RowCount is not null && table.Rows.Count != PartsTotal(batch, item))
        {
            return new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "table", "failed", 0, 0, 0, [],
                $"The data now has {table.Rows.Count} rows, but this table is split across slides for a different row count. Re-create it with data create-table (slide_app_id with replace via compose) to repaginate.");
        }

        return batch.Execute((ctx, ct) =>
        {
            RefreshOutcome? outcome = null;
            WithShape(ctx.Presentation, item.SlideId, item.ShapeId, shape =>
            {
                PowerPoint.Table? native = null;
                PowerPoint.Rows? nativeRows = null;
                PowerPoint.Columns? nativeColumns = null;
                try
                {
                    native = shape.Table;
                    nativeRows = native.Rows;
                    nativeColumns = native.Columns;
                    int columns = nativeColumns.Count;
                    if (columns != table.Columns.Count)
                    {
                        outcome = new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "table", "failed", 0, 0, 0, [],
                            $"The table has {columns} columns but the data now has {table.Columns.Count}; columns are never added or dropped silently. Re-create the table.");
                        return false;
                    }
                    int headerRows = item.Binding.Header ? 1 : 0;
                    int existing = nativeRows.Count - headerRows;
                    int added = 0, removed = 0, changed = 0;
                    var samples = new List<string>();

                    if (!dryRun)
                    {
                        while (nativeRows.Count - headerRows < rows.Count)
                        {
                            // Rows.Add appends a row that copies the last row's formatting.
                            var appended = nativeRows.Add();
                            ComUtilities.Release(ref appended);
                            added++;
                        }
                        while (nativeRows.Count - headerRows > rows.Count && nativeRows.Count > headerRows + 1)
                        {
                            PowerPoint.Row? last = null;
                            try
                            {
                                last = nativeRows[nativeRows.Count];
                                last.Delete();
                                removed++;
                            }
                            finally
                            {
                                if (last is not null) ComUtilities.Release(ref last);
                            }
                        }
                    }
                    else
                    {
                        added = Math.Max(0, rows.Count - existing);
                        removed = Math.Max(0, existing - rows.Count);
                    }

                    if (item.Binding.Header)
                        changed += WriteRow(native, 1, table.Columns.Select(column => column.Name).ToList(), dryRun, samples);
                    for (int r = 0; r < rows.Count && r + headerRows + 1 <= (dryRun ? existing + headerRows : nativeRows.Count); r++)
                        changed += WriteRow(native, r + headerRows + 1, rows[r].Select(value => value.Display).ToList(), dryRun, samples);

                    if (!dryRun)
                        SetTag(shape, DeckRoles.BindingTag, (item.Binding with { Hash = hash, RefreshedAt = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) }).ToJson());
                    outcome = new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "table", dryRun ? "would-change" : "refreshed", changed, added, removed, samples,
                        added > 0 ? "New rows copy the formatting of the last row; check banding and the total row with a preview." : null);
                    return true;
                }
                finally
                {
                    if (nativeColumns is not null) ComUtilities.Release(ref nativeColumns);
                    if (nativeRows is not null) ComUtilities.Release(ref nativeRows);
                    if (native is not null) ComUtilities.Release(ref native);
                }
            });
            return outcome ?? new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "table", "failed", 0, 0, 0, [], "The table shape no longer exists.");
        });
    }

    private static int PartsTotal(IPresentationBatch batch, BoundItem item) =>
        ListBound(batch).Where(other => other.Binding.Source == item.Binding.Source && other.Binding.Kind == "table" &&
            other.Binding.RowCount is not null && SameFamily(other.AppId, item.AppId))
            .Sum(other => other.Binding.RowCount!.Value);

    private static bool SameFamily(string? a, string? b)
    {
        static string Root(string? id) => id is null ? "" : id.Split('/')[0].Split('~')[0];
        return a is not null && b is not null && Root(a) == Root(b);
    }

    private static int WriteRow(PowerPoint.Table table, int row, List<string> values, bool dryRun, List<string> samples)
    {
        int changed = 0;
        for (int column = 1; column <= values.Count; column++)
        {
            var current = DeckSnapshotReader.ReadCellText(table, row, column);
            var next = values[column - 1];
            if (current == next.Replace('\r', ' ').Replace('\v', ' '))
                continue;
            changed++;
            if (samples.Count < 10)
                samples.Add($"R{row}C{column}: \"{current}\" → \"{next}\"");
            if (dryRun)
                continue;
            PowerPoint.Cell? cell = null;
            PowerPoint.Shape? shape = null;
            try
            {
                cell = table.Cell(row, column);
                shape = cell.Shape;
                // Assigning Text keeps the first run's formatting (font, size, color, alignment).
                TextRuns.WithRange(shape, (_, range) => { range.Text = next.Replace("\r\n", "\v", StringComparison.Ordinal).Replace('\n', '\v'); return true; });
            }
            finally
            {
                if (shape is not null) ComUtilities.Release(ref shape);
                if (cell is not null) ComUtilities.Release(ref cell);
            }
        }
        return changed;
    }

    private static RefreshOutcome RefreshChart(IPresentationBatch batch, BoundItem item, DataTableContent table, string hash, bool dryRun)
    {
        if (ChartSeries(table, item.Binding, out var categories, out var series, out var format) is { } error)
            return new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "chart", "failed", 0, 0, 0, [], error);
        if (dryRun)
        {
            return new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "chart", "would-change", categories!.Count * series!.Count, 0, 0, [],
                $"Would write {categories.Count} categories x {series.Count} series ({string.Join(", ", series.Select(s => s.Name))}).");
        }
        return batch.Execute((ctx, ct) =>
        {
            RefreshOutcome? outcome = null;
            WithShape(ctx.Presentation, item.SlideId, item.ShapeId, shape =>
            {
                if (shape.HasChart != Office.MsoTriState.msoTrue)
                    return false;
                PowerPoint.Chart? chart = null;
                try
                {
                    chart = shape.Chart;
                    var reference = ChartDataWriter.Write(chart, categories!, series!, format ?? item.Binding.NumberFormat);
                    SetTag(shape, DeckRoles.BindingTag, (item.Binding with { Hash = hash, RefreshedAt = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) }).ToJson());
                    outcome = new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "chart", "refreshed", categories!.Count * series!.Count, 0, 0, [],
                        $"Wrote {categories.Count} categories x {series.Count} series to the embedded data ({reference}); chart formatting kept, new series take the chart's default colors.");
                    return true;
                }
                finally
                {
                    if (chart is not null) ComUtilities.Release(ref chart);
                }
            });
            return outcome ?? new RefreshOutcome(item.AppId, item.SlideIndex, item.ShapeId, "chart", "failed", 0, 0, 0, [], "The chart shape no longer exists or is not a chart.");
        });
    }

    /// <summary>Builds chart categories and numeric series; text in a value column is an error.</summary>
    private static string? ChartSeries(DataTableContent table, DataBinding binding, out List<string>? categories, out List<ChartSeriesData>? series, out string? format)
    {
        categories = null;
        series = null;
        format = binding.NumberFormat;
        int categoryIndex = binding.CategoryColumn is null ? 0 : table.ColumnIndex(binding.CategoryColumn);
        if (categoryIndex < 0)
            return $"category_column '{binding.CategoryColumn}' not found. Columns: {string.Join(", ", table.Columns.Select(column => column.Name))}.";
        var valueIndexes = binding.ValueColumns is { Count: > 0 } named
            ? named.Select(table.ColumnIndex).ToList()
            : table.Columns.Where(column => column.Index != categoryIndex && column.Type is "number" or "percent").Select(column => column.Index).ToList();
        if (valueIndexes.Any(index => index < 0))
            return $"value_columns contains unknown columns. Columns: {string.Join(", ", table.Columns.Select(column => column.Name))}.";
        if (valueIndexes.Count == 0)
            return "No numeric value columns found; pass value_columns, and check data preview for columns typed as text.";
        var problems = new List<string>();
        foreach (var index in valueIndexes)
        {
            var bad = table.Rows.Select((row, r) => (row[index], r)).Where(entry => entry.Item1.Kind is not ("number" or "percent" or "empty")).Take(5).ToList();
            if (bad.Count > 0)
                problems.Add($"column '{table.Columns[index].Name}' has non-numeric cells: {string.Join(", ", bad.Select(entry => $"row {entry.r + 1} \"{entry.Item1.Raw}\""))}");
        }
        if (problems.Count > 0)
            return "Chart values must be numbers; nothing is converted. " + string.Join("; ", problems) + ".";
        categories = table.Rows.Select(row => row[categoryIndex].Display).ToList();
        series = valueIndexes.Select(index => new ChartSeriesData(table.Columns[index].Name,
            table.Rows.Select(row => row[index].Number ?? double.NaN).ToList())).ToList();
        if (format is null && valueIndexes.All(index => table.Columns[index].Type == "percent"))
            format = "0%";
        return null;
    }

    private static string? TryLoad(DataBinding binding, out DataTableContent? table, out string? hash)
    {
        table = null;
        hash = null;
        if (string.IsNullOrWhiteSpace(binding.Source) || !Path.IsPathFullyQualified(binding.Source))
            return "source_path must be a full local path (e.g. C:\\Data\\sales.xlsx).";
        try
        {
            (table, hash) = DataLoader.Load(binding);
            return null;
        }
        catch (FileNotFoundException ex)
        {
            return ex.Message;
        }
        catch (ArgumentException ex)
        {
            return ex.Message;
        }
        catch (InvalidDataException ex)
        {
            return $"'{binding.Source}' could not be read: {ex.Message}";
        }
        catch (IOException ex)
        {
            return $"'{binding.Source}' could not be read (is it open in another program?): {ex.Message}";
        }
        catch (System.Xml.XmlException ex)
        {
            return $"'{binding.Source}' is not a valid .xlsx file: {ex.Message}";
        }
    }

    private static int BodyRowCount(PowerPoint.Shape shape)
    {
        PowerPoint.Table? table = null;
        PowerPoint.Rows? rows = null;
        try
        {
            table = shape.Table;
            rows = table.Rows;
            return rows.Count - 1;
        }
        finally
        {
            if (rows is not null) ComUtilities.Release(ref rows);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static bool WithShape(PowerPoint.Presentation presentation, int slideId, int shapeId, Func<PowerPoint.Shape, bool> action)
    {
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        try
        {
            slide = DeckShapeLocator.FindSlide(presentation, slideId);
            shape = slide is null ? null : DeckShapeLocator.FindShape(slide, shapeId);
            return shape is not null && action(shape);
        }
        finally
        {
            if (shape is not null) ComUtilities.Release(ref shape);
            if (slide is not null) ComUtilities.Release(ref slide);
        }
    }

    private static void SetTag(PowerPoint.Shape shape, string name, string value)
    {
        var tags = shape.Tags;
        try
        {
            tags.Add(name, value);
        }
        finally
        {
            ComUtilities.Release(ref tags);
        }
    }

    private static List<string>? Merge(IReadOnlyList<string> first, IReadOnlyList<string>? second)
    {
        var list = first.Concat(second ?? []).Distinct().ToList();
        return list.Count > 0 ? list : null;
    }

    private static DataOperationResult Fail(string message, IReadOnlyList<string>? details = null) =>
        new() { Success = false, ErrorMessage = message, Warnings = details is { Count: > 0 } ? details : null };
}
