using System.Text;
using Sbroenne.PowerPointMcp.Core.Chart;
using Sbroenne.PowerPointMcp.Core.Data;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.PageSetup;
using Sbroenne.PowerPointMcp.Core.Table;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for data import, binding, refresh, and the table/chart extensions.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Data")]
public class DataCommandsTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly DataCommands _data = new();
    private readonly DeckCommands _deck = new();

    public DataCommandsTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    private void FreshWideDeck()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new PageSetupCommands().SetSize(_fixture.Batch, 960, 540).Success);
    }

    private static string WriteCsv(string content, Encoding? encoding = null)
    {
        var path = Path.Combine(Path.GetTempPath(), "PowerPointMcpTests", $"данные-{Guid.NewGuid():N}.csv");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content, encoding ?? new UTF8Encoding(true));
        return path;
    }

    [Fact]
    public void CreateTable_ThenRefresh_UpdatesChangedCellsKeepingFormatting()
    {
        FreshWideDeck();
        var csv = WriteCsv("Регион;Выручка;Доля\nМосква;1 234,5;12%\nКазань;980,0;8%\n");

        var created = _data.CreateTable(_fixture.Batch, csv, "Выручка по регионам", culture: "ru-RU", slideAppId: "rev");
        Assert.True(created.Success, created.ErrorMessage);
        Assert.Equal(["text", "number", "percent"], created.Columns!.Select(column => column.Type));
        var before = _deck.Find(_fixture.Batch, "appid:rev/table", requireSingle: true).Objects![0];
        Assert.Contains("1 234,5", before.Text, StringComparison.Ordinal);

        File.WriteAllText(csv, "Регион;Выручка;Доля\nМосква;1 300,0;13%\nКазань;980,0;8%\nСочи;120,0;1%\n", new UTF8Encoding(true));
        var dry = _data.Refresh(_fixture.Batch, appId: "rev/table", dryRun: true);
        Assert.True(dry.Success, dry.ErrorMessage);
        Assert.Equal("would-change", dry.Refreshed![0].Status);
        Assert.Contains("1 234,5", _deck.Find(_fixture.Batch, "appid:rev/table", requireSingle: true).Objects![0].Text, StringComparison.Ordinal);

        var refreshed = _data.Refresh(_fixture.Batch, appId: "rev/table");
        Assert.True(refreshed.Success, refreshed.ErrorMessage);
        var outcome = Assert.Single(refreshed.Refreshed!);
        Assert.Equal(("refreshed", 2, 1), (outcome.Status, outcome.CellsChanged, outcome.RowsAdded));
        var after = _deck.Find(_fixture.Batch, "appid:rev/table", requireSingle: true).Objects![0];
        Assert.Contains("1 300,0", after.Text, StringComparison.Ordinal);
        Assert.Contains("Сочи", after.Text, StringComparison.Ordinal);
        Assert.Equal(before.MinFontSize, after.MinFontSize);

        var again = _data.Refresh(_fixture.Batch, appId: "rev/table");
        Assert.Equal("unchanged", again.Refreshed![0].Status);
    }

    [Fact]
    public void CreateChart_FromCsv_WithMissingValues_AndRefresh()
    {
        FreshWideDeck();
        var csv = WriteCsv("Quarter,2024,2025\nQ1,1.1,1.3\nQ2,1.3,1.5\nQ3,1.2,\nQ4,1.5,\n");

        var created = _data.CreateChart(_fixture.Batch, csv, "Revenue", chartType: "line", slideAppId: "trend");
        Assert.True(created.Success, created.ErrorMessage);
        var chart = _deck.Find(_fixture.Batch, "appid:trend/chart", requireSingle: true).Objects![0];

        var details = new ChartCommands().GetDetails(_fixture.Batch, chart.SlideIndex, chart.ShapeIndex!.Value);
        Assert.True(details.Success, details.ErrorMessage);
        Assert.Equal(["Q1", "Q2", "Q3", "Q4"], details.Details!.Categories);
        Assert.Equal(["2024", "2025"], details.Details.Series.Select(series => series.Name));
        Assert.Null(details.Details.Series[1].Values[2]);

        File.WriteAllText(csv, "Quarter,2024,2025\nQ1,1.1,1.3\nQ2,1.3,1.5\nQ3,1.2,1.6\nQ4,1.5,1.8\n");
        var refreshed = _data.Refresh(_fixture.Batch, appId: "trend/chart");
        Assert.True(refreshed.Success, refreshed.ErrorMessage);
        var updated = new ChartCommands().GetDetails(_fixture.Batch, chart.SlideIndex, chart.ShapeIndex!.Value);
        Assert.Equal(1.8, updated.Details!.Series[1].Values[3]);
    }

    [Fact]
    public void CreateChart_RejectsTextInValueColumns()
    {
        FreshWideDeck();
        var csv = WriteCsv("Quarter,Revenue\nQ1,1.1\nQ2,n/a\n");
        var result = _data.CreateChart(_fixture.Batch, csv, "Revenue", valueColumns: ["Revenue"]);
        Assert.False(result.Success);
        Assert.Contains("row 2 \"n/a\"", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void Bindings_ReportStaleAndMissingSources()
    {
        FreshWideDeck();
        var csv = WriteCsv("A,B\n1,2\n");
        Assert.True(_data.CreateTable(_fixture.Batch, csv, "T", slideAppId: "t").Success);
        File.AppendAllText(csv, "3,4\n");
        var stale = Assert.Single(_data.Bindings(_fixture.Batch).Bound!);
        Assert.True(stale.Stale);
        File.Delete(csv);
        var missing = Assert.Single(_data.Bindings(_fixture.Batch).Bound!);
        Assert.False(missing.SourceExists);
        Assert.False(_data.Refresh(_fixture.Batch).Success);
    }

    [Fact]
    public void TableExtensions_FormatNumbersConditionalFormatAndStyle()
    {
        FreshWideDeck();
        var tables = new TableCommands();
        Assert.True(tables.AddTable(_fixture.Batch, 1, 4, 2, 50, 100, 400, 120).Success);
        int index = new Shape.ShapeCommands().GetCount(_fixture.Batch, 1).ShapeCount!.Value;
        Assert.True(tables.SetData(_fixture.Batch, 1, index, ["Item|Delta", "A|1234.5", "B|-12", "C|n/a"]).Success);

        var formatted = tables.FormatNumbers(_fixture.Batch, 1, index, 2, "#,##0.0;(#,##0.0)", culture: "en-US");
        Assert.True(formatted.Success, formatted.ErrorMessage);
        Assert.Equal(2, formatted.CellsChanged);
        Assert.Single(formatted.Skipped!);
        Assert.Equal("1,234.5", tables.GetCellText(_fixture.Batch, 1, index, 2, 2).CellText);
        Assert.Equal("(12.0)", tables.GetCellText(_fixture.Batch, 1, index, 3, 2).CellText);

        var negative = tables.ConditionalFormat(_fixture.Batch, 1, index, 2, "negative", textColor: "#C62828", culture: "en-US");
        Assert.True(negative.Success, negative.ErrorMessage);
        Assert.Equal(1, negative.CellsChanged);

        Assert.True(tables.ApplyStyle(_fixture.Batch, 1, index, profile: "corporate-blue").Success);
        Assert.Equal("1,234.5", tables.GetCellText(_fixture.Batch, 1, index, 2, 2).CellText);
        Assert.True(tables.SetColumnWidths(_fixture.Batch, 1, index, [3, 1], mode: "weights").Success);
    }

    [Fact]
    public void ChartExtensions_ComboSecondaryAxisAndAxisBounds()
    {
        FreshWideDeck();
        var charts = new ChartCommands();
        var added = charts.AddChart(_fixture.Batch, 1, "bar", 50, 100, 500, 300, ["Q1", "Q2", "Q3"], "Revenue", [10, 20, 30]);
        Assert.True(added.Success, added.ErrorMessage);
        int index = added.ShapeIndex!.Value;
        Assert.True(charts.SetData(_fixture.Batch, 1, index, ["Q1", "Q2", "Q3"], ["Revenue", "Margin"], ["10", "20", "30", "0.1", "", "0.3"], "0.0").Success);
        Assert.False(charts.SetData(_fixture.Batch, 1, index, ["Q1"], ["Revenue"], ["ten"]).Success);

        var combo = charts.SetSeriesStyle(_fixture.Batch, 1, index, 2, color: "#C55A11", seriesType: "line", secondaryAxis: true);
        Assert.True(combo.Success, combo.ErrorMessage);
        Assert.True(charts.SetAxis(_fixture.Batch, 1, index, "secondary-value", minimum: 0, maximum: 0.5, numberFormat: "0%").Success);
        Assert.True(charts.SetMissingValues(_fixture.Batch, 1, index, "connect").Success);
        Assert.True(charts.SetDataLabels(_fixture.Batch, 1, index, true, seriesIndex: 1, numberFormat: "0").Success);

        var details = charts.GetDetails(_fixture.Batch, 1, index);
        Assert.True(details.Details!.Series[1].SecondaryAxis);
        Assert.Equal("connect", details.Details.MissingValues);
        Assert.Equal("#C55A11", details.Details.Series[1].Color);
    }
}
