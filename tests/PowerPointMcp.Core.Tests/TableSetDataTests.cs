using Sbroenne.PowerPointMcp.Core.Table;

namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>Real integration tests for table set-data against live PowerPoint COM.</summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Table")]
public class TableSetDataTests : IClassFixture<SharedPresentationFixture>
{
    private readonly SharedPresentationFixture _fixture;
    private readonly TableCommands _tables = new();

    public TableSetDataTests(SharedPresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetData_FillsEveryCellFromMarkdownRows()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_tables.AddTable(_fixture.Batch, 1, 3, 3, 50, 100, 500, 150).Success);

        var result = _tables.SetData(_fixture.Batch, 1, 1,
        [
            "| Region | Revenue | Growth |",
            "|---|---:|---:|",
            "| APAC | $2.4M | +24% |",
            "| EMEA | $3.1M | +8% |",
        ]);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(9, result.CellsWritten);
        Assert.Equal(("Region", "+24%", "$3.1M"), (Cell(1, 1), Cell(2, 3), Cell(3, 2)));
    }

    [Fact]
    public void SetData_WithStartRow_LeavesTheHeaderUntouched()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_tables.AddTable(_fixture.Batch, 1, 3, 2, 50, 100, 400, 150).Success);
        Assert.True(_tables.SetCellText(_fixture.Batch, 1, 1, 1, 1, "Header").Success);

        var result = _tables.SetData(_fixture.Batch, 1, 1, ["a;b", "c;d"], separator: ";", startRow: 2);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(4, result.CellsWritten);
        Assert.Equal(("Header", "a", "d"), (Cell(1, 1), Cell(2, 1), Cell(3, 2)));
    }

    [Theory]
    [InlineData(new[] { "a|b", "c|d", "e|f" }, "row")]
    [InlineData(new[] { "a|b|c" }, "column")]
    public void SetData_TooLargeForTheTable_IsRejectedWithoutWriting(string[] data, string dimension)
    {
        _fixture.CreateFreshPresentation();
        Assert.True(_tables.AddTable(_fixture.Batch, 1, 2, 2, 50, 100, 400, 100).Success);
        Assert.True(_tables.SetCellText(_fixture.Batch, 1, 1, 1, 1, "keep").Success);

        var result = _tables.SetData(_fixture.Batch, 1, 1, data);

        Assert.False(result.Success);
        Assert.Contains(dimension, result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("Nothing was written", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal("keep", Cell(1, 1));
    }

    [Fact]
    public void SetData_OnAShapeWithoutATable_Fails()
    {
        _fixture.CreateFreshPresentation();
        Assert.True(new Shape.ShapeCommands().AddRectangle(_fixture.Batch, 1, 10, 10, 50, 50).Success);

        var result = _tables.SetData(_fixture.Batch, 1, 1, ["a|b"]);

        Assert.False(result.Success);
        Assert.Contains("does not contain a table", result.ErrorMessage, StringComparison.Ordinal);
    }

    private string Cell(int row, int column)
    {
        var result = _tables.GetCellText(_fixture.Batch, 1, 1, row, column);
        Assert.True(result.Success, result.ErrorMessage);
        return result.CellText!;
    }
}
