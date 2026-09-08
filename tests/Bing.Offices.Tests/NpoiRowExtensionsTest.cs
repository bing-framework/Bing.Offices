using System;
using Bing.Offices;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// RowExtensions 的职责级测试。
/// </summary>
public sealed class NpoiRowExtensionsTest
{
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void GetOrCreateCell_ShouldCreateMissingCellAndReuseExistingCell(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var row = workbook.CreateSheet("Data").CreateRow(0);

        var created = RowExtensions.GetOrCreateCell(row, 2);
        var reused = RowExtensions.GetOrCreateCell(row, 2);

        Assert.Same(created, reused);
        Assert.Equal(2, created.ColumnIndex);
        Assert.Same(created, row.GetCell(2));
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void CreateCell_ShouldRunActionAndReturnSameRow(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var row = workbook.CreateSheet("Data").CreateRow(0);
        ICell actionCell = null;

        var returned = RowExtensions.CreateCell(row, 1, cell =>
        {
            actionCell = cell;
            cell.SetCellValue("configured");
        });

        Assert.Same(row, returned);
        Assert.Same(row.GetCell(1), actionCell);
        Assert.Equal("configured", row.GetCell(1).StringCellValue);
        Assert.Same(row, RowExtensions.CreateCell(row, 1, null));
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ClearContent_ShouldReplaceEveryExistingCellWithEmptyString(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var row = workbook.CreateSheet("Data").CreateRow(0);
        row.CreateCell(0).SetCellValue("text");
        row.CreateCell(1).SetCellValue(12.5d);
        row.CreateCell(2).SetCellValue(true);

        var returned = RowExtensions.ClearContent(row);

        Assert.Same(row, returned);
        Assert.All(row.Cells, cell =>
        {
            Assert.Equal(CellType.String, cell.CellType);
            Assert.Equal(string.Empty, cell.StringCellValue);
        });
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void IsEmptyRow_ShouldTreatNullBlankAndWhitespaceRowsAsEmpty(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var sheet = workbook.CreateSheet("Data");
        var blankRow = sheet.CreateRow(0);
        blankRow.CreateCell(0, CellType.Blank);
        blankRow.CreateCell(1).SetCellValue("  \t");
        var valueRow = sheet.CreateRow(1);
        valueRow.CreateCell(0).SetCellValue("value");

        Assert.True(RowExtensions.IsEmptyRow(null));
        Assert.True(RowExtensions.IsEmptyRow(sheet.CreateRow(2)));
        Assert.True(RowExtensions.IsEmptyRow(blankRow));
        Assert.False(RowExtensions.IsEmptyRow(valueRow));
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Value_ShouldWriteValueAndBindProvidedStyle(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var row = workbook.CreateSheet("Data").CreateRow(0);
        var style = workbook.CreateCellStyle();
        style.Alignment = HorizontalAlignment.Center;

        RowExtensions.Value(row, 0, "value", style);
        RowExtensions.Value<string>(row, 1, null);

        var valueCell = row.GetCell(0);
        Assert.Equal(CellType.String, valueCell.CellType);
        Assert.Equal("value", valueCell.StringCellValue);
        Assert.Equal(style.Index, valueCell.CellStyle.Index);
        Assert.Equal(HorizontalAlignment.Center, valueCell.CellStyle.Alignment);
        Assert.Equal(CellType.Blank, row.GetCell(1).CellType);
    }

    private static IWorkbook CreateWorkbook(ExcelFormat format) =>
        format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
}
