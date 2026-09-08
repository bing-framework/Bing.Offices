using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// WorkbookExtensions 的职责级测试。
/// </summary>
public sealed class NpoiWorkbookExtensionsTest
{
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void GetExcelFormat_ShouldIdentifyWorkbookFormat(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);

        var actual = WorkbookExtensions.GetExcelFormat(workbook);

        Assert.Equal(format, actual);
    }

    [Fact]
    public void GetExcelFormat_UnsupportedOrNull_ShouldThrowNotSupportedException()
    {
        var exception = Assert.Throws<NotSupportedException>(() =>
            WorkbookExtensions.GetExcelFormat(null));

        Assert.Equal("未知 Excel 格式类型。", exception.Message);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void GetSheets_ShouldKeepVisibleSheetsInWorkbookOrder(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        workbook.CreateSheet("Visible");
        var hidden = workbook.CreateSheet("Hidden");
        var veryHidden = workbook.CreateSheet("VeryHidden");
        workbook.SetSheetVisibility(workbook.GetSheetIndex(hidden), SheetVisibility.Hidden);
        workbook.SetSheetVisibility(workbook.GetSheetIndex(veryHidden), SheetVisibility.VeryHidden);

        var actual = WorkbookExtensions.GetSheets(workbook).Select(sheet => sheet.SheetName).ToArray();

        Assert.Equal(new[] { "Visible" }, actual);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void GetSheets_EmptyWorkbook_ShouldReturnEmptyCollection(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);

        var actual = WorkbookExtensions.GetSheets(workbook);

        Assert.Empty(actual);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetAllSheetAutoCompute_ShouldSetEverySheetAndAllowEmptyWorkbook(ExcelFormat format)
    {
        using var emptyWorkbook = CreateWorkbook(format);
        WorkbookExtensions.SetAllSheetAutoCompute(emptyWorkbook);
        Assert.Empty(WorkbookExtensions.GetSheets(emptyWorkbook));

        using var workbook = CreateWorkbook(format);
        var first = workbook.CreateSheet("First");
        var second = workbook.CreateSheet("Second");
        first.ForceFormulaRecalculation = false;
        second.ForceFormulaRecalculation = false;

        WorkbookExtensions.SetAllSheetAutoCompute(workbook);

        Assert.True(first.ForceFormulaRecalculation);
        Assert.True(second.ForceFormulaRecalculation);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void AddSheet_ShouldCreateHeaderRowWithDefaultHeaderStyle(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var heads = new List<string> { "Code", "Name" };

        var sheet = WorkbookExtensions.AddSheet(workbook, "Data", heads);

        Assert.Equal("Data", sheet.SheetName);
        Assert.Equal(1, workbook.NumberOfSheets);
        var row = sheet.GetRow(0);
        Assert.NotNull(row);
        Assert.Equal((short)(20 * 20), row.Height);
        for (var index = 0; index < heads.Count; index++)
        {
            var cell = row.GetCell(index);
            Assert.NotNull(cell);
            Assert.Equal(CellType.String, cell.CellType);
            Assert.Equal(heads[index], cell.StringCellValue);
            AssertHeaderStyle(workbook, cell.CellStyle);
        }
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void AddSheet_NullHeads_ShouldThrowArgumentNullException(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);

        var exception = Assert.Throws<ArgumentNullException>(() =>
            WorkbookExtensions.AddSheet(workbook, "Data", null));

        Assert.Equal("heads", exception.ParamName);
        Assert.Equal(0, workbook.NumberOfSheets);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void DefaultHeadStyle_ShouldUseExpectedDefaultProperties(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);

        var style = WorkbookExtensions.DefaultHeadStyle(workbook);

        Assert.Equal((short)13, style.FillForegroundColor);
        Assert.Equal(FillPattern.SolidForeground, style.FillPattern);
        Assert.Equal(BorderStyle.Thin, style.BorderTop);
        Assert.Equal(BorderStyle.Thin, style.BorderRight);
        Assert.Equal(BorderStyle.Thin, style.BorderBottom);
        Assert.Equal(BorderStyle.Thin, style.BorderLeft);
        Assert.Equal(HorizontalAlignment.Center, style.Alignment);
        Assert.Equal(VerticalAlignment.Center, style.VerticalAlignment);
        var font = workbook.GetFontAt(style.FontIndex);
        Assert.Equal("宋体", font.FontName);
        Assert.Equal((short)9, font.FontHeightInPoints);
        Assert.True(font.IsBold);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void DefaultBodyStyle_ShouldUseExpectedDefaultProperties(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);

        var style = WorkbookExtensions.DefaultBodyStyle(workbook);

        Assert.Equal(FillPattern.NoFill, style.FillPattern);
        Assert.Equal(BorderStyle.Thin, style.BorderTop);
        Assert.Equal(BorderStyle.Thin, style.BorderRight);
        Assert.Equal(BorderStyle.Thin, style.BorderBottom);
        Assert.Equal(BorderStyle.Thin, style.BorderLeft);
        Assert.Equal(HorizontalAlignment.Center, style.Alignment);
        Assert.Equal(VerticalAlignment.Center, style.VerticalAlignment);
        var font = workbook.GetFontAt(style.FontIndex);
        Assert.Equal("宋体", font.FontName);
        Assert.Equal((short)9, font.FontHeightInPoints);
        Assert.False(font.IsBold);
    }

    private static IWorkbook CreateWorkbook(ExcelFormat format) =>
        format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();

    private static void AssertHeaderStyle(IWorkbook workbook, ICellStyle style)
    {
        Assert.Equal((short)13, style.FillForegroundColor);
        Assert.Equal(FillPattern.SolidForeground, style.FillPattern);
        Assert.Equal(BorderStyle.Thin, style.BorderTop);
        Assert.Equal(BorderStyle.Thin, style.BorderRight);
        Assert.Equal(BorderStyle.Thin, style.BorderBottom);
        Assert.Equal(BorderStyle.Thin, style.BorderLeft);
        Assert.Equal(HorizontalAlignment.Center, style.Alignment);
        Assert.Equal(VerticalAlignment.Center, style.VerticalAlignment);
        var font = workbook.GetFontAt(style.FontIndex);
        Assert.True(font.IsBold);
        Assert.Equal("宋体", font.FontName);
        Assert.Equal((short)9, font.FontHeightInPoints);
    }
}
