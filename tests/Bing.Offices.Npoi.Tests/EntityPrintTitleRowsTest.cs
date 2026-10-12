using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证实体布局的重复打印标题行。
/// </summary>
public sealed class EntityPrintTitleRowsTest
{
    /// <summary>
    /// 验证保存重开后重复标题行生效且不插入物理行。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PrintTitleRows_Export_ShouldPersistWithoutInsertingRows(bool xls)
    {
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(), CreateLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.GetSheet("明细");
        var printTitle = Assert.Single(workbook.GetAllNames().Where(name =>
            name.NameName.EndsWith("Print_Titles", StringComparison.Ordinal)));
        Assert.Contains("$1", printTitle.RefersToFormula, StringComparison.Ordinal);
        Assert.Equal(0, sheet.RepeatingRows.FirstRow);
        Assert.Equal(1, sheet.RepeatingRows.LastRow);
        Assert.Equal("采购单", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Code", sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.Equal("B", sheet.GetRow(4).GetCell(0).StringCellValue);
        Assert.Null(sheet.GetRow(5));
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证模板既有打印设置和打印区域在新增标题行后保持不变。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PrintTitleRows_Template_ShouldPreserveUnrelatedPrintSettings(bool xls)
    {
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(), CreateLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.GetSheet("明细");
        Assert.Equal((short)PaperSize.US_Letter_Small, sheet.PrintSetup.PaperSize);
        Assert.Equal((short)73, sheet.PrintSetup.Scale);
        Assert.Contains("$A$1:$B$8", workbook.GetPrintArea(workbook.GetSheetIndex(sheet)));
    }

    /// <summary>
    /// 验证重复配置及与实体物理页眉并用均在布局构建时失败。
    /// </summary>
    [Fact]
    public void PrintTitleRows_DuplicateOrPhysicalHeader_ShouldRejectLayout()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<PrintDocument>(builder => builder
            .PrintTitleRows("明细", "1:2")
            .PrintTitleRows("明细", "1:3")));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<PrintDocument>(builder => builder
            .PrintTitleRows("明细", "1:2")
            .ListRegion("明细", "A3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:B3", "B1", "HEADER"))));
    }

    private static ExcelEntityLayout<PrintDocument> CreateLayout() =>
        ExcelEntity.Layout<PrintDocument>(builder => builder
            .PrintTitleRows("明细", "1:2")
            .ListRegion("明细", "A3", item => item.Lines));

    private static PrintDocument CreateDocument() => new()
    {
        Lines = new List<PrintLine> { new() { Code = "A" }, new() { Code = "B" } }
    };

    private static MemoryStream CreateTemplate(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("采购单");
        sheet.PrintSetup.PaperSize = (short)PaperSize.US_Letter_Small;
        sheet.PrintSetup.Scale = 73;
        workbook.SetPrintArea(workbook.GetSheetIndex(sheet), "$A$1:$B$8");
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 打印布局测试单据。
    /// </summary>
    public sealed class PrintDocument
    {
        /// <summary>
        /// 单据明细。
        /// </summary>
        public List<PrintLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 打印布局测试明细。
    /// </summary>
    public sealed class PrintLine
    {
        /// <summary>
        /// 明细编码。
        /// </summary>
        public string Code { get; set; }
    }
}
