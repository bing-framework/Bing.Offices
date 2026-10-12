using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证实体布局的重复打印标题行。
/// </summary>
public sealed class EntityPrintTitleRowsTest
{
    /// <summary>
    /// 验证保存重开后重复标题行生效且不插入物理行。
    /// </summary>
    [Fact]
    public void PrintTitleRows_Export_ShouldPersistWithoutInsertingRows()
    {
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(), CreateLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.Worksheet("明细");
        Assert.Equal(1, sheet.PageSetup.FirstRowToRepeatAtTop);
        Assert.Equal(2, sheet.PageSetup.LastRowToRepeatAtTop);
        Assert.Equal("采购单", sheet.Cell("A1").GetString());
        Assert.Equal("Code", sheet.Cell("A3").GetString());
        Assert.Equal("B", sheet.Cell("A5").GetString());
        Assert.True(sheet.Cell("A6").IsEmpty());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证模板既有打印设置和打印区域在新增标题行后保持不变。
    /// </summary>
    [Fact]
    public void PrintTitleRows_Template_ShouldPreserveUnrelatedPrintSettings()
    {
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(), CreateLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var setup = workbook.Worksheet("明细").PageSetup;
        Assert.Equal(XLPaperSize.LetterPaper, setup.PaperSize);
        Assert.Equal(73, setup.Scale);
        var printArea = Assert.Single(setup.PrintAreas);
        Assert.Equal(1, printArea.RangeAddress.FirstAddress.RowNumber);
        Assert.Equal(1, printArea.RangeAddress.FirstAddress.ColumnNumber);
        Assert.Equal(8, printArea.RangeAddress.LastAddress.RowNumber);
        Assert.Equal(2, printArea.RangeAddress.LastAddress.ColumnNumber);
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

    private static MemoryStream CreateTemplate()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Cell("A1").Value = "采购单";
        sheet.PageSetup.PaperSize = XLPaperSize.LetterPaper;
        sheet.PageSetup.Scale = 73;
        sheet.PageSetup.PrintAreas.Add("A1:B8");
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
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
