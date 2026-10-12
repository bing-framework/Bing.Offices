using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Styles;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using PageDocument = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证独立逐页签字区及命名锚点。
/// </summary>
public sealed class EntityIndependentPageAnchorTest
{
    /// <summary>
    /// 验证空页、整页和多页均输出完整签字区并往返明细。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="count">明细数。</param>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 0)]
    [InlineData(true, 1)]
    [InlineData(false, 1)]
    [InlineData(true, 2)]
    [InlineData(false, 2)]
    [InlineData(true, 3)]
    [InlineData(false, 3)]
    [InlineData(true, 4)]
    [InlineData(false, 4)]
    [InlineData(true, 5)]
    [InlineData(false, 5)]
    public void PageFooterNamed_AllPages_ShouldWriteCompleteWorkbookAndRoundTrip(bool xls, int count)
    {
        var source = CreateDocument(count);
        var layout = CreateLayout();
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            var row = 0;
            AssertRow(sheet, row++, "Group", "Code", "Amount");
            var pages = Math.Max(1, (count + 1) / 2);
            var breaks = new List<int>();
            Assert.Equal(pages, workbook.GetAllNames().Count);
            for (var page = 0; page < pages; page++)
            {
                var items = source.Lines.Skip(page * 2).Take(2).ToArray();
                foreach (var item in items)
                    AssertRow(sheet, row++, item.Group, item.Code, item.Amount);
                AssertRow(sheet, row++, "", "", "");
                AssertRow(sheet, row, "SIGN", items.Length, items.Sum(item => item.Amount));
                AssertAnchor(workbook, "Sign_" + (page + 1).ToString("D3"), row);
                Assert.True(workbook.GetFontAt(sheet.GetRow(row).GetCell(0).CellStyle.FontIndex).IsBold);
                row++;
                AssertRow(sheet, row, "审核", "张三", "");
                Assert.Contains(sheet.MergedRegions, merge => merge.FirstRow == row
                    && merge.LastRow == row && merge.FirstColumn == 1 && merge.LastColumn == 2);
                if (page < pages - 1)
                    breaks.Add(row);
                row++;
            }
            AssertRow(sheet, row, "END", count, source.Lines.Sum(item => item.Amount));
            Assert.Equal(row, sheet.LastRowNum);
            Assert.Equal(breaks.ToArray(), sheet.RowBreaks);
            Assert.Equal(pages, sheet.NumMergedRegions);
        }
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
        Assert.True(input.CanRead);
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证无分页小计时每页页眉和签字区均有独立名称。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeaderNamed_WithFooter_ShouldNameFirstAndLastPages(bool xls)
    {
        var layout = CreateLayout(header: true);
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        Assert.Equal(4, workbook.GetAllNames().Count);
        AssertAnchor(workbook, "Head_001", 0);
        AssertAnchor(workbook, "Sign_001", 6);
        AssertAnchor(workbook, "Head_002", 8);
        AssertAnchor(workbook, "Sign_002", 13);
        var sheet = workbook.GetSheet("明细");
        AssertRow(sheet, 0, "HEADER", "1/2", "");
        AssertRow(sheet, 8, "HEADER", "2/2", "");
        AssertRow(sheet, 6, "SIGN", 2, 3m);
        AssertRow(sheet, 13, "SIGN", 1, 3m);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证空明细和单页仍创建首页页眉及末页签字名称。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="count">明细数。</param>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 0)]
    [InlineData(true, 1)]
    [InlineData(false, 1)]
    [InlineData(true, 2)]
    [InlineData(false, 2)]
    public void PageHeaderNamed_EmptyOrSinglePage_ShouldKeepBothAnchors(bool xls, int count)
    {
        var layout = CreateLayout(header: true);
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(count), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        Assert.Equal(2, workbook.GetAllNames().Count);
        AssertAnchor(workbook, "Head_001", 0);
        AssertAnchor(workbook, "Sign_001", count + 4);
        var sheet = workbook.GetSheet("明细");
        AssertRow(sheet, 0, "HEADER", "1/1", "");
        AssertRow(sheet, 1, "采购单", "", "");
        AssertRow(sheet, 2, "Group", "Code", "Amount");
        for (var index = 0; index < count; index++)
            AssertRow(sheet, index + 3, "G", "L" + (index + 1), index + 1);
        AssertRow(sheet, count + 3, "", "", "");
        AssertRow(sheet, count + 4, "SIGN", count, Enumerable.Range(1, count).Sum());
        AssertRow(sheet, count + 5, "审核", "张三", "");
        AssertRow(sheet, count + 6, "END", count, Enumerable.Range(1, count).Sum());
        Assert.Equal(count + 6, sheet.LastRowNum);
        Assert.Empty(sheet.RowBreaks);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal(CreateDocument(count).Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
    }

    /// <summary>
    /// 验证缩短模板仅显式清理当前系列的合法旧名称。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="remove">是否允许清理。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void PageFooterNamed_StaleTemplate_ShouldRespectPolicy(bool xls, bool remove)
    {
        using var template = CreateTemplate(xls, names: true);
        var original = template.ToArray();
        using var output = new MemoryStream();
        output.Write(new byte[] { 7, 4, 2 });
        var layout = CreateLayout(staleNames: remove ? ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete
            : ExcelEntityPageAnchorStaleNamePolicy.Reject);
        if (!remove)
        {
            Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter()
                .ExportForTemplate(CreateDocument(1), layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
            Assert.Equal(new byte[] { 7, 4, 2 }, output.ToArray());
        }
        else
        {
            output.SetLength(0);
            output.Position = 0;
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
            Assert.Equal(new[] { "Keep", "Sign_001" }, workbook.GetAllNames()
                .Select(name => name.NameName).OrderBy(name => name));
            var current = Assert.Single(workbook.GetAllNames().Where(name => name.NameName == "Sign_001"));
            Assert.Equal(0, current.SheetIndex);
            AssertAnchor(workbook, "Sign_001", 3);
        }
        Assert.Equal(original, template.ToArray());
    }

    /// <summary>
    /// 验证缺失的中间签字标记按原始坐标拒绝导入。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooter_MissingMarker_ShouldReportOriginalCoordinate(bool xls)
    {
        var layout = CreateLayout();
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        workbook.GetSheet("明细").GetRow(4).GetCell(0).SetCellValue("BROKEN");
        using var changed = new MemoryStream();
        workbook.Write(changed, true);
        changed.Position = 0;
        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(5, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
    }

    /// <summary>
    /// 验证无最终尾部时完整跳过签字区的空白末行。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="count">明细数。</param>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 0)]
    [InlineData(true, 3)]
    [InlineData(false, 3)]
    public void PageFooter_WithoutFinalFooter_ShouldIncludeBlankHeight(bool xls, int count)
    {
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.PrintTitleRows("明细", "1:1")
            .ListRegion("明细", "A1", document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN", footer => footer.GapRows(1)
                    .Cell("B3", (object)null)).End("C30")));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(count), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        Assert.Equal(Math.Max(1, (count + 1) / 2) + 1, workbook.GetAllNames().Count);
        Assert.Equal(0, workbook.GetSheet("明细").RepeatingRows.FirstRow);
        Assert.Equal(0, workbook.GetSheet("明细").RepeatingRows.LastRow);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(CreateDocument(count).Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
    }

    /// <summary>
    /// 验证组页同边界的小计、签字区顺序及最终明细公式。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public async Task PageFooter_GroupPageBoundary_ShouldExcludeSignatureFromFinalSum(bool xls,
        bool asynchronous)
    {
        var source = CreateDocument(5);
        for (var index = 0; index < source.Lines.Count; index++)
            source.Lines[index].Group = ((char)('A' + index / 2)).ToString();
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .GroupSubtotal(item => item.Group, "GROUP", footer =>
                    footer.Cell("C1", items => items.Sum(item => item.Amount)))
                .PageSubtotalNamed("Page", "PAGE", footer =>
                    footer.Cell("C1", items => items.Sum(item => item.Amount)))
                .PageFooterNamed("Sign", "SIGN", footer => footer.GapRows(1)
                    .Cell("C1", items => 1000 + items.Count).Cell("A2", "审核"))
                .Footer("END", footer => footer.FormulaSumDetailRowsAbove("C1", "C"))));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        var exporter = new NpoiExcelExporter();
        if (asynchronous)
            await exporter.ExportForTemplateAsync(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        else
            exporter.ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.GetSheet("明细");
        var expected = new object[][]
        {
            new object[] { "Group", "Code", "Amount" },
            new object[] { "A", "L1", 1m }, new object[] { "A", "L2", 2m },
            new object[] { "GROUP", "", 3m }, new object[] { "PAGE", "", 3m },
            new object[] { "", "", "" }, new object[] { "SIGN", "", 1002 },
            new object[] { "审核", "", "" },
            new object[] { "B", "L3", 3m }, new object[] { "B", "L4", 4m },
            new object[] { "GROUP", "", 7m }, new object[] { "PAGE", "", 7m },
            new object[] { "", "", "" }, new object[] { "SIGN", "", 1002 },
            new object[] { "审核", "", "" }, new object[] { "C", "L5", 5m },
            new object[] { "GROUP", "", 5m }, new object[] { "", "", "" },
            new object[] { "SIGN", "", 1001 }, new object[] { "审核", "", "" }
        };
        for (var row = 0; row < expected.Length; row++)
            AssertRow(sheet, row, expected[row]);
        Assert.Equal(20, sheet.LastRowNum);
        Assert.Equal("END", sheet.GetRow(20).GetCell(0).StringCellValue);
        Assert.Equal("SUM(C2:C3,C9:C10,C16:C16)", sheet.GetRow(20).GetCell(2).CellFormula);
        Assert.Equal(5, workbook.GetAllNames().Count);
        AssertAnchor(workbook, "Page_001", 4);
        AssertAnchor(workbook, "Page_002", 11);
        AssertAnchor(workbook, "Sign_001", 6);
        AssertAnchor(workbook, "Sign_002", 13);
        AssertAnchor(workbook, "Sign_003", 18);
        Assert.Equal(new[] { 7, 14 }, sheet.RowBreaks);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var importer = new NpoiExcelImporter();
        var result = asynchronous ? await importer.ImportEntityAsync(input, layout)
            : importer.ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
    }

    /// <summary>
    /// 创建每页签字区布局。
    /// </summary>
    /// <param name="header">是否配置实际页眉。</param>
    /// <param name="staleNames">旧名称策略。</param>
    /// <returns>实体布局。</returns>
    private static ExcelEntityLayout<PageDocument> CreateLayout(bool header = false,
        ExcelEntityPageAnchorStaleNamePolicy staleNames = ExcelEntityPageAnchorStaleNamePolicy.Reject) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", header ? "A3" : "A1",
            document => document.Lines, region =>
            {
                region.PageBreak(2);
                if (header)
                    region.PageHeaderNamed("Head", "A1:C3", "A1", "HEADER", configure =>
                        configure.Cell("B1", context => context.PageNumber + "/" + context.PageCount)
                            .Cell("A2", "采购单").Merge("A2:C2"));
                region.PageFooterNamed("Sign", "SIGN", footer => footer.GapRows(1)
                    .MarkerStyle(new ExcelCellStyle { Bold = true })
                    .Cell("B1", items => items.Count).Cell("C1", items => items.Sum(item => item.Amount))
                    .Cell("A2", "审核").Cell("B2", "张三").Merge("B2:C2"), staleNames);
                region.Footer("END", footer => footer.Cell("B1", items => items.Count)
                    .Cell("C1", items => items.Sum(item => item.Amount)));
            }));

    /// <summary>
    /// 创建采购明细。
    /// </summary>
    /// <param name="count">明细数。</param>
    /// <returns>测试单据。</returns>
    private static PageDocument CreateDocument(int count) => new()
    {
        Lines = Enumerable.Range(1, count).Select(index => new PageLine
            { Group = "G", Code = "L" + index, Amount = index }).ToList()
    };

    /// <summary>
    /// 创建指定格式的模板。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="names">是否包含旧名称。</param>
    /// <returns>调用方持有的模板流。</returns>
    private static MemoryStream CreateTemplate(bool xls, bool names = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("明细");
        if (names)
            foreach (var text in new[] { "Sign_001", "Sign_002", "Keep" })
            {
                var name = workbook.CreateName();
                name.NameName = text;
                name.SheetIndex = 0;
                name.RefersToFormula = "'明细'!$Z$20";
            }
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 核对整行三个单元格。
    /// </summary>
    /// <param name="sheet">工作表。</param>
    /// <param name="row">零基行号。</param>
    /// <param name="expected">三个预期值。</param>
    private static void AssertRow(ISheet sheet, int row, params object[] expected)
    {
        for (var column = 0; column < expected.Length; column++)
        {
            var cell = sheet.GetRow(row)?.GetCell(column);
            if (expected[column] is string text)
                Assert.Equal(text, cell == null || cell.CellType == CellType.Blank ? "" : cell.StringCellValue);
            else
                Assert.Equal(Convert.ToDouble(expected[column]), cell.NumericCellValue);
        }
    }

    /// <summary>
    /// 核对单格名称引用。
    /// </summary>
    /// <param name="workbook">工作簿。</param>
    /// <param name="name">预期名称。</param>
    /// <param name="row">零基标记行。</param>
    private static void AssertAnchor(IWorkbook workbook, string name, int row)
    {
        var actual = Assert.Single(workbook.GetAllNames().Where(item => item.NameName == name));
        Assert.Equal("明细!$A$" + (row + 1), actual.RefersToFormula.Replace("'明细'!", "明细!",
            StringComparison.Ordinal));
    }
}
