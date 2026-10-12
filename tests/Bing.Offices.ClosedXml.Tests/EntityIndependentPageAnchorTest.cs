using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Styles;
using ClosedXML.Excel;
using Xunit;
using PageDocument = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML 独立逐页签字区及命名锚点。
/// </summary>
public sealed class EntityIndependentPageAnchorTest
{
    /// <summary>
    /// 验证空页、整页和多页均输出完整签字区、名称和分页符，并完整往返明细。
    /// </summary>
    /// <param name="count">明细数。</param>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void PageFooterNamed_AllPages_ShouldWriteCompleteWorkbookAndRoundTrip(int count)
    {
        var source = CreateDocument(count);
        var layout = CreateLayout();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            var pages = Math.Max(1, (count + 1) / 2);
            var expected = new List<string> { "Group" };
            var signRows = new List<int>();
            var breakRows = new List<int>();
            var nextRow = 2;
            for (var page = 0; page < pages; page++)
            {
                var items = source.Lines.Skip(page * 2).Take(2).ToArray();
                foreach (var item in items)
                {
                    expected.Add(item.Group);
                    nextRow++;
                }
                expected.Add(string.Empty);
                nextRow++;
                signRows.Add(nextRow);
                expected.Add("SIGN");
                nextRow++;
                expected.Add("审核");
                if (page < pages - 1)
                    breakRows.Add(nextRow);
                nextRow++;
            }
            expected.Add("END");

            Assert.Equal(expected, ReadFirstColumn(sheet, expected.Count));
            Assert.Equal(pages, signRows.Count);
            Assert.Equal(breakRows.ToArray(), sheet.PageSetup.RowBreaks);
            Assert.Equal(pages, GetNames(workbook, sheet)
                .Count(name => name.Name.StartsWith("Sign_", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal(pages, sheet.MergedRanges.Count);

            for (var page = 0; page < pages; page++)
            {
                var items = source.Lines.Skip(page * 2).Take(2).ToArray();
                var signRow = signRows[page];
                Assert.Equal((double)items.Length, sheet.Cell(signRow, 2).GetDouble());
                Assert.Equal((double)items.Sum(item => item.Amount), sheet.Cell(signRow, 3).GetDouble());
                Assert.True(sheet.Cell(signRow, 1).Style.Font.Bold);
                Assert.Contains(sheet.MergedRanges, merge => merge.RangeAddress.FirstAddress.RowNumber == signRow + 1
                    && merge.RangeAddress.LastAddress.RowNumber == signRow + 1
                    && merge.RangeAddress.FirstAddress.ColumnNumber == 2
                    && merge.RangeAddress.LastAddress.ColumnNumber == 3);
                AssertNamedAnchor(workbook, sheet, "Sign_" + (page + 1).ToString("D3"), signRow);
            }

            var finalRow = expected.Count;
            Assert.Equal((double)count, sheet.Cell(finalRow, 2).GetDouble());
            Assert.Equal((double)source.Lines.Sum(item => item.Amount),
                sheet.Cell(finalRow, 3).GetDouble());
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证没有分页小计时页眉和每页尾部仍分别创建独立名称并完整往返。
    /// </summary>
    [Fact]
    public void PageHeaderNamed_WithFooter_ShouldNameAllPagesAndRoundTrip()
    {
        var layout = CreateLayout(header: true);
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument(3), layout, output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            var headerRows = FindRows(sheet, "HEADER");
            var signRows = FindRows(sheet, "SIGN");

            Assert.Equal(new[] { "1/2", "2/2" }, headerRows.Select(row => sheet.Cell(row, 2).GetString()));
            Assert.Equal(new[] { 2, 1 }, signRows.Select(row => (int)sheet.Cell(row, 2).GetDouble()));
            Assert.Equal(2, headerRows.Length);
            Assert.Equal(2, signRows.Length);
            Assert.Equal(4, GetNames(workbook, sheet).Count(name =>
                name.Name.StartsWith("Head_", StringComparison.OrdinalIgnoreCase)
                || name.Name.StartsWith("Sign_", StringComparison.OrdinalIgnoreCase)));
            AssertNamedAnchor(workbook, sheet, "Head_001", headerRows[0]);
            AssertNamedAnchor(workbook, sheet, "Head_002", headerRows[1]);
            AssertNamedAnchor(workbook, sheet, "Sign_001", signRows[0]);
            AssertNamedAnchor(workbook, sheet, "Sign_002", signRows[1]);
            foreach (var row in headerRows)
                Assert.Contains(sheet.MergedRanges, merge => merge.RangeAddress.FirstAddress.RowNumber == row + 1
                    && merge.RangeAddress.LastAddress.RowNumber == row + 1
                    && merge.RangeAddress.FirstAddress.ColumnNumber == 1
                    && merge.RangeAddress.LastAddress.ColumnNumber == 3);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证局部名称模板按 Reject 或 RemoveObsolete 策略处理旧系列，并保留 Keep 和输入字节。
    /// </summary>
    /// <param name="remove">是否显式删除当前系列中已过期的名称。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageFooterNamed_LocalStaleTemplate_ShouldRespectPolicyAndPreserveInput(bool remove)
    {
        using var template = CreateTemplateWithLocalNames();
        var original = template.ToArray();
        var layout = CreateLayout(staleNames: remove
            ? ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete
            : ExcelEntityPageAnchorStaleNamePolicy.Reject);
        using var output = new MemoryStream();
        output.Write(new byte[] { 7, 4, 2 });
        output.Position = 0;

        if (!remove)
        {
            Assert.Throws<BingOfficesConfigurationException>(() => new ClosedXmlExcelExporter()
                .ExportForTemplate(CreateDocument(1), layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
            Assert.Equal(new byte[] { 7, 4, 2 }, output.ToArray());
        }
        else
        {
            output.SetLength(0);
            output.Position = 0;
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(1), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
            var sheet = workbook.Worksheet("明细");
            Assert.Equal(new[] { "Keep", "Sign_001" }, GetNames(workbook, sheet)
                .Select(name => name.Name).OrderBy(name => name));
            AssertNamedAnchor(workbook, sheet, "Sign_001", FindRows(sheet, "SIGN").Single());
        }

        Assert.Equal(original, template.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证缺失的第二页签字 marker 返回完整原始坐标。
    /// </summary>
    [Fact]
    public void PageFooter_MissingMarker_ShouldReportOriginalCoordinate()
    {
        var layout = CreateLayout();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument(3), layout, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.Worksheet("明细");
        var signRows = FindRows(sheet, "SIGN");
        Assert.Equal(2, signRows.Length);
        sheet.Cell(signRows[1], 1).Value = "BROKEN";
        using var changed = new MemoryStream();
        workbook.SaveAs(changed);
        changed.Position = 0;

        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(signRows[1], error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Contains("SIGN", error.Message);
    }

    /// <summary>
    /// 验证没有最终尾部时签字区的空白高度、打印标题设置和明细往返。
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(3)]
    public void PageFooter_WithoutFinalFooter_ShouldIncludeBlankHeightAndPrintTitles(int count)
    {
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.PrintTitleRows("明细", "1:1")
            .ListRegion("明细", "A1", document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN", footer => footer.GapRows(1)
                    .Cell("B3", (object)null)).End("C30")));
        using var template = CreateTemplateWithPrintTitle();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(count), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            var pages = Math.Max(1, (count + 1) / 2);
            Assert.Equal(pages, FindRows(sheet, "SIGN").Length);
            Assert.Equal(pages, GetNames(workbook, sheet)
                .Count(name => name.Name.StartsWith("Sign_", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal(1, sheet.PageSetup.FirstRowToRepeatAtTop);
            Assert.Equal(1, sheet.PageSetup.LastRowToRepeatAtTop);
            Assert.DoesNotContain(sheet.CellsUsed(), cell => cell.GetString() == "END");
            foreach (var row in FindRows(sheet, "SIGN"))
            {
                Assert.True(sheet.Cell(row - 1, 1).IsEmpty());
                Assert.True(sheet.Cell(row, 1).Style.Font.Bold == false);
            }
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(CreateDocument(count).Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证组小计、页小计和签字区同页边界的物理顺序及最终明细公式。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出和导入入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageFooter_GroupPageBoundary_ShouldExcludeSignatureFromFinalSum(bool asynchronous)
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
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();
        if (asynchronous)
            await exporter.ExportForTemplateAsync(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        else
            exporter.ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            var expected = new[]
            {
                "Group", "A", "A", "GROUP", "PAGE", "", "SIGN", "审核",
                "B", "B", "GROUP", "PAGE", "", "SIGN", "审核",
                "C", "GROUP", "", "SIGN", "审核", "END"
            };
            Assert.Equal(expected, ReadFirstColumn(sheet, expected.Length));
            Assert.Equal(new[] { 8, 15 }, sheet.PageSetup.RowBreaks);
            var finalRow = expected.Length;
            Assert.Equal("SUM(C2:C3,C9:C10,C16:C16)", sheet.Cell(finalRow, 3).FormulaA1);
            Assert.DoesNotContain("1000", sheet.Cell(finalRow, 3).FormulaA1);
            Assert.Equal(new[] { "Page_001", "Page_002" },
                GetNames(workbook, sheet).Where(name => name.Name.StartsWith("Page_", StringComparison.Ordinal))
                    .Select(name => name.Name).OrderBy(name => name));
            Assert.Equal(new[] { "Sign_001", "Sign_002", "Sign_003" },
                GetNames(workbook, sheet).Where(name => name.Name.StartsWith("Sign_", StringComparison.Ordinal))
                    .Select(name => name.Name).OrderBy(name => name));
            Assert.Equal(5, GetNames(workbook, sheet).Count(name =>
                name.Name.StartsWith("Page_", StringComparison.Ordinal)
                || name.Name.StartsWith("Sign_", StringComparison.Ordinal)));
            var expectedRows = new[]
            {
                "Group|Code|Amount", "A|L1|1", "A|L2|2", "GROUP||3", "PAGE||3", "||",
                "SIGN||1002", "审核||", "B|L3|3", "B|L4|4", "GROUP||7", "PAGE||7", "||",
                "SIGN||1002", "审核||", "C|L5|5", "GROUP||5", "||", "SIGN||1001", "审核||",
                "END||=SUM(C2:C3,C9:C10,C16:C16)"
            };
            Assert.Equal(expectedRows, Enumerable.Range(1, expectedRows.Length).Select(row =>
                string.Join("|", Enumerable.Range(1, 3).Select(column =>
                {
                    var cell = sheet.Cell(row, column);
                    return cell.HasFormula ? "=" + cell.FormulaA1 : cell.GetString();
                }))));
            Assert.Equal(2, sheet.Cell(7, 3).GetDouble() - 1000);
            Assert.Equal(2, sheet.Cell(14, 3).GetDouble() - 1000);
            Assert.Equal(1, sheet.Cell(19, 3).GetDouble() - 1000);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var importer = new ClosedXmlExcelImporter();
        var result = asynchronous ? await importer.ImportEntityAsync(input, layout)
            : importer.ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 创建每页签字区布局。
    /// </summary>
    /// <param name="header">是否配置实际重复页眉。</param>
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
    /// 创建含局部旧名称的 ClosedXML 模板。
    /// </summary>
    /// <returns>调用方持有的模板流。</returns>
    private static MemoryStream CreateTemplateWithLocalNames()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("说明");
        var sheet = workbook.Worksheets.Add("明细");
        foreach (var (name, address) in new[]
        {
            ("Sign_001", "Z20"), ("Sign_002", "Z21"), ("Keep", "Z22")
        })
            sheet.DefinedNames.Add(name, sheet.Range(address));
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建带有打印标题工作表的模板。
    /// </summary>
    /// <returns>调用方持有的模板流。</returns>
    private static MemoryStream CreateTemplateWithPrintTitle()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("明细");
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建单工作表模板。
    /// </summary>
    /// <returns>调用方持有的模板流。</returns>
    private static MemoryStream CreateTemplate()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("明细");
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 读取工作表第一列的完整物理行序列。
    /// </summary>
    /// <param name="sheet">工作表。</param>
    /// <param name="rowCount">读取的行数。</param>
    /// <returns>第一列文本。</returns>
    private static string[] ReadFirstColumn(IXLWorksheet sheet, int rowCount) =>
        Enumerable.Range(1, rowCount).Select(row => sheet.Cell(row, 1).GetString()).ToArray();

    /// <summary>
    /// 查找第一列中等于指定文本的行号。
    /// </summary>
    /// <param name="sheet">工作表。</param>
    /// <param name="text">待查找文本。</param>
    /// <returns>一基行号。</returns>
    private static int[] FindRows(IXLWorksheet sheet, string text) =>
        Enumerable.Range(1, sheet.LastRowUsed()?.RowNumber() ?? 0)
            .Where(row => sheet.Cell(row, 1).GetString() == text).ToArray();

    /// <summary>
    /// 读取工作簿和工作表级的有效名称。
    /// </summary>
    /// <param name="workbook">工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <returns>有效名称集合。</returns>
    private static IReadOnlyList<IXLDefinedName> GetNames(XLWorkbook workbook, IXLWorksheet sheet) =>
        workbook.DefinedNames.ValidNamedRanges()
            .Concat(sheet.DefinedNames.ValidNamedRanges()).ToArray();

    /// <summary>
    /// 核对名称指向指定工作表的单格 marker。
    /// </summary>
    /// <param name="workbook">工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="nameText">名称文本。</param>
    /// <param name="row">marker 一基行号。</param>
    private static void AssertNamedAnchor(XLWorkbook workbook, IXLWorksheet sheet,
        string nameText, int row)
    {
        var name = Assert.Single(GetNames(workbook, sheet).Where(item => item.Name == nameText));
        var range = Assert.Single(name.Ranges);
        Assert.Equal(sheet.Name, range.Worksheet.Name);
        Assert.Equal(row, range.RangeAddress.FirstAddress.RowNumber);
        Assert.Equal(1, range.RangeAddress.FirstAddress.ColumnNumber);
        Assert.Equal(range.RangeAddress.FirstAddress, range.RangeAddress.LastAddress);
    }
}
