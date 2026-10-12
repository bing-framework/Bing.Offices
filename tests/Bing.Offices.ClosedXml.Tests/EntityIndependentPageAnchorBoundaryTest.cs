using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;
using PageDocument = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证独立逐页页眉和每页尾部命名锚点的边界与失败保全职责。
/// </summary>
public sealed class EntityIndependentPageAnchorBoundaryTest
{
    /// <summary>
    /// 验证空明细、单条明细和整页明细均保留首页页眉及末页尾部名称。
    /// </summary>
    /// <param name="count">明细数。</param>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void PageHeaderNamed_EmptySingleAndFullPage_ShouldKeepBothAnchors(int count)
    {
        var layout = CreateHeaderFooterLayout();
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(count), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            var headerRows = FindRows(sheet, "HEADER");
            var footerRows = FindRows(sheet, "SIGN");
            Assert.Single(headerRows);
            Assert.Single(footerRows);
            Assert.Equal("1/1", sheet.Cell(headerRows[0], 2).GetString());
            Assert.Equal(count, sheet.Cell(footerRows[0], 2).GetDouble());
            Assert.Equal((double)Enumerable.Range(1, count).Sum(), sheet.Cell(footerRows[0], 3).GetDouble());
            Assert.Equal("END", FindRows(sheet, "END").Length == 1
                ? sheet.Cell(FindRows(sheet, "END")[0], 1).GetString()
                : string.Empty);
            Assert.Empty(sheet.PageSetup.RowBreaks);
            AssertNamedAnchor(workbook, sheet, "Head_001", headerRows[0]);
            AssertNamedAnchor(workbook, sheet, "Sign_001", footerRows[0]);
            Assert.Equal(2, GetNames(workbook, sheet).Count(name =>
                name.Name.StartsWith("Head_", StringComparison.OrdinalIgnoreCase)
                || name.Name.StartsWith("Sign_", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal("Group", sheet.Cell(headerRows[0] + 2, 1).GetString());
            for (var index = 0; index < count; index++)
            {
                var row = headerRows[0] + 3 + index;
                Assert.Equal("L" + (index + 1), sheet.Cell(row, 2).GetString());
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
    /// 验证 RemoveObsolete 对所有非法旧系列引用均在写入前拒绝并保全输入。
    /// </summary>
    /// <param name="kind">非法旧名称类型。</param>
    [Theory]
    [InlineData("foreign")]
    [InlineData("two-dimensional")]
    [InlineData("formula")]
    [InlineData("invalid")]
    [InlineData("ambiguous")]
    public void PageFooterNamed_RemoveObsolete_InvalidOldReference_ShouldPreserveInputs(string kind)
    {
        using var template = CreateInvalidOldNameTemplate(kind);
        var templateBytes = template.ToArray();
        using var output = new MemoryStream();
        var original = new byte[] { 8, 6, 4, 2 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(1),
                CreateFooterLayout(ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        AssertPlanError(exception);
        Assert.Equal(original, output.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证每页尾部与固定字段重叠时在写入前拒绝并保全目标。
    /// </summary>
    [Fact]
    public void PageFooter_ShouldRejectFixedCellOverlapBeforeOutput()
    {
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder
            .CellNamed("明细", "FixedField", document => document.Title)
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false).PageBreak(1)
                .PageFooter("SIGN", footer => footer.GapRows(1))));
        using var template = CreateTemplateWithFixedAnchor();
        using var output = CreateTarget(new byte[] { 1, 3, 5 });

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(1), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        AssertPlanError(exception);
        Assert.Contains("固定单元格与列表区域重叠", exception.Message, StringComparison.Ordinal);
        Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 1, 3, 5 }, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证每页尾部纳入运行时范围后与第二列表区域重叠时拒绝。
    /// </summary>
    [Fact]
    public void PageFooter_ShouldRejectRuntimeOverlapWithSecondList()
    {
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false).PageBreak(1).PageFooter("SIGN", footer => footer.GapRows(1)))
            .ListRegion("明细", "A3", document => document.Lines, region => region
                .Header(false).End("C4")));
        using var output = CreateTarget(new byte[] { 2, 4, 6 });

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument(1), layout, output));

        AssertPlanError(exception);
        Assert.Contains("运行时范围", exception.Message, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 2, 4, 6 }, output.ToArray());
    }

    /// <summary>
    /// 验证每页尾部在工作表最大行和列边界外写入时提前返回结构化错误。
    /// </summary>
    [Fact]
    public void PageFooter_ShouldRejectPhysicalRowAndColumnOverflow()
    {
        var rowLayout = ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", "A1048576", document => document.Lines, region => region
                .Header(false).PageBreak(1)
                .PageFooter("SIGN", footer => footer.Cell("A2", "越界"))));
        using var rowOutput = new MemoryStream();
        var rowException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument(1), rowLayout, rowOutput));
        AssertPlanError(rowException);
        Assert.Empty(rowOutput.ToArray());

        var columnLayout = ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", "XFB1", document => document.Lines, region => region
                .Header(false).PageBreak(1)
                .PageFooter("SIGN", footer => footer.Cell("D1", "越界"))));
        using var columnOutput = new MemoryStream();
        var columnException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument(1), columnLayout, columnOutput));
        AssertPlanError(columnException);
        Assert.Empty(columnOutput.ToArray());
    }

    /// <summary>
    /// 验证缺失、重复和错位的每页尾部 marker 均报告实际原始坐标。
    /// </summary>
    /// <param name="mode">损坏模式：缺失、重复或错位。</param>
    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("shifted")]
    public void PageFooter_MarkerCorruption_ShouldReportExactOriginalCoordinate(string mode)
    {
        var layout = CreateFooterLayout();
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument(3), layout, exported);
        using var workbook = new XLWorkbook(new MemoryStream(exported.ToArray(), writable: false));
        var sheet = workbook.Worksheet("明细");
        var footerRows = FindRows(sheet, "SIGN");
        Assert.Equal(2, footerRows.Length);

        var expectedRow = footerRows[1];
        if (mode == "missing")
            sheet.Cell(footerRows[1], 1).Value = "BROKEN";
        else if (mode == "duplicate")
        {
            expectedRow = footerRows[1] + 1;
            sheet.Cell(expectedRow, 1).Value = "SIGN";
        }
        else
        {
            expectedRow = footerRows[1] + 1;
            sheet.Cell(footerRows[1], 1).Value = "BROKEN";
            sheet.Cell(expectedRow, 1).Value = "SIGN";
        }

        using var corrupted = new MemoryStream();
        workbook.SaveAs(corrupted);
        var corruptedBytes = corrupted.ToArray();
        corrupted.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(corrupted, layout));

        AssertMarkerError(exception, "明细", expectedRow);
        Assert.Equal(corruptedBytes, corrupted.ToArray());
        Assert.True(corrupted.CanRead);
    }

    /// <summary>
    /// 创建页眉和每页尾部均带名称的布局。
    /// </summary>
    private static ExcelEntityLayout<PageDocument> CreateHeaderFooterLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A3",
            document => document.Lines, region => region
            .PageBreak(2)
            .PageHeaderNamed("Head", "A1:C3", "A1", "HEADER", configure =>
                configure.Cell("B1", context => context.PageNumber + "/" + context.PageCount)
                    .Cell("A2", "采购单").Merge("A2:C2"))
            .PageFooterNamed("Sign", "SIGN", footer => footer.GapRows(1)
                .Cell("B1", items => items.Count)
                .Cell("C1", items => items.Sum(item => item.Amount))
                .Cell("A2", "审核"))
            .Footer("END", footer => footer.Cell("B1", items => items.Count)
                .Cell("C1", items => items.Sum(item => item.Amount)))));

    /// <summary>
    /// 创建只包含每页尾部的布局。
    /// </summary>
    /// <param name="staleNames">过期名称策略。</param>
    private static ExcelEntityLayout<PageDocument> CreateFooterLayout(
        ExcelEntityPageAnchorStaleNamePolicy staleNames = ExcelEntityPageAnchorStaleNamePolicy.Reject) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
            .PageBreak(2)
            .PageFooterNamed("Sign", "SIGN", footer => footer.GapRows(1)
                .Cell("B1", items => items.Count)
                .Cell("A2", "审核"), staleNames)
            .Footer("END")));

    /// <summary>
    /// 创建指定数量的采购明细。
    /// </summary>
    private static PageDocument CreateDocument(int count = 3) => new()
    {
        Lines = Enumerable.Range(1, count).Select(index => new PageLine
        {
            Group = "G",
            Code = "L" + index,
            Amount = index
        }).ToList()
    };

    /// <summary>
    /// 创建空白 ClosedXML 模板。
    /// </summary>
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
    /// 创建包含固定字段命名锚点的模板。
    /// </summary>
    private static MemoryStream CreateTemplateWithFixedAnchor()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        workbook.DefinedNames.Add("FixedField", sheet.Range("A3"));
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建包含非法旧逐页名称的模板。
    /// </summary>
    /// <param name="kind">非法名称类型。</param>
    private static MemoryStream CreateInvalidOldNameTemplate(string kind)
    {
        using var workbook = new XLWorkbook();
        var note = workbook.Worksheets.Add("说明");
        var detail = workbook.Worksheets.Add("明细");
        switch (kind)
        {
            case "foreign":
                workbook.DefinedNames.Add("Sign_003", note.Range("Z20"));
                break;
            case "two-dimensional":
                workbook.DefinedNames.Add("Sign_003", detail.Range("Z20:AA21"));
                break;
            case "formula":
                workbook.DefinedNames.Add("Sign_003", "OFFSET('明细'!$Z$20,0,0)");
                break;
            case "invalid":
                workbook.DefinedNames.Add("Sign_003", "#REF!");
                break;
            case "ambiguous":
                workbook.DefinedNames.Add("Sign_003", detail.Range("Z20"));
                detail.DefinedNames.Add("Sign_003", detail.Range("Z21"));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }

        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建带有旧字节的目标流。
    /// </summary>
    private static MemoryStream CreateTarget(byte[] original)
    {
        var stream = new MemoryStream();
        stream.Write(original, 0, original.Length);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 查找第一列中的标记行。
    /// </summary>
    private static int[] FindRows(IXLWorksheet sheet, string text) =>
        Enumerable.Range(1, sheet.LastRowUsed()?.RowNumber() ?? 0)
            .Where(row => sheet.Cell(row, 1).GetString() == text).ToArray();

    /// <summary>
    /// 读取工作簿和工作表级的有效名称。
    /// </summary>
    private static IReadOnlyList<IXLDefinedName> GetNames(XLWorkbook workbook, IXLWorksheet sheet) =>
        workbook.DefinedNames.ValidNamedRanges()
            .Concat(sheet.DefinedNames.ValidNamedRanges()).ToArray();

    /// <summary>
    /// 核对命名锚点指向指定 marker 单元格。
    /// </summary>
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

    /// <summary>
    /// 断言计划阶段结构化错误及原始工作表坐标。
    /// </summary>
    private static void AssertPlanError(BingOfficesConfigurationException exception)
    {
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.False(string.IsNullOrWhiteSpace(exception.Message));
    }

    /// <summary>
    /// 断言 marker 导入错误的原始工作表坐标。
    /// </summary>
    private static void AssertMarkerError(BingOfficesConfigurationException exception,
        string sheetName, int row)
    {
        AssertPlanError(exception);
        Assert.Equal(sheetName, exception.SheetName);
        Assert.Equal(row, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }
}
