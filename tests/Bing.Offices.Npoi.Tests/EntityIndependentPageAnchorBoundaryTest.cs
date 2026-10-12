using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using PageDocument = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI XLS/HSSF 与 XLSX/XSSF 独立逐页锚点的边界职责。
/// </summary>
public sealed class EntityIndependentPageAnchorBoundaryTest
{
    /// <summary>
    /// 验证空明细、单条明细和整页明细均保留首页页眉及末页尾部名称。
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
    public void PageHeaderNamed_EmptySingleAndFullPage_ShouldKeepBothAnchors(bool xls, int count)
    {
        var layout = CreateHeaderFooterLayout();
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(count), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(2, workbook.GetAllNames().Count);
            AssertAnchor(workbook, "Head_001", 0);
            AssertAnchor(workbook, "Sign_001", count + 4);
            AssertRow(sheet, 0, "HEADER", "1/1", "");
            AssertRow(sheet, 1, "采购单", "", "");
            AssertRow(sheet, 2, "Group", "Code", "Amount");
            for (var index = 0; index < count; index++)
                AssertRow(sheet, index + 3, "G", "L" + (index + 1), index + 1);
            AssertRow(sheet, count + 3, "", "", "");
            AssertRow(sheet, count + 4, "SIGN", count, Enumerable.Range(1, count).Sum());
            AssertRow(sheet, count + 5, "审核", "", "");
            AssertRow(sheet, count + 6, "END", count, Enumerable.Range(1, count).Sum());
            Assert.Empty(sheet.RowBreaks);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(CreateDocument(count).Lines.Select(item => (item.Group, item.Code, item.Amount)),
            result.Entity.Lines.Select(item => (item.Group, item.Code, item.Amount)));
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证 RemoveObsolete 对非法旧名称引用在两个 NPOI 格式中均拒绝并保全输入。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="kind">非法旧名称类型。</param>
    [Theory]
    [InlineData(true, "foreign")]
    [InlineData(false, "foreign")]
    [InlineData(true, "two-dimensional")]
    [InlineData(false, "two-dimensional")]
    [InlineData(true, "formula")]
    [InlineData(false, "formula")]
    [InlineData(true, "invalid")]
    [InlineData(false, "invalid")]
    [InlineData(true, "ambiguous")]
    [InlineData(false, "ambiguous")]
    public void PageFooterNamed_RemoveObsolete_InvalidOldReference_ShouldPreserveInputs(bool xls,
        string kind)
    {
        using var template = CreateInvalidOldNameTemplate(xls, kind);
        var templateBytes = template.ToArray();
        using var output = new MemoryStream();
        var original = new byte[] { 8, 6, 4, 2 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1),
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
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooter_ShouldRejectFixedCellOverlapBeforeOutput(bool xls)
    {
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder
            .CellNamed("明细", "FixedField", document => document.Title)
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false).PageBreak(1)
                .PageFooter("SIGN", footer => footer.GapRows(1))));
        using var template = CreateTemplateWithFixedAnchor(xls);
        using var output = CreateTarget(new byte[] { 1, 3, 5 });

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), layout,
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
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooter_ShouldRejectRuntimeOverlapWithSecondList(bool xls)
    {
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false).PageBreak(1).PageFooter("SIGN", footer => footer.GapRows(1)))
            .ListRegion("明细", "A3", document => document.Lines, region => region
                .Header(false).End("C4")));
        using var template = CreateTemplate(xls);
        using var output = CreateTarget(new byte[] { 2, 4, 6 });

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        AssertPlanError(exception);
        Assert.Contains("运行时范围", exception.Message, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 2, 4, 6 }, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证每页尾部在 HSSF/XSSF 最大行和列边界外写入时提前失败。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooter_ShouldRejectPhysicalRowAndColumnOverflow(bool xls)
    {
        var lastRow = xls ? "A65536" : "A1048576";
        var lastColumn = xls ? "IT1" : "XFB1";
        var rowLayout = ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", lastRow, document => document.Lines, region => region
                .Header(false).PageBreak(1).PageFooter("SIGN",
                    footer => footer.Cell("A2", "越界"))));
        using var rowTemplate = CreateTemplate(xls);
        using var rowOutput = new MemoryStream();
        var rowException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), rowLayout,
                new ExcelEntityTemplateOptions(rowTemplate, leaveOpen: true), rowOutput));
        AssertPlanError(rowException);
        Assert.Empty(rowOutput.ToArray());
        Assert.True(rowTemplate.CanRead);

        var columnLayout = ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", lastColumn, document => document.Lines, region => region
                .Header(false).PageBreak(1).PageFooter("SIGN",
                    footer => footer.Cell("D1", "越界"))));
        using var columnTemplate = CreateTemplate(xls);
        using var columnOutput = new MemoryStream();
        var columnException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), columnLayout,
                new ExcelEntityTemplateOptions(columnTemplate, leaveOpen: true), columnOutput));
        AssertPlanError(columnException);
        Assert.Empty(columnOutput.ToArray());
        Assert.True(columnTemplate.CanRead);
    }

    /// <summary>
    /// 验证缺失、重复和错位的每页尾部 marker 均报告实际原始坐标。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="mode">损坏模式。</param>
    [Theory]
    [InlineData(true, "missing")]
    [InlineData(false, "missing")]
    [InlineData(true, "duplicate")]
    [InlineData(false, "duplicate")]
    [InlineData(true, "shifted")]
    [InlineData(false, "shifted")]
    public void PageFooter_MarkerCorruption_ShouldReportExactOriginalCoordinate(bool xls,
        string mode)
    {
        var layout = CreateFooterLayout();
        using var template = CreateTemplate(xls);
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), exported);
        using var workbook = WorkbookFactory.Create(new MemoryStream(exported.ToArray(), writable: false));
        Assert.Equal(xls, workbook is HSSFWorkbook);
        var sheet = workbook.GetSheet("明细");
        var footerRows = FindRows(sheet, "SIGN");
        Assert.Equal(2, footerRows.Length);

        var expectedRow = footerRows[1];
        if (mode == "missing")
            sheet.GetRow(footerRows[1]).GetCell(0).SetCellValue("BROKEN");
        else if (mode == "duplicate")
        {
            expectedRow = footerRows[1] + 1;
            GetOrCreateCell(sheet, expectedRow, 0).SetCellValue("SIGN");
        }
        else
        {
            expectedRow = footerRows[1] + 1;
            sheet.GetRow(footerRows[1]).GetCell(0).SetCellValue("BROKEN");
            GetOrCreateCell(sheet, expectedRow, 0).SetCellValue("SIGN");
        }

        using var corrupted = new MemoryStream();
        workbook.Write(corrupted, true);
        var corruptedBytes = corrupted.ToArray();
        corrupted.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(corrupted, layout));

        AssertMarkerError(exception, "明细", expectedRow + 1);
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
    /// 创建指定格式的空白模板。
    /// </summary>
    private static MemoryStream CreateTemplate(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("明细");
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建包含固定字段命名锚点的模板。
    /// </summary>
    private static MemoryStream CreateTemplateWithFixedAnchor(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var name = workbook.CreateName();
        name.NameName = "FixedField";
        name.SheetIndex = workbook.GetSheetIndex(sheet);
        name.RefersToFormula = "'明细'!$A$3";
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建包含非法旧逐页名称的模板。
    /// </summary>
    private static MemoryStream CreateInvalidOldNameTemplate(bool xls, string kind)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("说明");
        var detail = workbook.CreateSheet("明细");
        void AddName(string nameText, int sheetIndex, string formula)
        {
            var name = workbook.CreateName();
            name.SheetIndex = sheetIndex;
            name.NameName = nameText;
            name.RefersToFormula = formula;
        }

        switch (kind)
        {
            case "foreign":
                AddName("Sign_003", 0, "'说明'!$Z$20");
                break;
            case "two-dimensional":
                AddName("Sign_003", 1, "'明细'!$Z$20:$AA$21");
                break;
            case "formula":
                AddName("Sign_003", 1, "OFFSET('明细'!$Z$20,0,0)");
                break;
            case "invalid":
                AddName("Sign_003", 1, "#REF!");
                break;
            case "ambiguous":
                AddName("Sign_003", -1, "'明细'!$Z$20");
                AddName("Sign_003", workbook.GetSheetIndex(detail), "'明细'!$Z$21");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }

        var stream = new MemoryStream();
        workbook.Write(stream, true);
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
    private static int[] FindRows(ISheet sheet, string text) =>
        Enumerable.Range(0, sheet.LastRowNum + 1)
            .Where(row => sheet.GetRow(row)?.GetCell(0)?.StringCellValue == text).ToArray();

    /// <summary>
    /// 获取指定行的单元格并在缺少时创建。
    /// </summary>
    private static ICell GetOrCreateCell(ISheet sheet, int row, int column) =>
        (sheet.GetRow(row) ?? sheet.CreateRow(row)).GetCell(column)
        ?? sheet.GetRow(row).CreateCell(column);

    /// <summary>
    /// 核对整行三个单元格。
    /// </summary>
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
    private static void AssertAnchor(IWorkbook workbook, string name, int row)
    {
        var actual = Assert.Single(workbook.GetAllNames().Where(item => item.NameName == name));
        Assert.Equal("明细!$A$" + (row + 1), actual.RefersToFormula.Replace("'明细'!", "明细!",
            StringComparison.Ordinal));
    }

    /// <summary>
    /// 断言计划阶段结构化错误。
    /// </summary>
    private static void AssertPlanError(BingOfficesConfigurationException exception)
    {
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.False(string.IsNullOrWhiteSpace(exception.Message));
    }

    /// <summary>
    /// 断言 marker 错误的原始工作表坐标。
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
