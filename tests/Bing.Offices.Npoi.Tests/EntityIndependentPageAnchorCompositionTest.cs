using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using PageDocument = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageLine;
using PhysicalDocument = Bing.Offices.Tests.EntityRepeatedPageHeaderTest.PhysicalHeaderDocument;
using PhysicalLine = Bing.Offices.Tests.EntityRepeatedPageHeaderTest.PhysicalHeaderLine;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 独立页尾名称与长页序列、模板名称及物理列组合的职责。
/// </summary>
public sealed class EntityIndependentPageAnchorCompositionTest
{
    /// <summary>
    /// 验证 1001 个单条逻辑页的每个页尾名称、引用和明细往返。
    /// </summary>
    /// <remarks>
    /// 同时覆盖四位页码名称和 HSSF/XSSF 两种工作簿格式。
    /// </remarks>
    /// <param name="xls">是否使用 HSSF；否则使用 XSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooterNamed_1001SingleItemPages_ShouldKeepAllNamesAndRoundTrip(bool xls)
    {
        var source = CreatePageDocument(1001);
        var layout = CreateNamedFooterLayout("Sign");
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            var names = workbook.GetAllNames().Where(name =>
                name.NameName.StartsWith("Sign_", StringComparison.OrdinalIgnoreCase)).ToArray();
            Assert.Equal(1001, names.Length);
            Assert.Equal(1001, workbook.GetAllNames().Count);
            foreach (var page in new[] { 999, 1000, 1001 })
                AssertAnchor(workbook, "Sign_" + page.ToString("D3"),
                    FindRow(sheet, "SIGN", page - 1));

            for (var page = 1; page <= 1001; page++)
            {
                var footerRow = FindRow(sheet, "SIGN", page - 1);
                Assert.Equal("SIGN", sheet.GetRow(footerRow).GetCell(0).StringCellValue);
                Assert.Equal(1d, sheet.GetRow(footerRow).GetCell(1).NumericCellValue);
                Assert.Equal("L" + page, sheet.GetRow(footerRow - 1).GetCell(1).StringCellValue);
                AssertAnchor(workbook, "Sign_" + page.ToString("D3"), footerRow);
            }
            Assert.Equal("END", sheet.GetRow(sheet.LastRowNum).GetCell(0).StringCellValue);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => (line.Code, line.Amount)),
            result.Entity.Lines.Select(line => (line.Code, line.Amount)));
    }

    /// <summary>
    /// 验证第 1000 页首次超出名称上限时不写入目标和模板。
    /// </summary>
    /// <remarks>
    /// 失败应保留已有目标内容及输入模板字节。
    /// </remarks>
    /// <param name="xls">是否使用 HSSF；否则使用 XSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooterNamed_LongPrefix_ShouldFailBeforeWritingAtPage1000(bool xls)
    {
        var layout = CreateNamedFooterLayout(new string('P', 251));
        using var template = CreateTemplate(xls);
        var templateBytes = template.ToArray();
        using var output = new MemoryStream();
        var original = new byte[] { 7, 4, 2 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreatePageDocument(1000), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, error.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
    }

    /// <summary>
    /// 验证 RemoveObsolete 不清理外部旧名称且写入前保全目标。
    /// </summary>
    /// <remarks>
    /// 外部名称由同格式外部工作簿关联产生。
    /// </remarks>
    /// <param name="xls">是否使用 HSSF；否则使用 XSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooterNamed_RemoveObsolete_ExternalOldName_ShouldPreserveInputs(bool xls)
    {
        using var template = CreateExternalNameTemplate(xls);
        var templateBytes = template.ToArray();
        using var output = new MemoryStream();
        var original = new byte[] { 8, 6, 4, 2 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreatePageDocument(1),
                CreateNamedFooterLayout("Sign", ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, error.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
    }

    /// <summary>
    /// 验证命名模板起点与动态列、计算列及校验的物理坐标组合。
    /// </summary>
    /// <remarks>
    /// 同时检查每页 Footer 计数、最终 Footer 计数和完整成功数据。
    /// </remarks>
    /// <param name="xls">是否使用 HSSF；否则使用 XSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ListRegionNamed_DynamicCalculatedAndPageFooter_ShouldPreservePhysicalCoordinates(bool xls)
    {
        var source = CreatePhysicalDocument();
        var layout = CreatePhysicalLayout();
        using var template = CreateNamedListTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("序号", sheet.GetRow(2).GetCell(1).StringCellValue);
            Assert.Equal("Code", sheet.GetRow(2).GetCell(2).StringCellValue);
            Assert.Equal("Quantity", sheet.GetRow(2).GetCell(3).StringCellValue);
            Assert.Equal("颜色", sheet.GetRow(2).GetCell(4).StringCellValue);
            Assert.Equal("L1", sheet.GetRow(3).GetCell(2).StringCellValue);
            Assert.Equal("L2", sheet.GetRow(4).GetCell(2).StringCellValue);
            Assert.Equal(1d, sheet.GetRow(3).GetCell(3).NumericCellValue);
            Assert.Equal(2d, sheet.GetRow(4).GetCell(3).NumericCellValue);
            Assert.Equal("红", sheet.GetRow(3).GetCell(4).StringCellValue);
            Assert.Equal("蓝", sheet.GetRow(4).GetCell(4).StringCellValue);
            Assert.Equal("SIGN", sheet.GetRow(5).GetCell(1).StringCellValue);
            Assert.Equal(2d, sheet.GetRow(5).GetCell(2).NumericCellValue);
            Assert.Equal("L3", sheet.GetRow(6).GetCell(2).StringCellValue);
            Assert.Equal(3d, sheet.GetRow(6).GetCell(1).NumericCellValue);
            Assert.Equal("绿", sheet.GetRow(6).GetCell(4).StringCellValue);
            Assert.Equal("SIGN", sheet.GetRow(7).GetCell(1).StringCellValue);
            Assert.Equal(1d, sheet.GetRow(7).GetCell(2).NumericCellValue);
            Assert.Equal("END", sheet.GetRow(8).GetCell(1).StringCellValue);
            Assert.Equal(3d, sheet.GetRow(8).GetCell(2).NumericCellValue);
        }

        using (var input = new MemoryStream(output.ToArray(), writable: false))
        {
            var result = new NpoiExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
            Assert.Equal(new[] { "红", "蓝", "绿" }, result.Entity.Lines.Select(line =>
                (string)line.Values["color"]));
        }

        foreach (var nativeFailure in new[] { true, false })
        {
            using var changed = new MemoryStream();
            using (var workbook = WorkbookFactory.Create(
                new MemoryStream(output.ToArray(), writable: false)))
            {
                var sheet = workbook.GetSheet("明细");
                sheet.GetRow(6).GetCell(2).SetCellValue(nativeFailure ? "非法" : "L3");
                sheet.GetRow(6).GetCell(4).SetCellValue("红");
                workbook.Write(changed, true);
            }
            changed.Position = 0;
            var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
                ExcelImportValidationMode.ConfiguredAndWorkbook, ExcelUnsupportedFeaturePolicy.Fail);
            var invalid = new NpoiExcelImporter().ImportEntity(changed, layout, options);
            var error = Assert.Single(invalid.Errors);
            Assert.Equal(nativeFailure ? ExcelImportErrorCode.WorkbookValidation
                : ExcelImportErrorCode.Validation, error.Code);
            Assert.Equal("明细", error.SheetName);
            Assert.Equal(7, error.RowIndex);
            Assert.Equal(nativeFailure ? 3 : 5, error.ColumnIndex);
            Assert.Equal(nativeFailure ? "Code" : "color", error.ColumnKey);
            Assert.Equal(new[] { "L1", "L2" }, invalid.Entity.Lines.Select(line => line.Code));
            Assert.Equal(new[] { 1, 2 }, invalid.Entity.Lines.Select(line => line.Quantity));
            Assert.Equal(new[] { "红", "蓝" }, invalid.Entity.Lines.Select(line =>
                (string)line.Values["color"]));
        }
    }

    /// <summary>
    /// 验证原生 Code 失败后动态 Unique 预留回滚，使页尾后的下一条重复动态值仍可成功导入。
    /// </summary>
    /// <param name="xls">是否使用 HSSF；否则使用 XSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DynamicUnique_NativeFailure_ShouldRollbackReservationForNextRow(bool xls)
    {
        var source = CreatePhysicalDocument(new[] { "红", "蓝", "黄", "绿" });
        var layout = CreatePhysicalLayout(7, new[] { "L1", "L2", "L3", "L4" });
        using var template = CreateNamedListTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("L3", sheet.GetRow(6).GetCell(2).StringCellValue);
            Assert.Equal("L4", sheet.GetRow(7).GetCell(2).StringCellValue);
            Assert.Equal("SIGN", sheet.GetRow(5).GetCell(1).StringCellValue);
            Assert.Equal(2d, sheet.GetRow(5).GetCell(2).NumericCellValue);
            Assert.Equal("SIGN", sheet.GetRow(8).GetCell(1).StringCellValue);
            Assert.Equal(2d, sheet.GetRow(8).GetCell(2).NumericCellValue);
            Assert.Equal("END", sheet.GetRow(9).GetCell(1).StringCellValue);
            Assert.Equal(4d, sheet.GetRow(9).GetCell(2).NumericCellValue);
        }

        using var changed = new MemoryStream();
        using (var workbook = WorkbookFactory.Create(
            new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            sheet.GetRow(6).GetCell(2).SetCellValue("非法");
            sheet.GetRow(6).GetCell(4).SetCellValue("重试");
            sheet.GetRow(7).GetCell(2).SetCellValue("L4");
            sheet.GetRow(7).GetCell(4).SetCellValue("重试");
            workbook.Write(changed, true);
        }
        changed.Position = 0;
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredAndWorkbook, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new NpoiExcelImporter().ImportEntity(changed, layout, options);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(7, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(new[] { "L1", "L2", "L4" }, result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(new[] { "红", "蓝", "重试" }, result.Entity.Lines.Select(line =>
            (string)line.Values["color"]));
    }

    /// <summary>
    /// 创建指定名称前缀的单列分页尾部布局。
    /// </summary>
    /// <remarks>
    /// 每页 Footer 和最终 Footer 均写入当前页明细计数。
    /// </remarks>
    /// <param name="prefix">页尾名称前缀。</param>
    /// <param name="staleNames">旧名称处理策略。</param>
    /// <returns>配置完成的实体布局。</returns>
    private static ExcelEntityLayout<PageDocument> CreateNamedFooterLayout(string prefix,
        ExcelEntityPageAnchorStaleNamePolicy staleNames = ExcelEntityPageAnchorStaleNamePolicy.Reject) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.Header(false).PageBreak(1)
                .PageFooterNamed(prefix, "SIGN", footer => footer.Cell("B1", lines => lines.Count), staleNames)
                .Footer("END", footer => footer.Cell("B1", lines => lines.Count))));

    /// <summary>
    /// 创建带命名起点、动态列、计算列和双来源校验的布局。
    /// </summary>
    /// <remarks>
    /// Footer 使用相对 B1 的计数以覆盖 B3 起点坐标。
    /// </remarks>
    /// <param name="validationEndRow">原生 Code 校验的最后一基行号。</param>
    /// <param name="allowedCodes">原生 Code 校验允许的值。</param>
    /// <returns>配置完成的实体布局。</returns>
    private static ExcelEntityLayout<PhysicalDocument> CreatePhysicalLayout(int validationEndRow = 6,
        IReadOnlyList<string> allowedCodes = null) =>
        ExcelEntity.Layout<PhysicalDocument>(builder => builder
            .ListRegionNamed("明细", "Lines", document => document.Lines, region => region
                .CalculatedColumn<int>("sequence", "序号", context => context.Index + 1,
                    column => column.Placement(ExcelColumnPlacement.Before("Code")))
                .DynamicColumnGroup("商品", line => line.Values, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "color", Title = "颜色", DataType = typeof(string),
                        Placement = ExcelColumnPlacement.After("Quantity"),
                        ValidationRules = new[]
                        {
                            new ExcelMappingDynamicValidationConfiguration
                            {
                                Name = "unique", IgnoreEmpty = false
                            }
                        }
                    }
                })
                .PageBreak(2)
                .PageFooter("SIGN", footer => footer.Cell("B1", lines => lines.Count))
                .Footer("END", footer => footer.Cell("B1", lines => lines.Count)))
            .LookupList("AllowedCodes", allowedCodes ?? new[] { "L1", "L2", "L3" })
            .DataValidation("明细", new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                {
                    StartRow = 3, EndRow = validationEndRow, StartColumn = 2, EndColumn = 2
                },
                Type = ExcelDataValidationType.List, ListSourceName = "AllowedCodes"
            }));

    /// <summary>
    /// 创建指定数量的分页明细。
    /// </summary>
    /// <remarks>
    /// 每条明细使用稳定的分组、编码和金额值。
    /// </remarks>
    private static PageDocument CreatePageDocument(int count) => new()
    {
        Lines = Enumerable.Range(1, count).Select(index => new PageLine
            { Group = "G", Code = "L" + index, Amount = index }).ToList()
    };

    /// <summary>
    /// 创建动态列物理校验明细。
    /// </summary>
    /// <remarks>
    /// 每条明细包含用于 Unique 校验的颜色值。
    /// </remarks>
    /// <param name="colors">每条明细对应的动态颜色。</param>
    /// <returns>动态列物理校验实体。</returns>
    private static PhysicalDocument CreatePhysicalDocument(IReadOnlyList<string> colors = null) => new()
    {
        Lines = (colors ?? new[] { "红", "蓝", "绿" }).Select((color, index) => new PhysicalLine
        {
            Code = "L" + (index + 1), Quantity = index + 1,
            Values = new Dictionary<string, object> { ["color"] = color }
        }).ToList()
    };

    /// <summary>
    /// 创建指定 NPOI 格式的空模板。
    /// </summary>
    /// <remarks>
    /// 模板包含供导出定位的明细工作表。
    /// </remarks>
    /// <param name="xls">是否创建 HSSF 模板。</param>
    /// <returns>可读写模板流。</returns>
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
    /// 创建含明细命名起点的指定格式模板。
    /// </summary>
    /// <remarks>
    /// 命名区域从 B3 开始以验证物理坐标传递。
    /// </remarks>
    /// <param name="xls">是否创建 HSSF 模板。</param>
    /// <returns>可读写模板流。</returns>
    private static MemoryStream CreateNamedListTemplate(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var name = workbook.CreateName();
        name.NameName = "Lines";
        name.SheetIndex = workbook.GetSheetIndex(sheet);
        name.RefersToFormula = "'明细'!$B$3";
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建含外部工作簿旧名称的指定格式模板。
    /// </summary>
    /// <remarks>
    /// 外部链接关系和名称在序列化后仍可被 Provider 读取。
    /// </remarks>
    /// <param name="xls">是否创建 HSSF 模板。</param>
    /// <returns>可读写模板流。</returns>
    private static MemoryStream CreateExternalNameTemplate(bool xls)
    {
        var externalName = xls ? "External.xls" : "External.xlsx";
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        using IWorkbook externalWorkbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        externalWorkbook.CreateSheet("明细");
        workbook.CreateSheet("明细");
        if (workbook is XSSFWorkbook xssfWorkbook)
        {
            var externalLink = (NPOI.XSSF.Model.ExternalLinksTable)xssfWorkbook
                .CreateRelationship(XSSFRelation.EXTERNAL_LINKS, XSSFFactory.GetInstance());
            externalLink.CTExternalLink.AddNewExternalBook();
            externalLink.LinkedFileName = externalName;
            externalLink.CTExternalLink.externalBook.sheetNames = new CT_ExternalSheetNames
            {
                sheetName = new[]
                {
                    new CT_ExternalSheetName { val = externalWorkbook.GetSheetName(0) }
                }
            };
            var workbookXml = xssfWorkbook.GetCTWorkbook();
            var externalReferences = workbookXml.externalReferences ?? new CT_ExternalReferences();
            (externalReferences.externalReference ??= new List<CT_ExternalReference>()).Add(
                new CT_ExternalReference { id = xssfWorkbook.GetRelationId(externalLink) });
            workbookXml.externalReferences = externalReferences;
        }
        else
        {
            workbook.LinkExternalWorkbook(externalName, externalWorkbook);
        }
        var name = workbook.CreateName();
        name.NameName = "Sign_002";
        if (workbook is XSSFWorkbook)
            name.RefersToFormula = "'明细'!$Z$20";
        else
            name.RefersToFormula = "'[" + externalName + "]明细'!$Z$20";
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        if (workbook is XSSFWorkbook)
        {
            using var reloaded = new XSSFWorkbook(stream);
            Assert.Equal(externalName, Assert.Single(reloaded.ExternalLinksTable).LinkedFileName);
            reloaded.GetName("Sign_002").RefersToFormula = "'[" + externalName + "]明细'!$Z$20";
            var linkedStream = new MemoryStream();
            reloaded.Write(linkedStream, true);
            linkedStream.Position = 0;
            return linkedStream;
        }
        return stream;
    }

    /// <summary>
    /// 查找工作表中指定文本的第 N 个行号。
    /// </summary>
    /// <remarks>
    /// 返回值使用 NPOI 的零基行号。
    /// </remarks>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="text">待查找文本。</param>
    /// <param name="occurrence">从零开始的出现序号。</param>
    /// <returns>匹配行的零基行号。</returns>
    private static int FindRow(ISheet sheet, string text, int occurrence)
    {
        var rows = Enumerable.Range(0, sheet.LastRowNum + 1).Where(row =>
            sheet.GetRow(row)?.GetCell(0)?.CellType == CellType.String
            && sheet.GetRow(row).GetCell(0).StringCellValue == text).ToArray();
        return rows[occurrence];
    }

    /// <summary>
    /// 查找工作表第一列中指定文本的单元格。
    /// </summary>
    /// <remarks>
    /// 用于确认分页标记和最终 Footer 的原始内容。
    /// </remarks>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="text">待查找文本。</param>
    /// <returns>匹配的单元格。</returns>
    private static ICell FindFirstCell(ISheet sheet, string text) =>
        Enumerable.Range(0, sheet.LastRowNum + 1).Select(row => sheet.GetRow(row)?.GetCell(0))
            .First(cell => cell?.CellType == CellType.String && cell.StringCellValue == text);

    /// <summary>
    /// 断言名称指向明细表指定行的 A 列单元格。
    /// </summary>
    /// <remarks>
    /// 同时确认名称引用保持为单格范围。
    /// </remarks>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="nameText">待检查名称。</param>
    /// <param name="row">目标零基行号。</param>
    private static void AssertAnchor(IWorkbook workbook, string nameText, int row)
    {
        var name = Assert.Single(workbook.GetAllNames().Where(item => item.NameName == nameText));
        Assert.Equal("明细!$A$" + (row + 1), name.RefersToFormula.Replace("'明细'!", "明细!",
            StringComparison.Ordinal));
    }
}
