using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;
using PageDocument = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageLine;
using DynamicDocument = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.DynamicHeaderDocument;
using DynamicLine = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.DynamicHeaderLine;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML 独立页尾名称与长页序列、模板名称及物理列组合的职责。
/// </summary>
public sealed class EntityIndependentPageAnchorCompositionTest
{
    /// <summary>
    /// 验证 1001 个单条逻辑页的每个页尾名称、引用和明细往返。
    /// </summary>
    /// <remarks>
    /// 同时覆盖四位页码名称及完整成功数据。
    /// </remarks>
    [Fact]
    public void PageFooterNamed_1001SingleItemPages_ShouldKeepAllNamesAndRoundTrip()
    {
        var source = CreatePageDocument(1001);
        var layout = CreateNamedFooterLayout("Sign");
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            var names = GetNames(workbook, sheet).Where(name =>
                name.Name.StartsWith("Sign_", StringComparison.OrdinalIgnoreCase)).ToArray();
            Assert.Equal(1001, names.Length);
            foreach (var page in new[] { 999, 1000, 1001 })
                AssertNamedAnchor(workbook, sheet, "Sign_" + page.ToString("D3"),
                    FindRow(sheet, "SIGN", page - 1));

            for (var page = 1; page <= 1001; page++)
            {
                var footerRow = FindRow(sheet, "SIGN", page - 1);
                Assert.Equal(1d, sheet.Cell(footerRow, 2).GetDouble());
                Assert.Equal("L" + page, sheet.Cell(footerRow - 1, 2).GetString());
                AssertNamedAnchor(workbook, sheet, "Sign_" + page.ToString("D3"), footerRow);
            }
            Assert.Equal("END", sheet.Cell(sheet.LastRowUsed().RowNumber(), 1).GetString());
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => (line.Code, line.Amount)),
            result.Entity.Lines.Select(line => (line.Code, line.Amount)));
    }

    /// <summary>
    /// 验证第 1000 页首次超出名称上限时不写入目标。
    /// </summary>
    /// <remarks>
    /// 失败应保留已有目标内容。
    /// </remarks>
    [Fact]
    public void PageFooterNamed_LongPrefix_ShouldFailBeforeWritingAtPage1000()
    {
        var layout = CreateNamedFooterLayout(new string('P', 251));
        using var output = new MemoryStream();
        var original = new byte[] { 7, 4, 2 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreatePageDocument(1000), layout, output));
        Assert.Equal(BingOfficesStage.Plan, error.Stage);
        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 验证 RemoveObsolete 不清理外部旧名称且写入前保全模板与目标。
    /// </summary>
    /// <remarks>
    /// 同时检查模板流和目标流仍可继续使用。
    /// </remarks>
    [Fact]
    public void PageFooterNamed_RemoveObsolete_ExternalOldName_ShouldPreserveInputs()
    {
        using var template = CreateExternalNameTemplate();
        var templateBytes = template.ToArray();
        using var output = new MemoryStream();
        var original = new byte[] { 8, 6, 4, 2 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreatePageDocument(1),
                CreateNamedFooterLayout("Sign", ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, error.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证命名模板起点与动态列、计算列及校验的物理坐标组合。
    /// </summary>
    /// <remarks>
    /// 同时检查每页 Footer 计数、最终 Footer 计数和完整成功数据。
    /// </remarks>
    [Fact]
    public void ListRegionNamed_DynamicCalculatedAndPageFooter_ShouldPreservePhysicalCoordinates()
    {
        var source = CreateDynamicDocument();
        var layout = CreateDynamicLayout();
        using var template = CreateNamedListTemplate();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("扩展", sheet.Cell("D3").GetString());
            Assert.Equal("序号", sheet.Cell("E3").GetString());
            Assert.Equal("A", sheet.Cell("B4").GetString());
            Assert.Equal("L1", sheet.Cell("C4").GetString());
            Assert.Equal(1d, sheet.Cell("E4").GetDouble());
            Assert.Equal("XL1", sheet.Cell("D4").GetString());
            Assert.Equal("L2", sheet.Cell("C5").GetString());
            Assert.Equal("B", sheet.Cell("B5").GetString());
            Assert.Equal(2d, sheet.Cell("E5").GetDouble());
            Assert.Equal("XL2", sheet.Cell("D5").GetString());
            Assert.Equal("SIGN", sheet.Cell("B6").GetString());
            Assert.Equal(2d, sheet.Cell("C6").GetDouble());
            Assert.Equal("L3", sheet.Cell("C7").GetString());
            Assert.Equal("C", sheet.Cell("B7").GetString());
            Assert.Equal(3d, sheet.Cell("E7").GetDouble());
            Assert.Equal("XL3", sheet.Cell("D7").GetString());
            Assert.Equal("SIGN", sheet.Cell("B8").GetString());
            Assert.Equal(1d, sheet.Cell("C8").GetDouble());
            Assert.Equal("END", sheet.Cell("B9").GetString());
            Assert.Equal(3d, sheet.Cell("C9").GetDouble());
        }

        using (var input = new MemoryStream(output.ToArray(), writable: false))
        {
            var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
            Assert.Equal(new[] { "XL1", "XL2", "XL3" }, result.Entity.Lines.Select(line =>
                (string)line.Fields["extra"]));
        }

        foreach (var nativeFailure in new[] { true, false })
        {
            using var changed = new MemoryStream();
            using (var workbook = new XLWorkbook(
                new MemoryStream(output.ToArray(), writable: false)))
            {
                var sheet = workbook.Worksheet("明细");
                sheet.Cell("C7").Value = nativeFailure ? "非法" : "L3";
                sheet.Cell("D7").Value = "XL1";
                workbook.SaveAs(changed);
            }
            changed.Position = 0;
            var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
                ExcelImportValidationMode.ConfiguredAndWorkbook, ExcelUnsupportedFeaturePolicy.Fail);
            var invalid = new ClosedXmlExcelImporter().ImportEntity(changed, layout, options);
            var error = Assert.Single(invalid.Errors);
            Assert.Equal(nativeFailure ? ExcelImportErrorCode.WorkbookValidation
                : ExcelImportErrorCode.Validation, error.Code);
            Assert.Equal("明细", error.SheetName);
            Assert.Equal(7, error.RowIndex);
            Assert.Equal(nativeFailure ? 3 : 4, error.ColumnIndex);
            Assert.Equal(nativeFailure ? "Code" : "extra", error.ColumnKey);
            Assert.Equal(new[] { "L1", "L2" }, invalid.Entity.Lines.Select(line => line.Code));
            Assert.Equal(new[] { "A", "B" }, invalid.Entity.Lines.Select(line => line.Group));
            Assert.Equal(new[] { "XL1", "XL2" }, invalid.Entity.Lines.Select(line =>
                (string)line.Fields["extra"]));
        }
    }

    /// <summary>
    /// 验证原生 Code 失败后动态 Unique 预留回滚，使页尾后的下一条重复动态值仍可成功导入。
    /// </summary>
    [Fact]
    public void DynamicUnique_NativeFailure_ShouldRollbackReservationForNextRow()
    {
        var source = CreateDynamicDocument(new[] { "红", "蓝", "黄", "绿" });
        var layout = CreateDynamicLayout(7, new[] { "L1", "L2", "L3", "L4" });
        using var template = CreateNamedListTemplate();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("L3", sheet.Cell("C7").GetString());
            Assert.Equal("L4", sheet.Cell("C8").GetString());
            Assert.Equal("SIGN", sheet.Cell("B6").GetString());
            Assert.Equal(2d, sheet.Cell("C6").GetDouble());
            Assert.Equal("SIGN", sheet.Cell("B9").GetString());
            Assert.Equal(2d, sheet.Cell("C9").GetDouble());
            Assert.Equal("END", sheet.Cell("B10").GetString());
            Assert.Equal(4d, sheet.Cell("C10").GetDouble());
        }

        using var changed = new MemoryStream();
        using (var workbook = new XLWorkbook(
            new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            sheet.Cell("C7").Value = "非法";
            sheet.Cell("D7").Value = "重试";
            sheet.Cell("C8").Value = "L4";
            sheet.Cell("D8").Value = "重试";
            workbook.SaveAs(changed);
        }
        changed.Position = 0;
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredAndWorkbook, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new ClosedXmlExcelImporter().ImportEntity(changed, layout, options);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(7, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(new[] { "L1", "L2", "L4" }, result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(new[] { "红", "蓝", "重试" }, result.Entity.Lines.Select(line =>
            (string)line.Fields["extra"]));
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
    private static ExcelEntityLayout<DynamicDocument> CreateDynamicLayout(int validationEndRow = 6,
        IReadOnlyList<string> allowedCodes = null) =>
        ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegionNamed("明细", "Lines", document => document.Lines, region => region
                .DynamicColumnGroup("扩展字段组", line => line.Fields, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "extra", Title = "扩展", DataType = typeof(string),
                        Placement = ExcelColumnPlacement.After("Code"),
                        ValidationRules = new[]
                        {
                            new ExcelMappingDynamicValidationConfiguration
                            {
                                Name = "unique", IgnoreEmpty = false
                            }
                        }
                    }
                })
                .CalculatedColumn<int>("sequence", "序号", context => context.Index + 1,
                    column => column.Placement(ExcelColumnPlacement.After("extra")))
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
    /// <param name="count">明细数量。</param>
    /// <returns>分页明细实体。</returns>
    private static PageDocument CreatePageDocument(int count) => new()
    {
        Lines = Enumerable.Range(1, count).Select(index => new PageLine
            { Group = "G", Code = "L" + index, Amount = index }).ToList()
    };

    /// <summary>
    /// 创建动态列物理校验明细。
    /// </summary>
    /// <remarks>
    /// 每条明细包含用于 Unique 校验的扩展值。
    /// </remarks>
    /// <param name="extras">每条明细对应的动态扩展值。</param>
    /// <returns>动态列物理校验实体。</returns>
    private static DynamicDocument CreateDynamicDocument(IReadOnlyList<string> extras = null) => new()
    {
        Lines = (extras ?? new[] { "XL1", "XL2", "XL3" }).Select((extra, index) => new DynamicLine
        {
            Group = ((char)('A' + index)).ToString(), Code = "L" + (index + 1),
            Fields = new Dictionary<string, object> { ["extra"] = extra }
        }).ToList()
    };

    /// <summary>
    /// 创建含明细命名起点的模板。
    /// </summary>
    /// <remarks>
    /// 命名区域从 B3 开始以验证物理坐标传递。
    /// </remarks>
    /// <returns>可读写模板流。</returns>
    private static MemoryStream CreateNamedListTemplate()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        workbook.DefinedNames.Add("Lines", sheet.Range("B3"));
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建含外部工作簿旧名称的 ClosedXML 模板。
    /// </summary>
    /// <remarks>
    /// 外部名称用于验证 RemoveObsolete 的保全边界。
    /// </remarks>
    /// <returns>可读写模板流。</returns>
    private static MemoryStream CreateExternalNameTemplate()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("明细");
        workbook.DefinedNames.Add("Sign_002", "'[1]明细'!$A$20");
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 查找工作表中指定文本的第 N 个行号。
    /// </summary>
    /// <remarks>
    /// 返回值使用 ClosedXML 的一基行号。
    /// </remarks>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="text">待查找文本。</param>
    /// <param name="occurrence">从零开始的出现序号。</param>
    /// <returns>匹配行的一基行号。</returns>
    private static int FindRow(IXLWorksheet sheet, string text, int occurrence) =>
        Enumerable.Range(1, sheet.LastRowUsed()?.RowNumber() ?? 0)
            .Where(row => sheet.Cell(row, 1).GetString() == text).ElementAt(occurrence);

    /// <summary>
    /// 读取工作簿和工作表级的有效名称。
    /// </summary>
    /// <remarks>
    /// 合并结果用于检查名称数量和作用域。
    /// </remarks>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <returns>有效名称集合。</returns>
    private static IReadOnlyList<IXLDefinedName> GetNames(XLWorkbook workbook, IXLWorksheet sheet) =>
        workbook.DefinedNames.ValidNamedRanges()
            .Concat(sheet.DefinedNames.ValidNamedRanges()).ToArray();

    /// <summary>
    /// 断言名称指向明细表指定行的 A 列单元格。
    /// </summary>
    /// <remarks>
    /// 同时确认名称引用保持为单格范围。
    /// </remarks>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="nameText">待检查名称。</param>
    /// <param name="row">目标一基行号。</param>
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
