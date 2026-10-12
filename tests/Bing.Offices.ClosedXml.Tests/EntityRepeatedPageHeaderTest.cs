using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证实体列表实际重复页眉的导出与导入。
/// </summary>
public sealed class EntityRepeatedPageHeaderTest
{
    /// <summary>
    /// 验证多页模板的页眉复制、分页小计名称和明细读回。
    /// </summary>
    [Fact]
    public void PageHeader_Template_ShouldCopyContentStyleMergesAndRoundTrip()
    {
        var layout = CreateLayout();
        var source = CreateDocument("A", "A", "B", "B", "C");
        var bytes = Export(source, layout);

        using (var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("HEAD", sheet.Cell("A1").GetString());
            Assert.Equal("HEAD", sheet.Cell("A7").GetString());
            Assert.Equal("HEAD", sheet.Cell("A13").GetString());
            Assert.Equal("采购单", sheet.Cell("A2").GetString());
            Assert.Equal("采购单", sheet.Cell("A8").GetString());
            Assert.Equal("采购单", sheet.Cell("A14").GetString());
            Assert.Equal("仓库一", sheet.Cell("A3").GetString());
            Assert.Equal("仓库一", sheet.Cell("A9").GetString());
            Assert.Equal("第1/3页 数量2", sheet.Cell("C1").GetString());
            Assert.Equal("第2/3页 数量2", sheet.Cell("C7").GetString());
            Assert.Equal("第3/3页 数量1", sheet.Cell("C13").GetString());
            Assert.Equal("Group", sheet.Cell("B3").GetString());
            Assert.Equal("Group", sheet.Cell("B9").GetString());
            Assert.Equal("Group", sheet.Cell("B15").GetString());
            Assert.Equal("Code", sheet.Cell("C3").GetString());
            Assert.Equal("Code", sheet.Cell("C9").GetString());
            Assert.Equal("Code", sheet.Cell("C15").GetString());
            Assert.Equal(new[] { "A", "A", "B", "B", "C" },
                new[] { 4, 5, 10, 11, 16 }.Select(row => sheet.Cell(row, 2).GetString()));
            Assert.Equal(new[] { "L1", "L2", "L3", "L4", "L5" },
                new[] { 4, 5, 10, 11, 16 }.Select(row => sheet.Cell(row, 3).GetString()));
            Assert.Equal("PAGE", sheet.Cell("B6").GetString());
            Assert.Equal("PAGE", sheet.Cell("B12").GetString());
            Assert.Equal("END", sheet.Cell("B17").GetString());
            Assert.Equal(28d, sheet.Row(7).Height, 2);
            Assert.Equal(28d, sheet.Row(13).Height, 2);
            Assert.True(sheet.Cell("A8").Style.Font.Bold);
            Assert.True(sheet.Cell("A14").Style.Font.Bold);
            Assert.True(HasMerge(sheet, 8, 1, 8, 3));
            Assert.True(HasMerge(sheet, 14, 1, 14, 3));
            Assert.Equal(new[] { 6, 12 }, sheet.PageSetup.RowBreaks);
            Assert.Equal(new[] { 6, 12 }, workbook.DefinedNames.ValidNamedRanges()
                .Where(name => name.Name.StartsWith("PageTotal_", StringComparison.Ordinal))
                .OrderBy(name => name.Name)
                .Select(name => Assert.Single(name.Ranges).RangeAddress.FirstAddress.RowNumber));
        }

        using var input = new MemoryStream(bytes, writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证空明细和单页明细仅保留首 页页眉。
    /// </summary>
    /// <param name="count">明细行数。</param>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void PageHeader_EmptyOrSinglePage_ShouldNotDuplicateHeader(int count)
    {
        var source = CreateDocument(Enumerable.Repeat("A", count).ToArray());
        var bytes = Export(source, CreateLayout());
        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        var sheet = workbook.Worksheet("明细");

        Assert.Equal("HEAD", sheet.Cell("A1").GetString());
        Assert.Equal($"第1/1页 数量{count}", sheet.Cell("C1").GetString());
        Assert.Equal("Group", sheet.Cell("B3").GetString());
        Assert.Equal("Code", sheet.Cell("C3").GetString());
        Assert.Equal("END", sheet.Cell(4 + count, 2).GetString());
        Assert.Empty(sheet.PageSetup.RowBreaks);
        Assert.Empty(workbook.DefinedNames.ValidNamedRanges().Where(name =>
            name.Name.StartsWith("PageTotal_", StringComparison.Ordinal)));
        Assert.True(sheet.Cell(5 + count, 1).IsEmpty());

        using var input = new MemoryStream(bytes, writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, CreateLayout());
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(count, result.Entity.Lines.Count);
    }

    /// <summary>
    /// 验证分组小计与分页小计在同一边界时，页眉仍紧随分页符。
    /// </summary>
    [Fact]
    public void PageHeader_WithGroupSubtotal_ShouldPreserveSegmentOrder()
    {
        var layout = CreateLayout(grouped: true);
        var bytes = Export(CreateDocument("A", "A", "B"), layout);
        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        var sheet = workbook.Worksheet("明细");
        Assert.Equal("GROUP", sheet.Cell("B6").GetString());
        Assert.Equal("PAGE", sheet.Cell("B7").GetString());
        Assert.Equal("HEAD", sheet.Cell("A8").GetString());
        Assert.Equal("B", sheet.Cell("B11").GetString());
        Assert.Equal("L3", sheet.Cell("C11").GetString());
        Assert.Equal("GROUP", sheet.Cell("B12").GetString());
        Assert.Equal("END", sheet.Cell("B13").GetString());
        Assert.Equal(new[] { 7 }, sheet.PageSetup.RowBreaks);

        using var input = new MemoryStream(bytes, writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证缺失或错位的复制页眉标记被识别为结构化配置错误。
    /// </summary>
    /// <param name="move">是否把标记移动到错误行。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_MissingOrShiftedMarker_ShouldRejectImport(bool move)
    {
        var layout = CreateLayout();
        var bytes = Export(CreateDocument("A", "A", "B"), layout);
        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        var sheet = workbook.Worksheet("明细");
        sheet.Cell("A7").Value = "INVALID";
        if (move)
            sheet.Cell("A8").Value = "HEAD";
        using var corrupted = new MemoryStream();
        workbook.SaveAs(corrupted);
        corrupted.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(corrupted, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(move ? 8 : 7, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>验证首页 marker 缺失及重复 marker 都返回原始坐标。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_FirstOrDuplicateMarker_ShouldRejectImport(bool duplicate)
    {
        var layout = CreateLayout();
        var bytes = Export(CreateDocument("A", "A", "B"), layout);
        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        var sheet = workbook.Worksheet("明细");
        if (duplicate)
            sheet.Cell("A2").Value = "HEAD";
        else
            sheet.Cell("A1").Value = "BROKEN";
        using var corrupted = new MemoryStream();
        workbook.SaveAs(corrupted);
        corrupted.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(corrupted, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(duplicate ? 2 : 1, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>验证布局构建时拒绝 marker 与固定字段冲突。</summary>
    [Fact]
    public void PageHeader_MarkerConflictsWithFixedCell_ShouldRejectLayout()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<HeaderDocument>(builder => builder
            .Cell("明细", "A1", document => document.Title)
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD"))));
    }

    /// <summary>验证源页眉未覆盖完整列表表头时预检失败。</summary>
    [Fact]
    public void PageHeader_SourceRangeMissingHeaderColumn_ShouldRejectBeforeOutput()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:B3", "A1", "HEAD")));
        using var template = CreateTemplate();
        using var output = new MemoryStream();

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>验证模板打印标题与实际页眉互斥。</summary>
    [Fact]
    public void PageHeader_TemplatePrintTitles_ShouldRejectBeforeOutput()
    {
        using var template = CreateTemplate(printTitles: true);
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>验证 marker 位于模板合并区域内部时预检失败并保留流所有权。</summary>
    [Fact]
    public void PageHeader_MarkerInsideMerge_ShouldPreserveOutputAndTemplate()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Range("A1:B1").Merge();
        using var template = new MemoryStream();
        workbook.SaveAs(template);
        template.Position = 0;
        using var output = new MemoryStream();
        var original = new byte[] { 5, 4, 3 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "B1", "HEAD")));

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>验证 Fluent 页眉合并与模板合并冲突时预检失败并保留流所有权。</summary>
    [Fact]
    public void PageHeader_FluentMergeConflictsWithTemplate_ShouldPreserveOutputAndTemplate()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD", header => header
                    .Merge("A2:B2"))));
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        var original = new byte[] { 5, 4, 3 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>验证同表无重叠的普通列表与实际页眉列表可以往返。</summary>
    [Fact]
    public void PageHeader_WithNonOverlappingOrdinaryList_ShouldRoundTrip()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD").End("D9"))
            .ListRegion("明细", "J3", document => document.OtherLines,
                region => region.End("K4")));
        var source = CreateDocument("A", "B", "C");
        source.OtherLines.Add(new HeaderLine { Code = "R1" });
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
            Assert.Equal("HEAD", workbook.Worksheet("明细").Cell("A1").GetString());
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(source.OtherLines.Select(line => line.Code), result.Entity.OtherLines.Select(line => line.Code));
    }

    /// <summary>验证未限定结束行的页眉列表不读取相邻列表的后续行。</summary>
    [Fact]
    public void PageHeader_WithoutEnd_WithLongerAdjacentList_ShouldRoundTrip()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD"))
            .ListRegion("明细", "J3", document => document.OtherLines,
                region => region.End("K12")));
        var source = CreateDocument("A", "B", "C");
        for (var index = 0; index < 8; index++)
            source.OtherLines.Add(new HeaderLine { Code = $"R{index}" });
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(source.OtherLines.Select(line => line.Code), result.Entity.OtherLines.Select(line => line.Code));
    }

    /// <summary>验证页眉值工厂失败时调用方目标流保持原样。</summary>
    [Fact]
    public void PageHeader_ValueFactoryFailure_ShouldPreserveTarget()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD", header => header
                    .Cell<string>("C1", (ExcelEntityPageHeaderContext<HeaderLine> _) =>
                        throw new InvalidOperationException("页眉值计算失败")))));
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        var original = new byte[] { 2, 4, 6 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesExportException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证无模板时可以由 Fluent 声明页眉内容和合并。
    /// </summary>
    [Fact]
    public void PageHeader_WithoutTemplate_ShouldWriteFluentHeader()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2)
                .PageHeader("A1:C3", "A1", "HEAD", header => header
                    .Cell("A2", "采购单")
                    .Cell("C1", page => $"第{page.PageNumber}页")
                    .Merge("A2:C2"))));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "A", "B"), layout, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.Worksheet("明细");

        Assert.Equal("采购单", sheet.Cell("A2").GetString());
        Assert.Equal("采购单", sheet.Cell("A7").GetString());
        Assert.Equal("第2页", sheet.Cell("C6").GetString());
        Assert.True(HasMerge(sheet, 7, 1, 7, 3));
    }

    /// <summary>
    /// 验证预取消时目标流不变且模板流仍由调用方持有。
    /// </summary>
    [Fact]
    public async Task PageHeader_AsyncPreCanceled_ShouldPreserveOutputAndTemplate()
    {
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        var original = new byte[] { 3, 1, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplateAsync(CreateDocument("A", "A", "B"),
                CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output,
                cancellation.Token));
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>验证后续页眉计算期间取消时不写入调用方流。</summary>
    [Fact]
    public async Task PageHeader_MidWriteCancellation_ShouldPreserveOutput()
    {
        using var cancellation = new CancellationTokenSource();
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD", header => header
                    .Cell("C1", (ExcelEntityPageHeaderContext<HeaderLine> context) =>
                    {
                        if (context.PageNumber == 2)
                            cancellation.Cancel();
                        return context.PageNumber;
                    }))));
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        var original = new byte[] { 3, 1, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplateAsync(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output,
                cancellation.Token));
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>验证复制页眉超出 XLSX 最大行数时预检失败。</summary>
    [Fact]
    public void PageHeader_MaxRow_ShouldRejectBeforeOutput()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B1048574", document => document.Lines, region => region
                .PageBreak(1).PageHeader("A1048573:C1048574", "A1048573", "HEAD")));
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B"), layout, output));
        Assert.Equal(0, output.Length);
    }

    /// <summary>验证无模板时固定字段占用的重复页眉目标不可覆盖。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_WithoutTemplateFixedCellOverlap_ShouldRejectBeforeOutput(bool emptyValue)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .Cell("明细", "A5", item => item.Title)
            .ListRegion("明细", "B2", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C2", "A1", "HEAD")));
        var source = CreateDocument("A", "B", "C");
        source.Title = emptyValue ? null : "BUSINESS";
        using var output = new MemoryStream();
        var original = new byte[] { 7, 2, 9 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(source, layout, output));
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>验证命名列表起点解析后仍须紧邻首页页眉。</summary>
    [Fact]
    public void PageHeader_NamedListWithoutHeader_ShouldValidateResolvedAdjacency()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegionNamed("明细", "Lines", item => item.Lines, region => region
                .Header(false).PageBreak(2).PageHeader("A1:C2", "A1", "HEAD")));
        using (var validWorkbook = new XLWorkbook())
        {
            var validSheet = validWorkbook.Worksheets.Add("明细");
            validSheet.Cell("A1").Value = "HEAD";
            validWorkbook.DefinedNames.Add("Lines", validSheet.Range("B3"));
            using var template = new MemoryStream();
            validWorkbook.SaveAs(template);
            template.Position = 0;
            using var output = new MemoryStream();
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            using var validInput = new MemoryStream(output.ToArray(), writable: false);
            var result = new ClosedXmlExcelImporter().ImportEntity(validInput, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
        }
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Cell("A1").Value = "HEAD";
        sheet.Cell("B3").Value = "A";
        sheet.Cell("C3").Value = "L1";
        workbook.DefinedNames.Add("Lines", sheet.Range("B4"));
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(input, layout));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
    }

    /// <summary>验证 End 预留空行不算作明细，页眉缺失时报告原始坐标。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_EndReservedRows_ShouldRoundTripOrReportMissingMarker(bool breakMarker)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD").End("D30")));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"), layout, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        if (breakMarker)
            workbook.Worksheet("明细").Cell("A6").Value = "BROKEN";
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;

        if (breakMarker)
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(input, layout));
            Assert.Equal(6, exception.RowIndex);
            Assert.Equal(1, exception.ColumnIndex);
        }
        else
        {
            var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
        }
    }

    /// <summary>验证无分页小计时缺失的页眉标记指向预期单元格。</summary>
    [Fact]
    public void PageHeader_WithoutSubtotalMissingMarker_ShouldReportExpectedCell()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"), layout, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        workbook.Worksheet("明细").Cell("A6").Value = "BROKEN";
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(input, layout));
        Assert.Equal(6, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>验证分组边界后的非首行标记缺失时仍报告实际位置。</summary>
    [Fact]
    public void PageHeader_GroupBoundaryMissingOffsetMarker_ShouldReportExpectedCell()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).GroupSubtotal(line => line.Group, "GROUP")
                .PageHeader("A1:C3", "A2", "HEAD")));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "A", "B"), layout, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.Worksheet("明细");
        var copiedMarker = Assert.Single(sheet.CellsUsed().Where(cell =>
            cell.Address.RowNumber > 3 && cell.Address.ColumnNumber == 1
            && cell.GetString() == "HEAD"));
        var expectedRow = copiedMarker.Address.RowNumber;
        copiedMarker.Value = "BROKEN";
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(input, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(expectedRow, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>验证异步导出和导入保持相同的页眉与明细结构。</summary>
    [Fact]
    public async Task PageHeader_AsyncRoundTrip_ShouldPreserveHeaderAndDetails()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportEntityAsync(CreateDocument("A", "B", "C"),
            layout, output);
        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
            Assert.Equal("HEAD", workbook.Worksheet("明细").Cell("A6").GetString());
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = await new ClosedXmlExcelImporter().ImportEntityAsync(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>验证动态列与计算列的最终表头范围参与重复页眉校验。</summary>
    [Fact]
    public void PageHeader_DynamicAndCalculatedColumns_ShouldUsePhysicalWidth()
    {
        var layout = ExcelEntity.Layout<DynamicHeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .DynamicColumnGroup("扩展字段组", item => item.Fields, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "extra", Title = "扩展", DataType = typeof(string),
                        Placement = ExcelColumnPlacement.After("Code")
                    }
                })
                .CalculatedColumn<int>("sequence", "序号", context => context.Index + 1,
                    column => column.Placement(ExcelColumnPlacement.After("extra")))
                .PageBreak(2).PageHeader("A1:E3", "A1", "HEAD")));
        var source = new DynamicHeaderDocument
        {
            Lines = new List<DynamicHeaderLine>
            {
                new() { Group = "A", Code = "L1", Fields = new Dictionary<string, object> { ["extra"] = "XL1" } },
                new() { Group = "B", Code = "L2", Fields = new Dictionary<string, object> { ["extra"] = "XL2" } },
                new() { Group = "C", Code = "L3", Fields = new Dictionary<string, object> { ["extra"] = "XL3" } }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("扩展", sheet.Cell("D3").GetString());
            Assert.Equal("序号", sheet.Cell("E3").GetString());
            Assert.Equal("扩展", sheet.Cell("D8").GetString());
            Assert.Equal("序号", sheet.Cell("E8").GetString());
            Assert.Equal("XL1", sheet.Cell("D4").GetString());
            Assert.Equal(3, sheet.Cell("E9").GetDouble());
        }
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
        Assert.Equal("XL3", result.Entity.Lines[2].Fields["extra"]);
    }

    /// <summary>验证页眉源高度不合法时布局阶段拒绝。</summary>
    [Theory]
    [InlineData("A1:C1")]
    [InlineData("A1:C2")]
    [InlineData("A1:C4")]
    public void PageHeader_InvalidHeight_ShouldRejectLayout(string sourceRange)
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader(sourceRange, "A1", "HEAD"))));
    }

    /// <summary>验证页眉文件提交失败只观察一次并保留目标文件。</summary>
    [Fact]
    public async Task PageHeader_FileCommitFailure_ShouldPreserveTargetAndObserveOnce()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-page-header-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, original);
        var observer = new HeaderFailureObserver();
        var failure = new BingOfficesFileCommitException("commit failed", provider: "ClosedXML");
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer },
            fileExportCommitter: new HeaderFailureCommitter(failure));
        try
        {
            var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                exporter.ExportEntityToFileAsync(CreateDocument("A", "B", "C"),
                    ExcelEntity.Layout<HeaderDocument>(builder => builder
                        .ListRegion("明细", "B3", item => item.Lines, region => region
                            .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD"))), path));
            Assert.Same(failure, exception);
            Assert.Equal(1, observer.Count);
            Assert.Same(exception, observer.LastException);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "entity.xlsx.*.tmp"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>验证页眉计算失败时默认提交器清理暂存文件并保留原目标。</summary>
    [Fact]
    public async Task PageHeader_FileWriteFailure_ShouldCleanStagingAndObserveOnce()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-page-header-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, original);
        var observer = new HeaderFailureObserver();
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD", header => header
                    .Cell("C1", (ExcelEntityPageHeaderContext<HeaderLine> page) =>
                        page.PageNumber == 2 ? throw new InvalidOperationException("header failed")
                            : "ok"))));
        try
        {
            var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
                new ClosedXmlExcelExporter(exceptionObservers: new[] { observer })
                    .ExportEntityToFileAsync(CreateDocument("A", "B", "C"), layout, path));
            Assert.Equal(1, observer.Count);
            Assert.Same(exception, observer.LastException);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "entity.xlsx.*.tmp"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>验证明细标记文本冲突不会改写调用方目标流。</summary>
    [Fact]
    public void PageHeader_DetailMarkerCollision_ShouldPreserveTarget()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "B1", "HEAD")));
        using var output = new MemoryStream();
        var original = new byte[] { 1, 2, 3 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument("HEAD", "B", "C"),
                layout, output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(4, exception.RowIndex);
        Assert.Equal(2, exception.ColumnIndex);
        Assert.Equal(original, output.ToArray());
    }

    /// <summary>验证动态列和计算列写出的标记值同样被拒绝。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_DynamicOrCalculatedMarkerCollision_ShouldReject(bool calculated)
    {
        var layout = ExcelEntity.Layout<DynamicHeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region =>
            {
                if (calculated)
                    region.CalculatedColumn<string>("marker", "标记", _ => "HEAD",
                        column => column.Placement(ExcelColumnPlacement.After("Code")));
                else
                    region.DynamicColumnGroup("扩展字段组", item => item.Fields, new[]
                    {
                        new ExcelDynamicColumnDefinition
                        {
                            Key = "extra", Title = "扩展", DataType = typeof(string),
                            Placement = ExcelColumnPlacement.After("Code")
                        }
                    });
                region.PageBreak(2).PageHeader(calculated ? "A1:E3" : "A1:D3",
                    "D1", "HEAD");
            }));
        var source = new DynamicHeaderDocument
        {
            Lines = new List<DynamicHeaderLine>
            {
                new() { Group = "A", Code = "L1",
                    Fields = new Dictionary<string, object> { ["extra"] = " HEAD " } }
            }
        };
        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(source, layout, output));
        Assert.True(exception.RowIndex == 4, exception.Message);
        Assert.True(exception.ColumnIndex == 4, exception.Message);
        Assert.Empty(output.ToArray());
    }

    /// <summary>验证 Fluent 单元格不得改写列表表头标题。</summary>
    [Fact]
    public void PageHeader_FluentTitleOverride_ShouldRejectButMatchingTitleRoundTrips()
    {
        ExcelEntityLayout<HeaderDocument> Layout(string title) =>
            ExcelEntity.Layout<HeaderDocument>(builder => builder
                .ListRegion("明细", "B3", document => document.Lines, region => region
                    .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD", header => header
                        .Cell("B3", title))));
        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"),
                Layout("错误标题"), output));
        Assert.Equal(3, exception.RowIndex);
        Assert.Equal(2, exception.ColumnIndex);
        Assert.Empty(output.ToArray());

        var layout = Layout("Group");
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"), layout,
            output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>验证 Fluent 合并不能覆盖表头且不影响单据头合并。</summary>
    [Fact]
    public void PageHeader_FluentMerge_ShouldRejectHeaderAndAllowDocumentTitle()
    {
        ExcelEntityLayout<HeaderDocument> Layout(string range) =>
            ExcelEntity.Layout<HeaderDocument>(builder => builder
                .ListRegion("明细", "B3", document => document.Lines, region => region
                    .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD", header => header
                        .Merge(range))));
        using var rejected = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"),
                Layout("B3:C3"), rejected));
        Assert.Empty(rejected.ToArray());

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"),
            Layout("A2:C2"), output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.Worksheet("明细");
        Assert.True(HasMerge(sheet, 2, 1, 2, 3));
        Assert.True(HasMerge(sheet, 7, 1, 7, 3));
    }

    /// <summary>验证导入拒绝已篡改的复制页表头。</summary>
    [Fact]
    public void PageHeader_CopiedTitleTampered_ShouldReportOriginalCell()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"),
            layout, exported);
        using var workbook = new XLWorkbook(new MemoryStream(exported.ToArray(), writable: false));
        workbook.Worksheet("明细").Cell("B8").Value = "错误标题";
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(input, layout));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(8, exception.RowIndex);
        Assert.Equal(2, exception.ColumnIndex);
    }

    /// <summary>验证无表头列表的标记列必须避开明细列。</summary>
    [Fact]
    public void PageHeader_WithoutListHeader_ShouldRequireDedicatedMarkerColumn()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B2", document => document.Lines, region => region
                .Header(false).PageBreak(2).PageHeader("A1:C1", "B1", "HEAD")));
        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"),
                layout, output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(1, exception.RowIndex);
        Assert.Equal(2, exception.ColumnIndex);
        Assert.Empty(output.ToArray());

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Cell("B1").Value = "HEAD";
        sheet.Cell("B2").Value = "A";
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;
        var importError = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(input, layout));
        Assert.Equal(BingOfficesStage.Plan, importError.Stage);
        Assert.Equal(1, importError.RowIndex);
        Assert.Equal(2, importError.ColumnIndex);
    }

    /// <summary>验证多页模板不会静默丢失源页眉元数据。</summary>
    [Theory]
    [InlineData("hyperlink")]
    [InlineData("comment")]
    [InlineData("rich-text")]
    [InlineData("validation")]
    [InlineData("conditional-format")]
    public void PageHeader_SourceMetadata_ShouldRejectBeforeOutput(string metadata)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        var cell = sheet.Cell("A2");
        switch (metadata)
        {
            case "hyperlink":
                cell.Value = "链接";
                cell.SetHyperlink(new XLHyperlink(new Uri("https://example.com")));
                break;
            case "comment":
                cell.CreateComment().AddText("批注");
                break;
            case "rich-text":
                cell.CreateRichText().AddText("富文本").SetBold();
                break;
            case "validation":
                sheet.Range("A2").CreateDataValidation().WholeNumber.Between(1, 10);
                break;
            case "conditional-format":
                sheet.Range("A2").AddConditionalFormat().WhenGreaterThan(5)
                    .Fill.SetBackgroundColor(XLColor.Yellow);
                break;
        }
        using var template = new MemoryStream();
        workbook.SaveAs(template);
        template.Position = 0;
        using var output = new MemoryStream();
        var original = new byte[] { 3, 2, 1 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
                layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>验证页眉隐藏状态和锁定样式随页复制。</summary>
    [Fact]
    public void PageHeader_HiddenRowAndUnlockedCell_ShouldCopyToNextPage()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Cell("A2").Value = "采购单";
        sheet.Cell("A2").Style.Protection.Locked = false;
        sheet.Row(2).Hide();
        using var template = new MemoryStream();
        workbook.SaveAs(template);
        template.Position = 0;
        using var output = new MemoryStream();
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));

        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "B", "C"),
            layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var actual = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var result = actual.Worksheet("明细");
        Assert.True(result.Row(2).IsHidden);
        Assert.True(result.Row(7).IsHidden);
        Assert.Equal("采购单", result.Cell("A7").GetString());
        Assert.False(result.Cell("A7").Style.Protection.Locked);
    }

    /// <summary>验证单页模板原有元数据不受复制边界限制。</summary>
    [Fact]
    public void PageHeader_SinglePage_ShouldPreserveTemplateMetadata()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Cell("A2").Value = "链接";
        sheet.Cell("A2").SetHyperlink(new XLHyperlink(new Uri("https://example.com")));
        sheet.Cell("A2").CreateComment().AddText("批注");
        sheet.Row(2).Hide();
        using var template = new MemoryStream();
        workbook.SaveAs(template);
        template.Position = 0;
        using var output = new MemoryStream();
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));

        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A"), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var actual = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var result = actual.Worksheet("明细");
        Assert.True(result.Cell("A2").HasHyperlink);
        Assert.True(result.Cell("A2").HasComment);
        Assert.True(result.Row(2).IsHidden);
    }

    /// <summary>验证配置式校验覆盖源页眉时在写出前拒绝多页。</summary>
    [Fact]
    public void PageHeader_ConfiguredValidationOnSource_ShouldRejectMultiPage()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .DataValidation("明细", new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                    { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
                Type = ExcelDataValidationType.List,
                Values = new[] { "A", "B" }
            })
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));
        using var output = new MemoryStream();
        var original = new byte[] { 2, 4, 6 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "B", "C"),
                layout, output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal(original, output.ToArray());
    }

    /// <summary>验证稀疏大工作表中的重复页眉列表只读取自身占用行。</summary>
    [Fact]
    public void PageHeader_SparseLargeEnd_ShouldIgnoreUnrelatedDistantCell()
    {
        const int distantRow = 20000;
        var layout = CreateSparseLargeEndLayout(distantRow);
        using var input = CreateSparseLargeEndWorkbook(distantRow, secondDetail: false);

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>验证稀疏大工作表的资源限制仍在实体读取时生效。</summary>
    [Fact]
    public void PageHeader_SparseLargeEnd_ShouldRespectRowLimit()
    {
        const int distantRow = 20000;
        var layout = CreateSparseLargeEndLayout(distantRow);
        using var input = CreateSparseLargeEndWorkbook(distantRow, secondDetail: true);
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxRows = 1 });

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit
            && error.RowIndex == 5);
    }

    /// <summary>验证长跨度页眉标记扫描在读取途中响应取消。</summary>
    [Fact]
    public void PageHeader_LongMarkerScan_ShouldObserveMidScanCancellation()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2000).PageHeader("A1:C3", "A1", "HEAD")));
        using var cancellation = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() =>
            ExcelEntitySubtotalImportScanner.FindSpans(layout.ListRegions[0], 3, 1000,
                row =>
                {
                    if (row == 300)
                        cancellation.Cancel();
                    return null;
                }, row => row == 0 ? "HEAD" : null, cancellation.Token));
    }

    /// <summary>验证复制页表头校验在读取途中响应取消。</summary>
    [Fact]
    public void PageHeader_CopiedTitleCheck_ShouldObserveMidCheckCancellation()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")));
        using var cancellation = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() =>
            ExcelEntityPageHeaderImportValidation.ValidateCopiedHeaders(
                layout.ListRegions[0], new[] { (StartRow: 6, EndRow: 8) }, 2,
                _ => "title", (_, _) =>
                {
                    cancellation.Cancel();
                    return "HEAD";
                }, cancellation.Token));
    }

    private static ExcelEntityLayout<HeaderDocument> CreateSparseLargeEndLayout(int distantRow) =>
        ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", document => document.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEAD")
                .End($"C{distantRow}")));

    private static MemoryStream CreateSparseLargeEndWorkbook(int distantRow,
        bool secondDetail)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        sheet.Cell("A1").Value = "HEAD";
        sheet.Cell("B3").Value = "Group";
        sheet.Cell("C3").Value = "Code";
        sheet.Cell("B4").Value = "A";
        sheet.Cell("C4").Value = "L1";
        if (secondDetail)
        {
            sheet.Cell("B5").Value = "B";
            sheet.Cell("C5").Value = "L2";
        }
        sheet.Cell(distantRow, 26).Value = "unrelated";
        var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;
        return output;
    }

    private static ExcelEntityLayout<HeaderDocument> CreateLayout(bool grouped = false) =>
        ExcelEntity.Layout<HeaderDocument>(builder => builder.ListRegion("明细", "B3",
            document => document.Lines, region =>
            {
                region.PageBreak(2)
                    .PageHeader("A1:C3", "A1", "HEAD", header => header
                        .Cell("C1", page =>
                            $"第{page.PageNumber}/{page.PageCount}页 数量{page.Items.Count}"))
                    .PageSubtotalNamed("PageTotal", "PAGE")
                    .Footer("END");
                if (grouped)
                    region.GroupSubtotal(line => line.Group, "GROUP");
            }));

    private static HeaderDocument CreateDocument(params string[] groups) => new()
    {
        Lines = groups.Select((group, index) => new HeaderLine
        {
            Group = group,
            Code = "L" + (index + 1)
        }).ToList()
    };

    private static byte[] Export(HeaderDocument source, ExcelEntityLayout<HeaderDocument> layout)
    {
        using var template = CreateTemplate();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        return output.ToArray();
    }

    private static MemoryStream CreateTemplate(bool printTitles = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("明细");
        if (printTitles)
            sheet.PageSetup.SetRowsToRepeatAtTop("1:3");
        sheet.Cell("A1").Value = "模板标记";
        sheet.Cell("A2").Value = "采购单";
        sheet.Cell("A2").Style.Font.Bold = true;
        sheet.Range("A2:C2").Merge();
        sheet.Cell("A3").Value = "仓库一";
        sheet.Row(1).Height = 28;
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    private static bool HasMerge(IXLWorksheet sheet, int firstRow, int firstColumn,
        int lastRow, int lastColumn) => sheet.MergedRanges.Any(range =>
        range.RangeAddress.FirstAddress.RowNumber == firstRow
        && range.RangeAddress.FirstAddress.ColumnNumber == firstColumn
        && range.RangeAddress.LastAddress.RowNumber == lastRow
        && range.RangeAddress.LastAddress.ColumnNumber == lastColumn);

    /// <summary>
    /// 保存页眉测试单据。
    /// </summary>
    public sealed class HeaderDocument
    {
        /// <summary>获取或设置单据标题。</summary>
        public string Title { get; set; }

        /// <summary>获取或设置明细。</summary>
        public List<HeaderLine> Lines { get; set; } = new();

        /// <summary>获取或设置并列表格行。</summary>
        public List<HeaderLine> OtherLines { get; set; } = new();
    }

    /// <summary>
    /// 保存页眉测试明细。
    /// </summary>
    public sealed class HeaderLine
    {
        /// <summary>获取或设置分组键。</summary>
        public string Group { get; set; }

        /// <summary>获取或设置行编码。</summary>
        public string Code { get; set; }

    }

    /// <summary>保存动态列表测试单据。</summary>
    public sealed class DynamicHeaderDocument
    {
        /// <summary>获取或设置明细。</summary>
        public List<DynamicHeaderLine> Lines { get; set; } = new();
    }

    /// <summary>保存动态列表测试明细。</summary>
    public sealed class DynamicHeaderLine
    {
        /// <summary>获取或设置分组。</summary>
        public string Group { get; set; }

        /// <summary>获取或设置编码。</summary>
        public string Code { get; set; }

        /// <summary>获取或设置扩展字段。</summary>
        public IDictionary<string, object> Fields { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>记录页眉导出异常的观察次数。</summary>
    private sealed class HeaderFailureObserver : IBingOfficesExceptionObserver
    {
        /// <summary>获取观察次数。</summary>
        public int Count { get; private set; }

        /// <summary>获取最近一次异常。</summary>
        public BingOfficesException LastException { get; private set; }

        /// <inheritdoc />
        public void Observe(BingOfficesException exception)
        {
            Count++;
            LastException = exception;
        }
    }

    /// <summary>模拟页眉导出文件提交失败。</summary>
    private sealed class HeaderFailureCommitter : IFileExportCommitter
    {
        /// <summary>保存提交异常。</summary>
        private readonly BingOfficesFileCommitException _failure;

        /// <summary>初始化一个 <see cref="HeaderFailureCommitter" /> 类型的实例。</summary>
        public HeaderFailureCommitter(BingOfficesFileCommitException failure) => _failure = failure;

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format) => throw _failure;

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format) => Task.FromException(_failure);
    }
}
