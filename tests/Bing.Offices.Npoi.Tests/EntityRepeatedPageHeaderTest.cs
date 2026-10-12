using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Entities;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证实体列表实际重复页眉的 NPOI 行为。
/// </summary>
public sealed class EntityRepeatedPageHeaderTest
{
    /// <summary>
    /// 验证模板页眉、分页小计和导入往返。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_Template_ShouldCopyAndRoundTrip(bool xls)
    {
        var layout = CreateLayout();
        var source = CreateDocument(5);
        var bytes = Export(source, layout, xls);
        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("HEADER", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal("HEADER", sheet.GetRow(6).GetCell(0).StringCellValue);
            Assert.Equal("HEADER", sheet.GetRow(12).GetCell(0).StringCellValue);
            Assert.Equal("1/3", sheet.GetRow(0).GetCell(2).StringCellValue);
            Assert.Equal("2/3", sheet.GetRow(6).GetCell(2).StringCellValue);
            Assert.Equal("3/3", sheet.GetRow(12).GetCell(2).StringCellValue);
            Assert.Equal(sheet.GetRow(0).Height, sheet.GetRow(6).Height);
            Assert.Equal("1+1", sheet.GetRow(7).GetCell(0).CellFormula);
            Assert.Equal(sheet.GetRow(1).GetCell(1).CellStyle.Index,
                sheet.GetRow(7).GetCell(1).CellStyle.Index);
            Assert.Contains(sheet.MergedRegions, merge => merge.FirstRow == 7
                && merge.LastRow == 7 && merge.FirstColumn == 1 && merge.LastColumn == 2);
            Assert.Equal(new[] { 5, 11 }, sheet.RowBreaks);
            Assert.Equal("PAGE", sheet.GetRow(5).GetCell(1).StringCellValue);
            Assert.Equal("END", sheet.GetRow(16).GetCell(1).StringCellValue);
            Assert.Equal(2, workbook.GetAllNames().Count(name => name.NameName.StartsWith(
                "PageTotal_", StringComparison.OrdinalIgnoreCase)));
        }
        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => item.Code),
            result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证无明细或单页时不复制页眉。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_EmptyOrSinglePage_ShouldNotCopy(bool xls)
    {
        foreach (var count in new[] { 0, 1 })
        {
            var bytes = Export(CreateDocument(count), CreateLayout(), xls);
            using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
            var sheet = workbook.GetSheet("明细");
            Assert.Null(sheet.GetRow(6));
            Assert.Empty(sheet.RowBreaks);
        }
    }

    /// <summary>
    /// 验证重复页眉 marker 丢失时拒绝导入。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_MissingMarker_ShouldRejectImport(bool xls)
    {
        var layout = CreateLayout();
        var bytes = Export(CreateDocument(3), layout, xls);
        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        workbook.GetSheet("明细").GetRow(6).GetCell(0).SetCellValue("BROKEN");
        using var changed = new MemoryStream();
        workbook.Write(changed, true);
        changed.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(7, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>验证首页页眉标记和额外重复标记均导致结构化导入错误。</summary>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void PageHeader_FirstOrDuplicateMarker_ShouldRejectImport(bool xls, bool duplicate)
    {
        var layout = CreateLayout();
        var bytes = Export(CreateDocument(3), layout, xls);
        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        var sheet = workbook.GetSheet("明细");
        if (duplicate)
            sheet.GetRow(1).CreateCell(0).SetCellValue("HEADER");
        else
            sheet.GetRow(0).GetCell(0).SetCellValue("BROKEN");
        using var changed = new MemoryStream();
        workbook.Write(changed, true);
        changed.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(duplicate ? 2 : 1, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>
    /// 验证复制页列表标题被改动时返回原始单元格坐标。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_CopiedListHeaderMismatch_ShouldReportSourceCoordinate(bool xls)
    {
        var layout = CreateLayout();
        var bytes = Export(CreateDocument(5), layout, xls);
        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        workbook.GetSheet("明细").GetRow(8).GetCell(1).SetCellValue("Changed");
        using var changed = new MemoryStream();
        workbook.Write(changed, true);
        changed.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));

        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(9, exception.RowIndex);
        Assert.Equal(2, exception.ColumnIndex);
    }

    /// <summary>验证布局构建时拒绝 marker 与固定绑定字段冲突。</summary>
    [Fact]
    public void PageHeader_MarkerConflictsWithFixedCell_ShouldRejectLayout()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<HeaderDocument>(builder => builder
            .Cell("明细", "A1", item => item.Title)
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER"))));
    }

    /// <summary>验证页眉源区域必须覆盖列表表头的全部列。</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_SourceRangeMissingHeaderColumn_ShouldRejectBeforeOutput(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:B3", "A1", "HEADER")));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证无法可靠复制的模板元数据使多页页眉导出在写入前失败。
    /// </summary>
    [Theory]
    [InlineData(true, "hyperlink")]
    [InlineData(false, "hyperlink")]
    [InlineData(true, "comment")]
    [InlineData(false, "comment")]
    [InlineData(true, "rich-text")]
    [InlineData(false, "rich-text")]
    [InlineData(true, "validation")]
    [InlineData(false, "validation")]
    [InlineData(true, "conditional-formatting")]
    [InlineData(false, "conditional-formatting")]
    public void PageHeader_SourceMetadata_ShouldRejectMultiPageBeforeOutput(bool xls, string metadata)
    {
        using var template = CreateMetadataHeaderTemplate(xls, metadata);
        using var output = new MemoryStream();
        var original = new byte[] { 4, 8, 1, 6 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(5), CreateLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证配置式数据校验与多页页眉源区域重叠时拒绝写入。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_ConfiguredValidationOverlap_ShouldRejectBeforeOutput(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER"))
            .DataValidation("明细", new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                {
                    StartRow = 0, EndRow = 0, StartColumn = 2, EndColumn = 2
                },
                Type = ExcelDataValidationType.Integer,
                Value1 = 1
            }));
        using var template = CreatePlainHeaderTemplate(xls);
        using var output = new MemoryStream();
        var original = new byte[] { 2, 7, 1, 8 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(5), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("配置式数据校验", exception.Message);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证重复页眉复制隐藏行和解锁单元格样式。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_HiddenRowAndUnlockedStyle_ShouldCopyAcrossPages(bool xls)
    {
        using var template = CreateMetadataHeaderTemplate(xls, "hidden-row", "unlocked");
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(5), CreateLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.GetSheet("明细");
        Assert.True(sheet.GetRow(1).ZeroHeight);
        Assert.True(sheet.GetRow(7).ZeroHeight);
        Assert.False(sheet.GetRow(6).GetCell(2).CellStyle.IsLocked);
    }

    /// <summary>
    /// 验证单页页眉无需复制时保留模板超链接、批注和隐藏行。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_SourceMetadata_ShouldRemainForSinglePage(bool xls)
    {
        using var template = CreateMetadataHeaderTemplate(xls,
            "hyperlink", "comment", "hidden-row");
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), CreateLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.GetSheet("明细");
        var cell = sheet.GetRow(0).GetCell(2);
        Assert.Equal("https://example.invalid/header", cell.Hyperlink.Address);
        Assert.Equal("页眉模板批注", cell.CellComment.String.String);
        Assert.Equal("template", cell.CellComment.Author);
        Assert.True(sheet.GetRow(1).ZeroHeight);
    }

    /// <summary>验证模板已有打印标题与实际页眉互斥。</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_TemplatePrintTitles_ShouldRejectBeforeOutput(bool xls)
    {
        using var template = CreateTemplate(xls, printTitles: true);
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), CreateLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>验证页眉预检失败时清除异步文件暂存。</summary>
    [Fact]
    public async Task PageHeader_PreflightFailure_ShouldCleanTemporaryStaging()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-offices-page-header-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var layout = CreateLayout();
            var staging = new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.TempFile,
                directory: directory);
            var exporter = new NpoiExcelExporter(null, staging);
            using var template = CreateTemplate(false, printTitles: true);

            await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                exporter.ExportForTemplateAsync(CreateDocument(3), layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), new MemoryStream()));
            Assert.Empty(Directory.EnumerateFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>验证同表无重叠的普通列表可以与实际页眉列表共存并往返。</summary>
    [Fact]
    public void PageHeader_WithNonOverlappingOrdinaryList_ShouldRoundTrip()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER").End("D9"))
            .ListRegion("明细", "J3", item => item.OtherLines, region => region.End("K4")));
        var source = CreateDocument(3);
        source.OtherLines.Add(new HeaderLine { Code = "R1" });
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
            Assert.Equal("HEADER", workbook.GetSheet("明细").GetRow(0).GetCell(0).StringCellValue);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(source.OtherLines.Select(line => line.Code), result.Entity.OtherLines.Select(line => line.Code));
    }

    /// <summary>验证未限定结束行的页眉列表不读取相邻列表的后续行。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_WithoutEnd_WithLongerAdjacentList_ShouldRoundTrip(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER"))
            .ListRegion("明细", "J3", item => item.OtherLines, region => region.End("K12")));
        var source = CreateDocument(3);
        for (var index = 0; index < 8; index++)
            source.OtherLines.Add(new HeaderLine { Code = $"R{index}" });
        var bytes = Export(source, layout, xls);
        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
            Assert.Equal("HEADER", workbook.GetSheet("明细").GetRow(0).GetCell(0).StringCellValue);
        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(source.OtherLines.Select(line => line.Code), result.Entity.OtherLines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证无模板时由 Fluent 填充并重复页眉。
    /// </summary>
    [Fact]
    public void PageHeader_WithoutTemplate_ShouldWriteFluentHeader()
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER", header => header
                    .Cell("B1", "采购单")
                    .Cell("C1", (ExcelEntityPageHeaderContext<HeaderLine> context) =>
                        $"{context.PageNumber}/{context.PageCount}"))
                .Footer("END")));
        var source = CreateDocument(3);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        var sheet = workbook.GetSheet("明细");
        Assert.Equal("HEADER", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("HEADER", sheet.GetRow(5).GetCell(0).StringCellValue);
        Assert.Equal("1/2", sheet.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("2/2", sheet.GetRow(5).GetCell(2).StringCellValue);
        Assert.Equal(new[] { 4 }, sheet.RowBreaks);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => item.Code),
            result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证组小计、分页小计和页眉按物理行顺序写入。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_GroupSubtotal_ShouldPreserveSegmentOrder(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER")
                .GroupSubtotal(line => line.Group, "GROUP")
                .PageSubtotalNamed("PageTotal", "PAGE")
                .Footer("END")));
        var source = CreateDocument(3);
        source.Lines[0].Group = "A";
        source.Lines[1].Group = "A";
        source.Lines[2].Group = "B";
        var bytes = Export(source, layout, xls);
        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("GROUP", sheet.GetRow(5).GetCell(1).StringCellValue);
            Assert.Equal("PAGE", sheet.GetRow(6).GetCell(1).StringCellValue);
            Assert.Equal("HEADER", sheet.GetRow(7).GetCell(0).StringCellValue);
            Assert.Equal("GROUP", sheet.GetRow(11).GetCell(1).StringCellValue);
            Assert.Equal("END", sheet.GetRow(12).GetCell(1).StringCellValue);
            Assert.Equal(new[] { 6 }, sheet.RowBreaks);
        }
        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(item => item.Code),
            result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证页眉值工厂失败时保留原目标文件。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_ValueFactoryFailure_ShouldPreserveTarget(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER", header => header
                    .Cell("C1", (ExcelEntityPageHeaderContext<HeaderLine> context) =>
                        context.PageNumber == 2 ? throw new InvalidOperationException("fail") : "ok"))));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + (xls ? ".xls" : ".xlsx"));
        try
        {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            Assert.Throws<BingOfficesExportException>(() => new NpoiExcelExporter().ExportEntityToFile(
                CreateDocument(3), layout, path));
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>验证标记处于模板合并区域内部时预检失败。</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_MarkerInsideMerge_ShouldRejectBeforeOutput(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        using var template = new MemoryStream();
        workbook.Write(template, true);
        template.Position = 0;
        using var output = new MemoryStream();
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "B1", "HEADER")));
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(0, output.Length);
    }

    /// <summary>验证相对合并与模板合并冲突时预检失败。</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_FluentMergeConflictsWithTemplate_ShouldRejectBeforeOutput(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER", header => header
                    .Merge("A2:B2"))));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(0, output.Length);
    }

    /// <summary>验证后续页眉计算期间取消时不写入调用方流。</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PageHeader_MidWriteCancellation_ShouldPreserveOutput(bool xls)
    {
        using var cancellation = new CancellationTokenSource();
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER", header => header
                    .Cell("C1", (ExcelEntityPageHeaderContext<HeaderLine> context) =>
                    {
                        if (context.PageNumber == 2)
                            cancellation.Cancel();
                        return context.PageNumber;
                    }))));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        var original = new byte[] { 3, 1, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        var tempDirectory = Path.GetTempPath();
        var stagedFilesBefore = new HashSet<string>(
            Directory.GetFiles(tempDirectory, "bing-offices-entity-template-async-*.tmp"),
            StringComparer.OrdinalIgnoreCase);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new NpoiExcelExporter().ExportForTemplateAsync(CreateDocument(3), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output,
                cancellation.Token));
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
        Assert.Empty(Directory.GetFiles(tempDirectory, "bing-offices-entity-template-async-*.tmp")
            .Where(path => !stagedFilesBefore.Contains(path)));
    }

    /// <summary>验证复制页眉时取消会在下一个来源单元格前停止。</summary>
    [Theory]
    [InlineData(true, 0, 0)]
    [InlineData(false, 0, 0)]
    [InlineData(true, 2, 2)]
    [InlineData(false, 2, 2)]
    public void PageHeader_CancellationDuringSourceCopy_ShouldStopBeforeFinalization(
        bool xls, int cancellationRow, int cancellationColumn)
    {
        using var cancellation = new CancellationTokenSource();
        var header = CreateLayout().ListRegions[0].PageHeader;
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        for (var rowIndex = 0; rowIndex < 3; rowIndex++)
        {
            var row = sheet.CreateRow(rowIndex);
            for (var columnIndex = 0; columnIndex < 3; columnIndex++)
                row.CreateCell(columnIndex, CellType.String).SetCellValue($"{rowIndex}:{columnIndex}");
        }
        var copiedCells = new List<(int Row, int Column)>();

        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiEntityExportExecutor.WritePageHeader(workbook, sheet, header, targetRow: 5,
                pageNumber: 2, pageCount: 2, items: Array.Empty<object>(), copySource: true,
                cancellationToken: cancellation.Token, afterSourceCellCopied: (row, column) =>
                {
                    copiedCells.Add((row, column));
                    if (row == cancellationRow && column == cancellationColumn)
                        cancellation.Cancel();
                }));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(cancellationRow * 3 + cancellationColumn + 1, copiedCells.Count);
        Assert.Equal((cancellationRow, cancellationColumn), copiedCells[^1]);
        Assert.Equal($"{cancellationRow}:{cancellationColumn}",
            sheet.GetRow(5 + cancellationRow).GetCell(cancellationColumn).StringCellValue);
        Assert.Equal(0, sheet.NumMergedRegions);
        if (cancellationRow == 0)
            Assert.Null(sheet.GetRow(5).GetCell(1));
        else
            Assert.Equal("0:0", sheet.GetRow(5).GetCell(0).StringCellValue);
        Assert.Equal("0:1", sheet.GetRow(0).GetCell(1).StringCellValue);
    }

    /// <summary>验证复制页眉超出 XLS 或 XLSX 最大行数时预检失败。</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_MaxRow_ShouldRejectBeforeOutput(bool xls)
    {
        var last = xls ? 65534 : 1048574;
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", $"B{last}", item => item.Lines, region => region
                .PageBreak(1).PageHeader($"A{last - 1}:C{last}", $"A{last - 1}", "HEADER")));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(2), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(0, output.Length);
    }

    /// <summary>
    /// 验证无模板页眉复制目标占用固定字段时拒绝导出，包括空固定值。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageHeader_WithoutTemplateFixedCellOverlap_ShouldRejectBeforeOutput(bool emptyValue)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .Cell("明细", "A5", item => item.Title)
            .ListRegion("明细", "B2", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C2", "A1", "HEADER")));
        var document = CreateDocument(3);
        document.Title = emptyValue ? null : "BUSINESS";
        using var output = new MemoryStream();
        var original = new byte[] { 7, 2, 9 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().ExportEntity(
            document, layout, output));

        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证 Header(false) 命名列表与页眉相邻时往返，改动命名起点后拒绝漏读。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_NamedListWithoutHeader_ShouldValidateResolvedAdjacency(bool xls)
    {
        var layout = CreateNamedHeaderLayout();
        var document = CreateDocument(3);
        byte[] bytes;
        using (var template = CreateNamedHeaderTemplate(xls, "$B$3"))
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(document, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            bytes = output.ToArray();
        }

        using (var input = new MemoryStream(bytes, writable: false))
        {
            var result = new NpoiExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(document.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
        }

        var moved = RewriteNamedAnchor(bytes, "$B$4");
        using var changed = new MemoryStream(moved, writable: false);
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
    }

    /// <summary>
    /// 验证 End 中的预留空行不计入明细，且损坏页眉 marker 报告精确原始坐标。
    /// </summary>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void PageHeader_EndReservedRows_ShouldRoundTripOrReportMissingMarker(bool xls, bool breakMarker)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2).PageHeader("A1:C3", "A1", "HEADER").End("D30")));
        var bytes = Export(CreateDocument(3), layout, xls);
        if (!breakMarker)
        {
            using var input = new MemoryStream(bytes, writable: false);
            var result = new NpoiExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
            return;
        }

        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        workbook.GetSheet("明细").GetRow(5).GetCell(0).SetCellValue("BROKEN");
        using var changed = new MemoryStream();
        workbook.Write(changed, true);
        changed.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(6, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>
    /// 验证稀疏工作簿配合格式最大结束行仍只导入实际明细。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_SparseLargeEnd_ShouldRoundTrip(bool xls)
    {
        var layout = CreateSparseLargeEndLayout(xls);
        var bytes = ExportWithPlainHeaderTemplate(CreateDocument(5), layout, xls);

        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3", "L4", "L5" },
            result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证稀疏大结束行布局仍受实体导入输入字节限额约束。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_SparseLargeEnd_InputLimit_ShouldReturnResourceError(bool xls)
    {
        var layout = CreateSparseLargeEndLayout(xls);
        var bytes = ExportWithPlainHeaderTemplate(CreateDocument(5), layout, xls);
        using var input = new MemoryStream(bytes, writable: false);
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = bytes.LongLength - 1
        });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new NpoiExcelImporter().ImportEntity(input, layout, options));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Open, exception.Stage);
        Assert.True(input.CanRead);
    }

    /// <summary>
    /// 验证稀疏大结束行布局尊重导入取消令牌。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PageHeader_SparseLargeEnd_CancelledImport_ShouldThrow(bool xls)
    {
        var layout = CreateSparseLargeEndLayout(xls);
        var bytes = ExportWithPlainHeaderTemplate(CreateDocument(5), layout, xls);
        using var input = new MemoryStream(bytes, writable: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new NpoiExcelImporter().ImportEntityAsync(input, layout, cancellation.Token));
    }

    /// <summary>
    /// 验证无分页小计时第二页缺失 marker 的错误仍指向 marker 原始位置。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_WithoutPageSubtotal_MissingMarkerShouldReportExactCoordinate(bool xls)
    {
        var layout = CreateNamedHeaderLayout();
        byte[] bytes;
        using (var template = CreateNamedHeaderTemplate(xls, "$B$3"))
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(3), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            bytes = output.ToArray();
        }

        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        workbook.GetSheet("明细").GetRow(5).GetCell(0).SetCellValue("BROKEN");
        using var changed = new MemoryStream();
        workbook.Write(changed, true);
        changed.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(changed, layout));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(6, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>
    /// 验证重复页眉的异步导出结果可完整读回。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PageHeader_AsyncExport_ShouldRoundTrip(bool xls)
    {
        var layout = CreateLayout();
        var document = CreateDocument(3);
        using var output = new MemoryStream();
        using var template = CreateTemplate(xls);
        await new NpoiExcelExporter().ExportForTemplateAsync(document, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false)))
        {
            Assert.Equal("HEADER", workbook.GetSheet("明细").GetRow(6).GetCell(0).StringCellValue);
            Assert.Equal("END", workbook.GetSheet("明细").GetRow(10).GetCell(1).StringCellValue);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(document.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证页眉工作簿提交失败时保留原文件、清理暂存文件并只观察一次异常。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageHeader_FileCommitFailure_ShouldPreserveTargetAndCleanupTemporaryFile(bool asynchronous)
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-page-header-commit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "invoice.xlsx");
        var original = new byte[] { 4, 2, 8, 1 };
        File.WriteAllBytes(path, original);
        try
        {
            using var lockedTarget = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var observer = new PageHeaderRecordingObserver();
            var exporter = new NpoiExcelExporter(exceptionObservers: new[] { observer });

            BingOfficesFileCommitException exception;
            if (asynchronous)
                exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                    exporter.ExportEntityToFileAsync(CreateDocument(3), CreateLayout(), path));
            else
                exception = Assert.Throws<BingOfficesFileCommitException>(() =>
                    exporter.ExportEntityToFile(CreateDocument(3), CreateLayout(), path));

            Assert.Equal(BingOfficesStage.Commit, exception.Stage);
            Assert.Equal(BingOfficesErrorCode.FileCommitFailed, exception.Code);
            Assert.Single(observer.Exceptions);
            Assert.Same(exception, observer.Exceptions[0]);
            Assert.Equal(original.Length, lockedTarget.Length);
            lockedTarget.Position = 0;
            var actual = new byte[original.Length];
            Assert.Equal(original.Length, lockedTarget.Read(actual, 0, actual.Length));
            Assert.Equal(original, actual);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    /// <summary>
    /// 验证动态列和计算列的最终物理宽度被完整复制，并拒绝过窄源区域。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_DynamicAndCalculatedColumns_ShouldUseCompletePhysicalRange(bool xls)
    {
        var layout = CreatePhysicalHeaderLayout("A1:D3");
        var source = CreatePhysicalHeaderDocument();
        byte[] bytes;
        using (var template = CreatePhysicalHeaderTemplate(xls))
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            bytes = output.ToArray();
        }

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            var firstHeader = Enumerable.Range(0, 4).Select(column =>
                sheet.GetRow(2).GetCell(column).StringCellValue).ToArray();
            var repeatedHeader = Enumerable.Range(0, 4).Select(column =>
                sheet.GetRow(6).GetCell(column).StringCellValue).ToArray();
            Assert.Equal(new[] { "序号", "Code", "Quantity", "颜色" }, firstHeader);
            Assert.Equal(firstHeader, repeatedHeader);
            Assert.Equal("2/2", sheet.GetRow(4).GetCell(1).StringCellValue);
            Assert.Equal("蓝", sheet.GetRow(7).GetCell(3).StringCellValue);
        }

        using (var input = new MemoryStream(bytes, writable: false))
        {
            var result = new NpoiExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(new[] { "L1", "L2" }, result.Entity.Lines.Select(line => line.Code));
            Assert.Equal(new[] { "红", "蓝" }, result.Entity.Lines.Select(line =>
                line.Values["color"].ToString()));
        }

        var narrowLayout = CreatePhysicalHeaderLayout("A1:C3");
        using var narrowTemplate = CreatePhysicalHeaderTemplate(xls);
        using var protectedOutput = new MemoryStream();
        var original = new byte[] { 2, 7, 1 };
        protectedOutput.Write(original, 0, original.Length);
        protectedOutput.Position = 0;
        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().ExportForTemplate(
            source, narrowLayout, new ExcelEntityTemplateOptions(narrowTemplate, leaveOpen: true), protectedOutput));
        Assert.Equal(original, protectedOutput.ToArray());
        Assert.True(narrowTemplate.CanRead);
    }

    /// <summary>
    /// 验证页眉源区域高度必须与 Header 模式的首个明细位置精确相邻。
    /// </summary>
    [Theory]
    [InlineData(true, true)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(false, false)]
    public void PageHeader_InvalidHeightForHeaderMode_ShouldRejectBeforeOutput(bool xls, bool includeHeader)
    {
        var sourceRange = includeHeader ? "A1:C3" : "A1:C2";
        var end = includeHeader ? "D6" : "D4";
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .Header(includeHeader)
                .PageBreak(1)
                .PageHeader(sourceRange, "A1", "HEADER")
                .End(end)));
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        var original = new byte[] { 8, 6, 7 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().ExportForTemplate(
            CreateDocument(2), layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证明细值与实际页眉 marker 同值时在写入目标流前返回配置错误。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_DetailValueMatchingMarker_ShouldRejectBeforeOutput(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "A3", item => item.Lines, region => region
                .PageBreak(1).PageHeader("A1:C3", "A1", "HEADER")));
        var document = CreateDocument(2);
        document.Lines[0].Code = "HEADER";
        using var template = CreatePlainHeaderTemplate(xls);
        using var output = new MemoryStream();
        var original = new byte[] { 5, 3, 1 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(document, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证无表头列表的页眉 marker 列与明细列重叠时在写入前拒绝。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_HeaderlessMarkerInsideDetailColumns_ShouldRejectBeforeOutput(bool xls)
    {
        var layout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .Header(false)
                .PageBreak(1)
                .PageHeader("A1:C2", "B2", "HEADER")));
        using var template = CreatePlainHeaderTemplate(xls);
        using var output = new MemoryStream();
        var original = new byte[] { 3, 1, 4, 1, 5 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(2), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证页眉不能改写启用的列表标题，但保留相同标题时可以往返。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_FluentCellWithDifferentListTitle_ShouldReject(bool xls)
    {
        var document = CreateDocument(2);
        var invalidLayout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(1).PageHeader("A1:C3", "A1", "HEADER", header =>
                    header.Cell("B3", "Renamed"))));
        using (var template = CreatePlainHeaderTemplate(xls))
        using (var output = new MemoryStream())
        {
            var original = new byte[] { 6, 2, 6 };
            output.Write(original, 0, original.Length);
            output.Position = 0;
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelExporter().ExportForTemplate(document, invalidLayout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
            Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
            Assert.Equal(original, output.ToArray());
            Assert.True(template.CanRead);
        }
    }

    /// <summary>
    /// 验证页眉保留相同列表标题时可以往返。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_FluentCellWithMatchingListTitle_ShouldRoundTrip(bool xls)
    {
        var document = CreateDocument(2);
        var validLayout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(1).PageHeader("A1:C3", "A1", "HEADER", header =>
                    header.Cell("B3", "Code"))));
        byte[] bytes;
        using (var template = CreatePlainHeaderTemplate(xls))
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(document, validLayout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            bytes = output.ToArray();
        }

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("Code", sheet.GetRow(2).GetCell(1).StringCellValue);
            Assert.Equal("Code", sheet.GetRow(6).GetCell(1).StringCellValue);
        }
        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, validLayout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(document.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证页眉合并不能跨列表表头，表头之外的合并仍可导出并往返。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_MergeOverlappingListHeader_ShouldReject(bool xls)
    {
        var document = CreateDocument(2);
        var invalidLayout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(1).PageHeader("A1:C3", "A1", "HEADER", header =>
                    header.Merge("B3:C3"))));
        using (var template = CreatePlainHeaderTemplate(xls))
        using (var output = new MemoryStream())
        {
            var original = new byte[] { 9, 1, 4 };
            output.Write(original, 0, original.Length);
            output.Position = 0;
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelExporter().ExportForTemplate(document, invalidLayout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
            Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Equal(original, output.ToArray());
            Assert.True(template.CanRead);
        }
    }

    /// <summary>
    /// 验证页眉表头之外的合并可以导出并往返。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_MergeOutsideListHeader_ShouldRoundTrip(bool xls)
    {
        var document = CreateDocument(2);
        var validLayout = ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(1).PageHeader("A1:C3", "A1", "HEADER", header =>
                    header.Merge("A2:B2"))));
        byte[] bytes;
        using (var template = CreatePlainHeaderTemplate(xls))
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(document, validLayout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            bytes = output.ToArray();
        }

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Contains(sheet.MergedRegions, merge => merge.FirstRow == 1 && merge.LastRow == 1
                && merge.FirstColumn == 0 && merge.LastColumn == 1);
            Assert.Contains(sheet.MergedRegions, merge => merge.FirstRow == 5 && merge.LastRow == 5
                && merge.FirstColumn == 0 && merge.LastColumn == 1);
        }
        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, validLayout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(document.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
    }

    private static ExcelEntityLayout<HeaderDocument> CreateLayout() =>
        ExcelEntity.Layout<HeaderDocument>(builder => builder
            .Cell("明细", "B1", item => item.Title)
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2)
                .PageHeader("A1:C3", "A1", "HEADER", header => header
                    .Cell("C1", (ExcelEntityPageHeaderContext<HeaderLine> context) =>
                        $"{context.PageNumber}/{context.PageCount}"))
                .PageSubtotalNamed("PageTotal", "PAGE")
                .Footer("END")));

    private static ExcelEntityLayout<HeaderDocument> CreateSparseLargeEndLayout(bool xls) =>
        ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegion("明细", "B3", item => item.Lines, region => region
                .PageBreak(2)
                .PageHeader("A1:C3", "A1", "HEADER")
                .End(xls ? "D65536" : "D1048576")));

    private static ExcelEntityLayout<HeaderDocument> CreateNamedHeaderLayout() =>
        ExcelEntity.Layout<HeaderDocument>(builder => builder
            .ListRegionNamed("明细", "LinesAnchor", item => item.Lines, region => region
                .Header(false)
                .PageBreak(2)
                .PageHeader("A1:C2", "A2", "HEADER")));

    private static ExcelEntityLayout<PhysicalHeaderDocument> CreatePhysicalHeaderLayout(string sourceRange) =>
        ExcelEntity.Layout<PhysicalHeaderDocument>(builder => builder
            .ListRegion("明细", "A3", document => document.Lines, region => region
                .CalculatedColumn<int>("line-number", "序号", context => context.Index + 1,
                    column => column.Placement(ExcelColumnPlacement.Before("Code")))
                .DynamicColumnGroup("商品", line => line.Values, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "color", Title = "颜色", DataType = typeof(string)
                    }
                })
                .PageBreak(1)
                .PageHeader(sourceRange, "A1", "PAGE", header => header
                    .Cell("B1", context => $"{context.PageNumber}/{context.PageCount}"))
                .End("D20")));

    private static PhysicalHeaderDocument CreatePhysicalHeaderDocument() => new()
    {
        Lines = new List<PhysicalHeaderLine>
        {
            new() { Code = "L1", Quantity = 1, Values = new Dictionary<string, object> { ["color"] = "红" } },
            new() { Code = "L2", Quantity = 2, Values = new Dictionary<string, object> { ["color"] = "蓝" } }
        }
    };

    private static HeaderDocument CreateDocument(int count) => new()
    {
        Title = "采购单",
        Lines = Enumerable.Range(1, count).Select(index => new HeaderLine
        {
            Code = "L" + index
        }).ToList()
    };

    private static byte[] Export(HeaderDocument source,
        ExcelEntityLayout<HeaderDocument> layout, bool xls)
    {
        using var template = CreateTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        return output.ToArray();
    }

    private static byte[] ExportWithPlainHeaderTemplate(HeaderDocument source,
        ExcelEntityLayout<HeaderDocument> layout, bool xls)
    {
        using var template = CreatePlainHeaderTemplate(xls);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        return output.ToArray();
    }

    private static MemoryStream CreateTemplate(bool xls, bool printTitles = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        if (printTitles)
            sheet.RepeatingRows = CellRangeAddress.ValueOf("$1:$3");
        var row = sheet.CreateRow(0);
        row.Height = 480;
        row.CreateCell(0).SetCellValue("HEADER");
        var second = sheet.CreateRow(1);
        second.CreateCell(0).SetCellFormula("1+1");
        var title = second.CreateCell(1);
        title.SetCellValue("采购信息");
        var style = workbook.CreateCellStyle();
        style.FillPattern = FillPattern.SolidForeground;
        style.FillForegroundColor = IndexedColors.LightYellow.Index;
        title.CellStyle = style;
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 2));
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreatePlainHeaderTemplate(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("HEADER");
        sheet.CreateRow(1);
        sheet.CreateRow(2);
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateMetadataHeaderTemplate(bool xls, params string[] metadata)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var marker = sheet.CreateRow(0).CreateCell(0);
        marker.SetCellValue("HEADER");
        var metadataCell = sheet.GetRow(0).CreateCell(2);
        if (metadata.Contains("hyperlink", StringComparer.Ordinal))
        {
            var hyperlink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
            hyperlink.Address = "https://example.invalid/header";
            metadataCell.Hyperlink = hyperlink;
        }
        if (metadata.Contains("comment", StringComparer.Ordinal))
        {
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = 2;
            anchor.Col2 = 4;
            anchor.Row1 = 0;
            anchor.Row2 = 2;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString("页眉模板批注");
            comment.Author = "template";
            metadataCell.CellComment = comment;
        }
        if (metadata.Contains("unlocked", StringComparer.Ordinal))
        {
            var style = workbook.CreateCellStyle();
            style.IsLocked = false;
            metadataCell.CellStyle = style;
        }
        if (metadata.Contains("rich-text", StringComparer.Ordinal))
        {
            var text = workbook.GetCreationHelper().CreateRichTextString("formatted");
            var font = workbook.CreateFont();
            text.ApplyFont(0, 1, font);
            metadataCell.SetCellValue(text);
        }
        if (metadata.Contains("validation", StringComparer.Ordinal))
        {
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateExplicitListConstraint(new[] { "A", "B" });
            sheet.AddValidationData(helper.CreateValidation(constraint,
                new CellRangeAddressList(0, 0, 2, 2)));
        }
        if (metadata.Contains("conditional-formatting", StringComparer.Ordinal))
        {
            var conditionalFormatting = sheet.SheetConditionalFormatting;
            var rule = conditionalFormatting.CreateConditionalFormattingRule(
                ComparisonOperator.GreaterThan, "0");
            conditionalFormatting.AddConditionalFormatting(
                new[] { new CellRangeAddress(0, 0, 2, 2) }, rule);
        }
        var hidden = sheet.CreateRow(1);
        hidden.ZeroHeight = metadata.Contains("hidden-row", StringComparer.Ordinal);
        sheet.CreateRow(2);

        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateNamedHeaderTemplate(bool xls, string anchorAddress)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("采购信息");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("HEADER");
        var anchor = workbook.CreateName();
        anchor.NameName = "LinesAnchor";
        anchor.RefersToFormula = "'明细'!" + anchorAddress;
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreatePhysicalHeaderTemplate(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("PAGE");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("采购信息");
        sheet.CreateRow(2);
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static byte[] RewriteNamedAnchor(byte[] bytes, string anchorAddress)
    {
        using var source = new MemoryStream(bytes, writable: false);
        using var workbook = WorkbookFactory.Create(source);
        var anchor = Assert.Single(workbook.GetAllNames(), name => name.NameName == "LinesAnchor");
        anchor.RefersToFormula = "'明细'!" + anchorAddress;
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>保存测试单据。</summary>
    public sealed class HeaderDocument
    {
        /// <summary>获取或设置标题。</summary>
        public string Title { get; set; }

        /// <summary>获取或设置明细。</summary>
        public List<HeaderLine> Lines { get; set; } = new();

        /// <summary>获取或设置并列表格行。</summary>
        public List<HeaderLine> OtherLines { get; set; } = new();
    }

    /// <summary>保存测试明细。</summary>
    public sealed class HeaderLine
    {
        /// <summary>获取或设置编码。</summary>
        public string Code { get; set; }

        /// <summary>获取或设置分组。</summary>
        public string Group { get; set; }
    }

    /// <summary>
    /// 保存动态和计算列页眉测试的根模型。
    /// </summary>
    public sealed class PhysicalHeaderDocument
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<PhysicalHeaderLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 保存动态和计算列页眉测试的明细。
    /// </summary>
    public sealed class PhysicalHeaderLine
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 获取或设置动态列值。
        /// </summary>
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 记录文件提交失败时的异常观察调用。
    /// </summary>
    private sealed class PageHeaderRecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 获取观察到的异常集合。
        /// </summary>
        public List<BingOfficesException> Exceptions { get; } = new();

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }
}
