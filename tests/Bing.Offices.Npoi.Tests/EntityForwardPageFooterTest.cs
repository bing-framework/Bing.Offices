using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI Forward 每页尾部边界测试。
/// </summary>
public sealed class EntityForwardPageFooterTest
{
    /// <summary>
    /// 验证 XLS/XLSX 的同步和异步 Forward 导入会跳过每页尾部及最终尾部。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls));
        var batches = new List<ExcelEntityImportBatch<PageDocument>>();
        var layout = CreateLayout();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, layout, options, batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 1, 2, 3, 4 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
        Assert.Equal(new[] { 2, 3, 7, 8 },
            batches.Select(batch => batch.FirstRowIndex).ToArray());
        Assert.Equal(4, summary.RowsRead);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
    }

    /// <summary>
    /// 验证 Forward 会同时跳过实际重复页眉和每页尾部。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageHeaderAndPageFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateHeaderFooterLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHeaderFooterLayout(), options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 1, 2, 3, 4 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
        Assert.Equal(4, summary.RowsRead);
    }

    /// <summary>
    /// 验证实际重复页眉可与分页小计、每页尾部和最终尾部组合。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageHeaderWithPageSubtotalAndFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterPageSubtotalWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source,
                CreateHeaderFooterPageSubtotalLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHeaderFooterPageSubtotalLayout(), options,
                _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(4, summary.RowsRead);
    }

    /// <summary>
    /// 验证实际重复页眉可与分组小计、每页尾部和最终尾部组合。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageHeaderWithGroupSubtotalAndFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterGroupSubtotalWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source,
                CreateHeaderFooterGroupSubtotalLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHeaderFooterGroupSubtotalLayout(), options,
                _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(4, summary.RowsRead);
    }

    /// <summary>
    /// 验证组合路径会裁剪最终尾部前的空白间隔和多行尾部内容。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldSkipGappedMultiRowFinalFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterGappedFinalFooterWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source,
                CreateHeaderFooterGappedFinalFooterLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source,
                CreateHeaderFooterGappedFinalFooterLayout(), options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(4, summary.RowsRead);
    }

    /// <summary>
    /// 验证组合页眉和每页尾部路径在异步回调取消后停止交付后续批次。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ImportEntityBatchesAsync_Forward_PageHeaderAndFooter_ShouldStopAfterCallbackCancellation(
        bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterWorkbook(xls));
        using var cancellation = new CancellationTokenSource();
        var callbackCount = 0;
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            importer.ImportEntityBatchesAsync(source, CreateHeaderFooterLayout(), options,
                (_, _) =>
                {
                    callbackCount++;
                    cancellation.Cancel();
                    return Task.CompletedTask;
                }, cancellation.Token));

        Assert.Equal(1, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证组合页眉和每页尾部路径可从不可定位输入暂存后读取。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_PageHeaderAndFooter_ShouldReadNonSeekableInput(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new NonSeekableReadStream(CreateHeaderFooterWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateHeaderFooterLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHeaderFooterLayout(), options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(4, summary.RowsRead);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实际重复页眉首行标记错位时返回原始坐标配置错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldRejectPageHeaderAndPageFooterMisalignment(
        bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterWorkbook(xls,
            invalidRepeatedHeader: true));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateHeaderFooterLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(6, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>
    /// 验证最后一页缺少每页尾部时返回结构化配置错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldRejectMissingLastPageFooter(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, omitLastPageFooter: true));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("每页尾部", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证同一逻辑页重复出现每页尾部标记时返回结构化错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldRejectDuplicatePageFooter(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, duplicatePageFooter: true));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("重复", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证最后一个逻辑页只有页脚而没有明细时返回结构化错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldRejectEmptyLastPageFooter(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, emptyLastPage: true));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("每页尾部", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证前向页脚读取在预取消时不交付批次并保持输入流可读。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldHonorPreCancellation(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        if (asynchronous)
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), options, _ => { },
                    cancellation.Token));
        }

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证每页尾部前的分页小计不会被识别为明细。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldReadPageSubtotalBeforePageFooter(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateSubtotalWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = importer.ImportEntityBatches(source, CreateSubtotalLayout(), options,
            _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 1, 2, 3, 4 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
    }

    /// <summary>
    /// 验证分组小计跨逻辑页时会在每页尾部前正确裁剪。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadGroupSubtotalBeforePageFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateGroupSubtotalWorkbook(xls));
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateGroupSubtotalLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateGroupSubtotalLayout(), options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2", "B-3" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(5, summary.RowsRead);
    }

    private static ExcelEntityLayout<PageDocument> CreateLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
                .PageBreak(2)
                .PageFooter("SIGN", footer => footer.GapRows(1).Cell("B2",
                    items => items.Count))
                .Footer("END", footer => footer.GapRows(1))));

    private static ExcelEntityLayout<PageDocument> CreateHeaderFooterLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A2",
            document => document.Lines, region => region
                .End("B10")
                .PageBreak(2)
                .PageHeader("A1:B2", "A1", "HDR")
                .PageFooter("SIGN")));

    private static ExcelEntityLayout<PageDocument> CreateHeaderFooterPageSubtotalLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A2",
            document => document.Lines, region => region
                .End("B12")
                .PageBreak(2)
                .PageHeader("A1:B2", "A1", "HDR")
                .PageSubtotal("PAGE")
                .PageFooter("SIGN")
                .Footer("END")));

    private static ExcelEntityLayout<GroupPageDocument> CreateHeaderFooterGroupSubtotalLayout() =>
        ExcelEntity.Layout<GroupPageDocument>(builder => builder.ListRegion("明细", "A2",
            document => document.Lines, region => region
                .End("B13")
                .PageBreak(2)
                .PageHeader("A1:B2", "A1", "HDR")
                .GroupSubtotal(item => item.Group, "GROUP")
                .PageFooter("SIGN")
                .Footer("END")));

    private static ExcelEntityLayout<PageDocument> CreateHeaderFooterGappedFinalFooterLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A2",
            document => document.Lines, region => region
                .End("B13")
                .PageBreak(2)
                .PageHeader("A1:B2", "A1", "HDR")
                .PageFooter("SIGN")
                .Footer("END", footer => footer.GapRows(1).Cell("A2", "TAIL"))));

    private static ExcelEntityLayout<PageDocument> CreateSubtotalLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE", footer => footer.GapRows(1).Cell("B2",
                    items => items.Sum(item => item.Quantity)))
                .PageFooter("SIGN", footer => footer.GapRows(1))
                .Footer("END", footer => footer.GapRows(1))));

    private static ExcelEntityLayout<GroupPageDocument> CreateGroupSubtotalLayout() =>
        ExcelEntity.Layout<GroupPageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
                .GroupSubtotal(item => item.Group, "GROUP")
                .PageBreak(3)
                .PageFooter("SIGN")
                .Footer("END")));

    private static byte[] CreateWorkbook(bool xls, bool omitLastPageFooter = false,
        bool duplicatePageFooter = false, bool emptyLastPage = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        CreateLine(sheet, 1, "A-1", 1);
        CreateLine(sheet, 2, "A-2", 2);
        sheet.CreateRow(4).CreateCell(0).SetCellValue("SIGN");
        sheet.CreateRow(5).CreateCell(1).SetCellValue(2d);
        if (duplicatePageFooter)
            sheet.GetRow(5).CreateCell(0).SetCellValue("SIGN");
        CreateLine(sheet, 6, "B-1", 3);
        CreateLine(sheet, 7, "B-2", 4);
        if (omitLastPageFooter)
            sheet.CreateRow(9).CreateCell(0).SetCellValue("END");
        else if (emptyLastPage)
        {
            sheet.CreateRow(6).CreateCell(0).SetCellValue("SIGN");
            sheet.CreateRow(7).CreateCell(0).SetCellValue("END");
        }
        else
        {
            sheet.CreateRow(9).CreateCell(0).SetCellValue("SIGN");
            sheet.CreateRow(10).CreateCell(1).SetCellValue(2d);
            sheet.CreateRow(12).CreateCell(0).SetCellValue("END");
        }

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterWorkbook(bool xls,
        bool invalidRepeatedHeader = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        WriteRow(sheet, 0, "HDR", "采购单");
        WriteRow(sheet, 1, "Code", "Quantity");
        CreateLine(sheet, 2, "A-1", 1);
        CreateLine(sheet, 3, "A-2", 2);
        WriteRow(sheet, 4, "SIGN");
        WriteRow(sheet, 5, invalidRepeatedHeader ? "WRONG" : "HDR", "采购单");
        WriteRow(sheet, 6, "Code", "Quantity");
        CreateLine(sheet, 7, "B-1", 3);
        CreateLine(sheet, 8, "B-2", 4);
        WriteRow(sheet, 9, "SIGN");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterPageSubtotalWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        WriteRow(sheet, 0, "HDR", "采购单");
        WriteRow(sheet, 1, "Code", "Quantity");
        CreateLine(sheet, 2, "A-1", 1);
        CreateLine(sheet, 3, "A-2", 2);
        WriteRow(sheet, 4, "PAGE");
        WriteRow(sheet, 5, "SIGN");
        WriteRow(sheet, 6, "HDR", "采购单");
        WriteRow(sheet, 7, "Code", "Quantity");
        CreateLine(sheet, 8, "B-1", 3);
        CreateLine(sheet, 9, "B-2", 4);
        WriteRow(sheet, 10, "SIGN");
        WriteRow(sheet, 11, "END");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterGroupSubtotalWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        WriteRow(sheet, 0, "HDR", "采购单");
        WriteRow(sheet, 1, "Code", "Group");
        CreateGroupedLine(sheet, 2, "A-1", "A");
        CreateGroupedLine(sheet, 3, "A-2", "A");
        WriteRow(sheet, 4, "GROUP");
        WriteRow(sheet, 5, "SIGN");
        WriteRow(sheet, 6, "HDR", "采购单");
        WriteRow(sheet, 7, "Code", "Group");
        CreateGroupedLine(sheet, 8, "B-1", "B");
        CreateGroupedLine(sheet, 9, "B-2", "B");
        WriteRow(sheet, 10, "GROUP");
        WriteRow(sheet, 11, "SIGN");
        WriteRow(sheet, 12, "END");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterGappedFinalFooterWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        WriteRow(sheet, 0, "HDR", "采购单");
        WriteRow(sheet, 1, "Code", "Quantity");
        CreateLine(sheet, 2, "A-1", 1);
        CreateLine(sheet, 3, "A-2", 2);
        WriteRow(sheet, 4, "SIGN");
        WriteRow(sheet, 5, "HDR", "采购单");
        WriteRow(sheet, 6, "Code", "Quantity");
        CreateLine(sheet, 7, "B-1", 3);
        CreateLine(sheet, 8, "B-2", 4);
        WriteRow(sheet, 9, "SIGN");
        sheet.CreateRow(10);
        WriteRow(sheet, 11, "END");
        WriteRow(sheet, 12, "TAIL");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 提供不可定位但可读取的实体导入输入流。
    /// </summary>
    private sealed class NonSeekableReadStream : Stream
    {
        /// <summary>
        /// 保存实际输入内容的只读流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个 <see cref="NonSeekableReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="buffer">输入工作簿内容。</param>
        internal NonSeekableReadStream(byte[] buffer)
        {
            _inner = new MemoryStream(buffer, writable: false);
        }

        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;

        /// <inheritdoc />
        public override bool CanSeek => false;

        /// <inheritdoc />
        public override bool CanWrite => false;

        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => throw new NotSupportedException();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) =>
            _inner.ReadAsync(buffer, offset, count, cancellationToken);

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default) =>
            _inner.ReadAsync(buffer, cancellationToken);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private static byte[] CreateSubtotalWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        CreateLine(sheet, 1, "A-1", 1);
        CreateLine(sheet, 2, "A-2", 2);
        sheet.CreateRow(4).CreateCell(0).SetCellValue("PAGE");
        sheet.CreateRow(5).CreateCell(1).SetCellValue(3d);
        sheet.CreateRow(7).CreateCell(0).SetCellValue("SIGN");
        CreateLine(sheet, 8, "B-1", 3);
        CreateLine(sheet, 9, "B-2", 4);
        sheet.CreateRow(11).CreateCell(0).SetCellValue("SIGN");
        sheet.CreateRow(13).CreateCell(0).SetCellValue("END");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static byte[] CreateGroupSubtotalWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Group");
        CreateGroupedLine(sheet, 1, "A-1", "A");
        CreateGroupedLine(sheet, 2, "A-2", "A");
        sheet.CreateRow(3).CreateCell(0).SetCellValue("GROUP");
        CreateGroupedLine(sheet, 4, "B-1", "B");
        sheet.CreateRow(5).CreateCell(0).SetCellValue("SIGN");
        CreateGroupedLine(sheet, 6, "B-2", "B");
        CreateGroupedLine(sheet, 7, "B-3", "B");
        sheet.CreateRow(8).CreateCell(0).SetCellValue("GROUP");
        sheet.CreateRow(9).CreateCell(0).SetCellValue("SIGN");
        sheet.CreateRow(10).CreateCell(0).SetCellValue("END");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static void CreateLine(ISheet sheet, int rowIndex, string code, int quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(quantity);
    }

    private static void CreateGroupedLine(ISheet sheet, int rowIndex, string code, string group)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(group);
    }

    private static void WriteRow(ISheet sheet, int rowIndex, params string[] values)
    {
        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
        for (var index = 0; index < values.Length; index++)
            row.CreateCell(index).SetCellValue(values[index]);
    }

    private sealed class PageDocument
    {
        public List<PageLine> Lines { get; } = new();
    }

    private sealed class PageLine
    {
        public string Code { get; set; }

        public int Quantity { get; set; }

    }

    private sealed class GroupPageDocument
    {
        public List<GroupPageLine> Lines { get; } = new();
    }

    private sealed class GroupPageLine
    {
        public string Code { get; set; }

        public string Group { get; set; }
    }
}
