using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// ClosedXML Forward 每页尾部边界测试。
/// </summary>
public sealed class EntityForwardPageFooterTest
{
    /// <summary>
    /// 验证同步和异步 Forward 导入会跳过每页尾部及最终尾部。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageFooter(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());
        var batches = new List<ExcelEntityImportBatch<PageDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateLayout(), options, batches.Add);

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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageHeaderAndPageFooter(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterWorkbook());
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageHeaderWithPageSubtotalAndFooter(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterPageSubtotalWorkbook());
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageHeaderWithGroupSubtotalAndFooter(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterGroupSubtotalWorkbook());
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldSkipGappedMultiRowFinalFooter(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterGappedFinalFooterWorkbook());
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
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_PageHeaderAndFooter_ShouldStopAfterCallbackCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterWorkbook());
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_PageHeaderAndFooter_ShouldReadNonSeekableInput(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new NonSeekableReadStream(CreateHeaderFooterWorkbook());
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
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectPageHeaderAndPageFooterMisalignment()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHeaderFooterWorkbook(
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
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectMissingLastPageFooter()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(omitLastPageFooter: true));
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
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectDuplicatePageFooter()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(duplicatePageFooter: true));
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
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectEmptyLastPageFooter()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(emptyLastPage: true));
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldHonorPreCancellation(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());
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
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadPageSubtotalBeforePageFooter()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateSubtotalWorkbook());
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadGroupSubtotalBeforePageFooter(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateGroupSubtotalWorkbook());
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

    private static byte[] CreateWorkbook(bool omitLastPageFooter = false,
        bool duplicatePageFooter = false, bool emptyLastPage = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Quantity";
        CreateLine(sheet, 2, "A-1", 1);
        CreateLine(sheet, 3, "A-2", 2);
        sheet.Cell(5, 1).Value = "SIGN";
        sheet.Cell(6, 2).Value = 2;
        if (duplicatePageFooter)
            sheet.Cell(6, 1).Value = "SIGN";
        CreateLine(sheet, 7, "B-1", 3);
        CreateLine(sheet, 8, "B-2", 4);
        if (omitLastPageFooter)
            sheet.Cell(10, 1).Value = "END";
        else if (emptyLastPage)
        {
            sheet.Cell(7, 1).Value = "SIGN";
            sheet.Cell(8, 1).Value = "END";
        }
        else
        {
            sheet.Cell(10, 1).Value = "SIGN";
            sheet.Cell(11, 2).Value = 2;
            sheet.Cell(13, 1).Value = "END";
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterWorkbook(bool invalidRepeatedHeader = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "HDR";
        sheet.Cell(1, 2).Value = "采购单";
        sheet.Cell(2, 1).Value = "Code";
        sheet.Cell(2, 2).Value = "Quantity";
        CreateLine(sheet, 3, "A-1", 1);
        CreateLine(sheet, 4, "A-2", 2);
        sheet.Cell(5, 1).Value = "SIGN";
        sheet.Cell(6, 1).Value = invalidRepeatedHeader ? "WRONG" : "HDR";
        sheet.Cell(6, 2).Value = "采购单";
        sheet.Cell(7, 1).Value = "Code";
        sheet.Cell(7, 2).Value = "Quantity";
        CreateLine(sheet, 8, "B-1", 3);
        CreateLine(sheet, 9, "B-2", 4);
        sheet.Cell(10, 1).Value = "SIGN";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterPageSubtotalWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "HDR";
        sheet.Cell(1, 2).Value = "采购单";
        sheet.Cell(2, 1).Value = "Code";
        sheet.Cell(2, 2).Value = "Quantity";
        CreateLine(sheet, 3, "A-1", 1);
        CreateLine(sheet, 4, "A-2", 2);
        sheet.Cell(5, 1).Value = "PAGE";
        sheet.Cell(6, 1).Value = "SIGN";
        sheet.Cell(7, 1).Value = "HDR";
        sheet.Cell(7, 2).Value = "采购单";
        sheet.Cell(8, 1).Value = "Code";
        sheet.Cell(8, 2).Value = "Quantity";
        CreateLine(sheet, 9, "B-1", 3);
        CreateLine(sheet, 10, "B-2", 4);
        sheet.Cell(11, 1).Value = "SIGN";
        sheet.Cell(12, 1).Value = "END";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterGroupSubtotalWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "HDR";
        sheet.Cell(1, 2).Value = "采购单";
        sheet.Cell(2, 1).Value = "Code";
        sheet.Cell(2, 2).Value = "Group";
        CreateGroupedLine(sheet, 3, "A-1", "A");
        CreateGroupedLine(sheet, 4, "A-2", "A");
        sheet.Cell(5, 1).Value = "GROUP";
        sheet.Cell(6, 1).Value = "SIGN";
        sheet.Cell(7, 1).Value = "HDR";
        sheet.Cell(7, 2).Value = "采购单";
        sheet.Cell(8, 1).Value = "Code";
        sheet.Cell(8, 2).Value = "Group";
        CreateGroupedLine(sheet, 9, "B-1", "B");
        CreateGroupedLine(sheet, 10, "B-2", "B");
        sheet.Cell(11, 1).Value = "GROUP";
        sheet.Cell(12, 1).Value = "SIGN";
        sheet.Cell(13, 1).Value = "END";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static byte[] CreateHeaderFooterGappedFinalFooterWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "HDR";
        sheet.Cell(1, 2).Value = "采购单";
        sheet.Cell(2, 1).Value = "Code";
        sheet.Cell(2, 2).Value = "Quantity";
        CreateLine(sheet, 3, "A-1", 1);
        CreateLine(sheet, 4, "A-2", 2);
        sheet.Cell(5, 1).Value = "SIGN";
        sheet.Cell(6, 1).Value = "HDR";
        sheet.Cell(6, 2).Value = "采购单";
        sheet.Cell(7, 1).Value = "Code";
        sheet.Cell(7, 2).Value = "Quantity";
        CreateLine(sheet, 8, "B-1", 3);
        CreateLine(sheet, 9, "B-2", 4);
        sheet.Cell(10, 1).Value = "SIGN";
        sheet.Row(11);
        sheet.Cell(12, 1).Value = "END";
        sheet.Cell(13, 1).Value = "TAIL";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
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

    private static byte[] CreateSubtotalWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Quantity";
        CreateLine(sheet, 2, "A-1", 1);
        CreateLine(sheet, 3, "A-2", 2);
        sheet.Cell(5, 1).Value = "PAGE";
        sheet.Cell(6, 2).Value = 3;
        sheet.Cell(8, 1).Value = "SIGN";
        CreateLine(sheet, 9, "B-1", 3);
        CreateLine(sheet, 10, "B-2", 4);
        sheet.Cell(12, 1).Value = "SIGN";
        sheet.Cell(14, 1).Value = "END";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static byte[] CreateGroupSubtotalWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Group";
        CreateGroupedLine(sheet, 2, "A-1", "A");
        CreateGroupedLine(sheet, 3, "A-2", "A");
        sheet.Cell(4, 1).Value = "GROUP";
        CreateGroupedLine(sheet, 5, "B-1", "B");
        sheet.Cell(6, 1).Value = "SIGN";
        CreateGroupedLine(sheet, 7, "B-2", "B");
        CreateGroupedLine(sheet, 8, "B-3", "B");
        sheet.Cell(9, 1).Value = "GROUP";
        sheet.Cell(10, 1).Value = "SIGN";
        sheet.Cell(11, 1).Value = "END";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static void CreateLine(IXLWorksheet sheet, int rowNumber, string code, int quantity)
    {
        sheet.Cell(rowNumber, 1).Value = code;
        sheet.Cell(rowNumber, 2).Value = quantity;
    }

    private static void CreateGroupedLine(IXLWorksheet sheet, int rowNumber, string code, string group)
    {
        sheet.Cell(rowNumber, 1).Value = code;
        sheet.Cell(rowNumber, 2).Value = group;
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
