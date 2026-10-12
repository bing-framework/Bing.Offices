using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI Workbook 批次与失败工作簿的组合能力。
/// </summary>
public sealed class NpoiWorkbookBatchFailureWorkbookTest
{
    /// <summary>
    /// NPOI 的 XLS 和 XLSX 批次应按物理工作表顺序交付，并生成失败工作簿。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ImportWorkbookBatches_ShouldDeliverPhysicalOrderAndFailureWorkbook(ExcelFormat format)
    {
        using var source = CreateWorkbook(format);
        using var failure = new MemoryStream();
        var request = CreateRequest(failure);
        using var provider = CreateProvider();
        var importer = provider.GetRequiredService<IExcelImporter>();
        var batches = new List<ExcelWorkbookImportBatch>();
        var callbackSawFailureWorkbook = false;

        var summary = provider.GetRequiredService<IExcelWorkbookBatchImporter>()
            .ImportWorkbookBatches(source, request, batch =>
            {
                callbackSawFailureWorkbook = failure.Length > 0;
                batches.Add(batch);
            });

        var capabilities = Assert.IsAssignableFrom<IExcelProviderFeatureDescriptor>(importer);
        Assert.True(capabilities.SupportsBatchImport);
        Assert.True(capabilities.Supports(ExcelProviderFeatures.MultiSheetBatchImport));
        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(1, summary.ErrorCount);
        Assert.Equal(new[] { "Second", "First" }, batches.Select(batch => batch.SheetName));
        Assert.Equal(2, batches.Select(batch => batch.BatchNumber).Distinct().Count());
        Assert.Equal("First", Assert.IsType<Row>(Assert.Single(batches[1].Items)).Code);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, Assert.Single(batches[0].Errors).Code);
        Assert.Equal(2, Assert.Single(batches[0].Errors).RowIndex);
        Assert.True(callbackSawFailureWorkbook);
        Assert.True(failure.Length > 0);
        failure.Position = 0;
        using var failureWorkbook = WorkbookFactory.Create(failure);
        Assert.Contains("_ImportErrors", Enumerable.Range(0, failureWorkbook.NumberOfSheets)
            .Select(failureWorkbook.GetSheetName));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 异步批次回调取消后应停止交付并保留调用方输入流所有权。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_CallbackCancellation_ShouldPreserveSource()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx);
        using var cancellation = new CancellationTokenSource();
        using var provider = CreateProvider();
        var request = CreateRequest();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            provider.GetRequiredService<IExcelWorkbookBatchImporter>().ImportWorkbookBatchesAsync(
                source, request, (batch, token) =>
                {
                    cancellation.Cancel();
                    token.ThrowIfCancellationRequested();
                    return Task.CompletedTask;
                }, cancellation.Token));

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 没有导入错误时不应写出失败工作簿。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_WithoutErrors_ShouldNotWriteFailureWorkbook()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx, invalidSecondRow: false);
        using var failure = new MemoryStream();
        using var provider = CreateProvider();
        var summary = provider.GetRequiredService<IExcelWorkbookBatchImporter>()
            .ImportWorkbookBatches(source, CreateRequest(failure), _ => { });

        Assert.True(summary.IsSuccess);
        Assert.Equal(0, failure.Length);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建 NPOI Provider 服务容器。
    /// </summary>
    /// <returns>已完成构建的服务提供器。</returns>
    private static ServiceProvider CreateProvider()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        return services.BuildServiceProvider();
    }

    /// <summary>
    /// 创建失败工作簿测试请求。
    /// </summary>
    /// <param name="failure">失败工作簿目标流。</param>
    /// <param name="invalidSecondRow">是否将第二个工作表的数量写为非法文本。</param>
    /// <returns>配置完成的 Workbook 导入请求。</returns>
    private static ExcelWorkbookImportRequest<BatchWorkbook> CreateRequest(Stream failure = null)
    {
        return ExcelImport.Workbook<BatchWorkbook>(builder =>
        {
            if (failure != null)
            {
                builder.FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                    Destination = failure
                });
            }
            builder.Sheet<Row>("First", root => root.First,
                sheet => sheet.BatchSize(1).Validate(ExcelValidationFailureMode.Continue));
            builder.Sheet<Row>("Second", root => root.Second,
                sheet => sheet.BatchSize(1).Validate(ExcelValidationFailureMode.Continue));
        });
    }

    /// <summary>
    /// 创建物理顺序为 Second、First 的测试工作簿。
    /// </summary>
    /// <param name="format">输出工作簿格式。</param>
    /// <param name="invalidSecondRow">是否写入非法数量。</param>
    /// <returns>定位到起始位置的工作簿流。</returns>
    private static MemoryStream CreateWorkbook(ExcelFormat format, bool invalidSecondRow = true)
    {
        IWorkbook workbook = format == ExcelFormat.Xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var second = workbook.CreateSheet("Second");
        second.CreateRow(0).CreateCell(0).SetCellValue("Code");
        second.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        second.CreateRow(1).CreateCell(0).SetCellValue("Second");
        second.GetRow(1).CreateCell(1).SetCellValue(invalidSecondRow ? "bad" : "2");
        var first = workbook.CreateSheet("First");
        first.CreateRow(0).CreateCell(0).SetCellValue("Code");
        first.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        first.CreateRow(1).CreateCell(0).SetCellValue("First");
        first.GetRow(1).CreateCell(1).SetCellValue("1");
        var output = new MemoryStream();
        workbook.Write(output, true);
        workbook.Close();
        output.Position = 0;
        return output;
    }

    /// <summary>
    /// Workbook 批次测试根模型。
    /// </summary>
    private sealed class BatchWorkbook
    {
        /// <summary>
        /// 获取 First 工作表实体集合。
        /// </summary>
        public List<Row> First { get; } = new();

        /// <summary>
        /// 获取 Second 工作表实体集合。
        /// </summary>
        public List<Row> Second { get; } = new();
    }

    /// <summary>
    /// Workbook 批次测试行模型。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
