using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML Workbook 批次与失败工作簿的组合能力。
/// </summary>
public sealed class ClosedXmlWorkbookBatchFailureWorkbookTest
{
    /// <summary>
    /// ClosedXML 批次应按物理工作表顺序交付，并生成带错误批注的失败工作簿。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_ShouldDeliverPhysicalOrderAndFailureWorkbook()
    {
        using var source = CreateWorkbook(invalidSecondRow: true);
        using var failure = new MemoryStream();
        using var provider = CreateProvider();
        var batches = new List<ExcelWorkbookImportBatch>();
        var callbackSawFailureWorkbook = false;

        var summary = provider.GetRequiredService<IExcelWorkbookBatchImporter>()
            .ImportWorkbookBatches(source, CreateRequest(failure), batch =>
            {
                callbackSawFailureWorkbook = failure.Length > 0;
                batches.Add(batch);
            });
        var importer = provider.GetRequiredService<IExcelImporter>();

        var capabilities = Assert.IsAssignableFrom<IExcelProviderFeatureDescriptor>(importer);
        Assert.True(capabilities.SupportsBatchImport);
        Assert.True(capabilities.Supports(ExcelProviderFeatures.MultiSheetBatchImport));
        Assert.False(summary.IsSuccess);
        Assert.Equal(new[] { "Second", "First" }, batches.Select(batch => batch.SheetName));
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, Assert.Single(batches[0].Errors).Code);
        Assert.True(callbackSawFailureWorkbook);
        Assert.True(failure.Length > 0);
        failure.Position = 0;
        using var failureWorkbook = new XLWorkbook(failure);
        Assert.Contains("_ImportErrors", failureWorkbook.Worksheets.Select(sheet => sheet.Name));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// ClosedXML 异步批次应保持失败工作簿先提交、批次后回调的顺序。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_ShouldDeliverAfterFailureWorkbook()
    {
        using var source = CreateWorkbook(invalidSecondRow: true);
        using var failure = new MemoryStream();
        using var provider = CreateProvider();
        var batches = new List<ExcelWorkbookImportBatch>();
        var callbackSawFailureWorkbook = false;

        var summary = await provider.GetRequiredService<IExcelWorkbookBatchImporter>()
            .ImportWorkbookBatchesAsync(source, CreateRequest(failure), (batch, token) =>
            {
                token.ThrowIfCancellationRequested();
                callbackSawFailureWorkbook = failure.Length > 0;
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.False(summary.IsSuccess);
        Assert.True(callbackSawFailureWorkbook);
        Assert.Equal(new[] { "Second", "First" }, batches.Select(batch => batch.SheetName));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 没有导入错误时不应写出失败工作簿。
    /// </summary>
    [Fact]
    public void ImportWorkbookBatches_WithoutErrors_ShouldNotWriteFailureWorkbook()
    {
        using var source = CreateWorkbook(invalidSecondRow: false);
        using var failure = new MemoryStream();
        using var provider = CreateProvider();

        var summary = provider.GetRequiredService<IExcelWorkbookBatchImporter>()
            .ImportWorkbookBatches(source, CreateRequest(failure), _ => { });

        Assert.True(summary.IsSuccess);
        Assert.Equal(0, failure.Length);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建 ClosedXML Provider 服务容器。
    /// </summary>
    /// <returns>已完成构建的服务提供器。</returns>
    private static ServiceProvider CreateProvider()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesClosedXml();
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
    /// <param name="invalidSecondRow">是否写入非法数量。</param>
    /// <returns>定位到起始位置的工作簿流。</returns>
    private static MemoryStream CreateWorkbook(bool invalidSecondRow)
    {
        using var workbook = new XLWorkbook();
        var second = workbook.Worksheets.Add("Second");
        second.Cell(1, 1).Value = "Code";
        second.Cell(1, 2).Value = "Quantity";
        second.Cell(2, 1).Value = "Second";
        second.Cell(2, 2).Value = invalidSecondRow ? "bad" : "2";
        var first = workbook.Worksheets.Add("First");
        first.Cell(1, 1).Value = "Code";
        first.Cell(1, 2).Value = "Quantity";
        first.Cell(2, 1).Value = "First";
        first.Cell(2, 2).Value = "1";
        var output = new MemoryStream();
        workbook.SaveAs(output);
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
