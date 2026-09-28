using Bing.Offices;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Starter.Provider;
using Xunit;

namespace Starter.Provider.Tests;

/// <summary>
/// 最小 Provider 的真实工作簿和契约测试。
/// </summary>
public sealed class StarterExcelExporterTests
{
    /// <summary>
    /// 验证同步和异步入口输出完整的多工作表数据。
    /// </summary>
    /// <param name="async">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_ShouldWriteAllSheets(bool async)
    {
        var exporter = new StarterExcelExporter();
        using var output = new MemoryStream();

        if (async) await exporter.ExportAsync(Request(), output);
        else exporter.Export(Request(), output);

        Assert.True(output.CanWrite);
        AssertWorkbook(output.ToArray());
    }

    /// <summary>
    /// 验证异步入口调用目标流的异步写入。
    /// </summary>
    [Fact]
    public async Task ExportAsync_ShouldUseDestinationAsyncIo()
    {
        using var output = new AsyncOnlyStream();

        await new StarterExcelExporter().ExportAsync(Request(), output);

        Assert.True(output.AsyncWrites > 0);
        AssertWorkbook(output.ToArray());
    }

    /// <summary>
    /// 验证声明的能力与实际不支持范围一致。
    /// </summary>
    [Fact]
    public void Capabilities_ShouldMatchImplementedBoundary()
    {
        IExcelProviderFeatureDescriptor provider = new StarterExcelExporter();
        Assert.Equal("StarterClosedXml", provider.ProviderName);
        Assert.True(provider.Supports(ExcelProviderCapabilities.List | ExcelProviderCapabilities.Workbook
            | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx));
        Assert.False(provider.Supports(ExcelProviderCapabilities.Entity));
        Assert.False(provider.Supports(ExcelProviderCapabilities.Template));
        Assert.False(provider.Supports(ExcelProviderFeatures.StreamingWorkbookCreation));
        Assert.Equal(new[] { ExcelFormat.Xlsx }, provider.WriteFormats);
        Assert.Empty(provider.ReadFormats);
        Assert.True(provider.SupportsCompleteWorkbookExport);
        Assert.False(provider.SupportsCompleteWorkbookImport);
        Assert.True(provider.SupportsTrueAsyncIo);
    }

    /// <summary>
    /// 验证宿主注册的提交器和观察器由 DI 注入。
    /// </summary>
    [Fact]
    public void AddStarterExcelProvider_ShouldPreserveHostServices()
    {
        var services = new ServiceCollection();
        var committer = new RecordingCommitter();
        var observer = new RecordingObserver();
        services.AddSingleton<IFileExportCommitter>(committer);
        services.AddSingleton<IBingOfficesExceptionObserver>(observer);
        services.AddStarterExcelProvider();
        using var provider = services.BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();

        Assert.IsType<StarterExcelExporter>(exporter);
        Assert.Equal(typeof(StarterExcelExporter),
            provider.GetRequiredService<IExcelProviderFeatureDescriptor>().GetType());
        Assert.Same(committer, provider.GetRequiredService<IFileExportCommitter>());
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            exporter.Export(ExcelExport.Workbook(book => book.Format(ExcelFormat.Xls)
                .AddSheet("Rows", new[] { new Row { Id = 1 } })), output));
        Assert.Equal(1, observer.Count);
        Assert.Empty(output.ToArray());

        var path = Path.Combine(Path.GetTempPath(), $"starter-provider-di-{Guid.NewGuid():N}.xlsx");
        try
        {
            exporter.ExportToFile(Request(), path);
            Assert.Equal(1, committer.CommitCalls);
            AssertWorkbook(File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    /// <summary>
    /// 验证未实现的格式、模板和高级布局在输出前被拒绝。
    /// </summary>
    [Fact]
    public void UnsupportedRequest_ShouldLeaveDestinationUntouched()
    {
        var exporter = new StarterExcelExporter();
        using var output = new MemoryStream();
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var requests = new[]
        {
            ExcelExport.Workbook(book => book.Format(ExcelFormat.Xls)
                .AddSheet("Rows", new[] { new Row { Id = 1 } })),
            ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
                .AddSheet("Rows", new[] { new Row { Id = 1 } })),
            ExcelExport.Workbook(book => book.AddSheet("Rows", new[] { new Row { Id = 1 } },
                sheet => sheet.FreezePane(new ExcelFreezePaneDefinition { Rows = 1 })))
        };

        foreach (var request in requests)
        {
            var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                exporter.Export(request, output));
            Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
            Assert.Equal(BingOfficesOperation.Export, exception.Operation);
            Assert.Equal("StarterClosedXml", exception.Provider);
            Assert.Empty(output.ToArray());
        }
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证未声明 Entity 时公共扩展入口结构化拒绝。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectBeforeOutput()
    {
        IExcelExporter exporter = new StarterExcelExporter();
        var layout = ExcelEntity.Layout<Row>(builder => builder.Cell("Rows", "A1", row => row.Id));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            exporter.ExportEntity(new Row { Id = 1 }, layout, output));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证行数限制和属性读取失败不会写入调用方目标。
    /// </summary>
    [Fact]
    public void FailedBuild_ShouldPreserveDestinationAndObserveOnce()
    {
        var observer = new RecordingObserver();
        var exporter = new StarterExcelExporter(exceptionObservers: new[] { observer });
        using var output = new MemoryStream();
        var rows = Enumerable.Range(0, 10001).Select(id => new Row { Id = id });
        var limitRequest = ExcelExport.Workbook(book => book.AddSheet("Rows", rows));

        var limit = Assert.Throws<BingOfficesResourceLimitException>(() => exporter.Export(limitRequest, output));
        Assert.Equal(BingOfficesStage.Write, limit.Stage);
        Assert.Equal(1, observer.Count);
        Assert.Empty(output.ToArray());

        var failingRequest = ExcelExport.Workbook(book => book.AddSheet("Rows",
            new[] { new ThrowingRow() }));
        var failure = Assert.Throws<BingOfficesExportException>(() => exporter.Export(failingRequest, output));
        Assert.Equal(BingOfficesStage.Write, failure.Stage);
        Assert.Equal(2, observer.Count);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证同步与异步文件入口替换原有目标。
    /// </summary>
    /// <param name="async">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportToFile_ShouldReplaceExistingWorkbook(bool async)
    {
        var path = Path.Combine(Path.GetTempPath(), $"starter-provider-{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            var exporter = new StarterExcelExporter();
            if (async) await exporter.ExportToFileAsync(Request(), path);
            else exporter.ExportToFile(Request(), path);

            AssertWorkbook(File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    /// <summary>
    /// 验证预取消不修改已有目标文件。
    /// </summary>
    [Fact]
    public async Task PreCancelledExportToFile_ShouldPreserveExistingBytes()
    {
        var path = Path.Combine(Path.GetTempPath(), $"starter-provider-{Guid.NewGuid():N}.xlsx");
        var previous = new byte[] { 1, 2, 3 };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try
        {
            File.WriteAllBytes(path, previous);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                new StarterExcelExporter().ExportToFileAsync(Request(), path, cancellation.Token));
            Assert.Equal(previous, File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    /// <summary>
    /// 构造两个工作表的基础导出请求。
    /// </summary>
    /// <returns>工作簿请求。</returns>
    private static ExcelWorkbookExportRequest Request() => ExcelExport.Workbook(book => book
        .AddSheet("Orders", new[]
        {
            new Row { Id = 1, Name = "采购", Amount = 12.5m, Active = true },
            new Row { Id = 2, Name = "赠品", Amount = 0m, Active = false }
        })
        .AddSheet("Summary", new[] { new SummaryRow { Count = 2 } }));

    /// <summary>
    /// 校验完整工作簿结构和数据。
    /// </summary>
    /// <param name="bytes">工作簿字节。</param>
    private static void AssertWorkbook(byte[] bytes)
    {
        using var source = new MemoryStream(bytes);
        using var workbook = new XLWorkbook(source);
        Assert.Equal(2, workbook.Worksheets.Count);
        var orders = workbook.Worksheet("Orders");
        Assert.Equal(new[] { "Active", "Amount", "Id", "Name" },
            Enumerable.Range(1, 4).Select(index => orders.Cell(1, index).GetString()));
        Assert.True(orders.Cell(2, 1).GetBoolean());
        Assert.Equal(12.5, orders.Cell(2, 2).GetDouble());
        Assert.Equal(1, orders.Cell(2, 3).GetDouble());
        Assert.Equal("采购", orders.Cell(2, 4).GetString());
        Assert.False(orders.Cell(3, 1).GetBoolean());
        Assert.Equal(0, orders.Cell(3, 2).GetDouble());
        Assert.Equal(2, orders.Cell(3, 3).GetDouble());
        Assert.Equal("赠品", orders.Cell(3, 4).GetString());
        var summary = workbook.Worksheet("Summary");
        Assert.Equal("Count", summary.Cell(1, 1).GetString());
        Assert.Equal(2, summary.Cell(2, 1).GetDouble());
    }

    /// <summary>
    /// 导出明细数据。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置是否有效。
        /// </summary>
        public bool Active { get; set; }
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
        /// <summary>
        /// 获取或设置编号。
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 汇总数据。
    /// </summary>
    public sealed class SummaryRow
    {
        /// <summary>
        /// 获取或设置记录数。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 用于触发属性读取失败的数据。
    /// </summary>
    public sealed class ThrowingRow
    {
        /// <summary>
        /// 获取值时抛出异常。
        /// </summary>
        public int Value => throw new InvalidOperationException("getter failed");
    }

    /// <summary>
    /// 记录异常观察次数。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 获取观察次数。
        /// </summary>
        public int Count { get; private set; }

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Count++;
    }

    /// <summary>
    /// 用于确认 DI 保留宿主实例的提交器。
    /// </summary>
    private sealed class RecordingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 用于转发文件提交操作的默认提交器。
        /// </summary>
        private readonly DefaultFileExportCommitter _inner = new DefaultFileExportCommitter();

        /// <summary>
        /// 获取同步提交次数。
        /// </summary>
        public int CommitCalls { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        {
            CommitCalls++;
            _inner.Commit(path, write, cancellationToken, format);
        }

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format) =>
            _inner.CommitAsync(path, writeAsync, cancellationToken, format);
    }

    /// <summary>
    /// 禁止同步写入并记录异步写入次数。
    /// </summary>
    private sealed class AsyncOnlyStream : MemoryStream
    {
        /// <summary>
        /// 获取异步写入次数。
        /// </summary>
        public int AsyncWrites { get; private set; }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new InvalidOperationException("同步写入不允许。");

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncWrites++;
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWrites++;
            var bytes = buffer.ToArray();
            base.Write(bytes, 0, bytes.Length);
            return ValueTask.CompletedTask;
        }
    }
}
