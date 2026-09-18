using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Csv;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 Async API 使用真实 Stream 异步入口并保持同步业务合同。
/// </summary>
[Collection("Excel Async staging")]
public sealed class AsyncPipelineTest
{
    /// <summary>
    /// 验证 Excel 异步入口使用异步外部流，并与同步导入结果一致。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_ShouldUseAsyncOuterStream_AndMatchSyncImport()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        var rows = new[] { new ExcelRow { Code = "A", Count = 1 }, new ExcelRow { Code = "B", Count = 2 } };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1", rows));
        var importRequest = ExcelImport.Workbook<ExcelWorkbook>(workbook => workbook.Sheet("Sheet1",
            root => root.Rows));

        using var syncOutput = new MemoryStream();
        exporter.Export(exportRequest, syncOutput);
        var source = syncOutput.ToArray();
        using var asyncInput = new AsyncOnlyReadStream(source);
        var asyncImport = await importer.ImportAsync(asyncInput, importRequest);

        using var asyncOutput = new AsyncOnlyWriteStream();
        await exporter.ExportAsync(exportRequest, asyncOutput);

        Assert.Equal(rows.Select(row => row.Code), asyncImport.Workbook.Rows.Select(row => row.Code));
        Assert.NotEmpty(asyncOutput.ToArray());
        Assert.True(asyncInput.AsyncReadCount > 0);
        Assert.Equal(0, asyncInput.SyncReadCount);
        Assert.True(asyncOutput.AsyncWriteCount > 0);
        Assert.Equal(0, asyncOutput.SyncWriteCount);
    }

    /// <summary>
    /// 验证 Excel 异步暂存策略能够生成输出并清理文件。
    /// </summary>
    /// <param name="strategyName">暂存策略名称。</param>
    [Theory]
    [InlineData("Memory")]
    [InlineData("TempFile")]
    [InlineData("Hybrid")]
    public async Task ExcelAsync_StagingStrategies_ShouldProduceOutputAndCleanFiles(string strategyName)
    {
        var strategy = Enum.Parse<NpoiAsyncStagingStrategy>(strategyName);
        var stagingDirectory = CreateStagingDirectory();
        var exporter = new NpoiExcelExporter(new DefaultFileExportCommitter(),
            new NpoiAsyncStagingFactory(strategy, hybridThresholdBytes: 1, directory: stagingDirectory));
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            Enumerable.Range(0, 100).Select(index => new ExcelRow { Code = $"C-{index}", Count = index })));
        using var output = new MemoryStream();

        try
        {
            await exporter.ExportAsync(request, output);

            Assert.NotEmpty(output.ToArray());
            Assert.Empty(Directory.GetFiles(stagingDirectory));
        }
        finally
        {
            Directory.Delete(stagingDirectory, true);
        }
    }

    /// <summary>
    /// 验证混合暂存会在写入前迁移并刷新数据。
    /// </summary>
    [Fact]
    public async Task HybridStaging_ShouldMigrateDuringWriteBeforeFlush()
    {
        var directory = CreateStagingDirectory();
        var staging = new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.Hybrid,
            hybridThresholdBytes: 4, directory: directory).Create("bing-offices-excel-async-");
        try
        {
            staging.WriteStream.Write(Encoding.UTF8.GetBytes("12345"));

            Assert.Single(Directory.GetFiles(directory));
            await staging.FlushAsync(CancellationToken.None);
            using var output = new MemoryStream();
            await staging.CopyToAsync(output, CancellationToken.None);
            Assert.Equal("12345", Encoding.UTF8.GetString(output.ToArray()));
        }
        finally
        {
            staging.Dispose();
            Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证混合暂存迁移期间会保留流位置并回填数据。
    /// </summary>
    [Fact]
    public async Task HybridStaging_ShouldPreserveSeekPositionAndBackpatchDuringMigration()
    {
        var directory = CreateStagingDirectory();
        var staging = new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.Hybrid,
            hybridThresholdBytes: 8, directory: directory).Create("bing-offices-excel-async-");
        try
        {
            staging.WriteStream.Write(Encoding.UTF8.GetBytes("abcdefgh"));
            staging.WriteStream.Position = 3;
            staging.WriteStream.Write(Encoding.UTF8.GetBytes("123456"));

            Assert.Single(Directory.GetFiles(directory));
            Assert.Equal(9, staging.WriteStream.Position);

            await staging.FlushAsync(CancellationToken.None);
            using var output = new MemoryStream();
            await staging.CopyToAsync(output, CancellationToken.None);

            Assert.Equal("abc123456", Encoding.UTF8.GetString(output.ToArray()));
        }
        finally
        {
            staging.Dispose();
            Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证混合暂存长度超过阈值时会迁移数据并保留流位置。
    /// </summary>
    [Fact]
    public async Task HybridStaging_SetLengthBeyondThreshold_ShouldMigrateAndPreservePosition()
    {
        var directory = CreateStagingDirectory();
        var staging = new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.Hybrid,
            hybridThresholdBytes: 4, directory: directory).Create("bing-offices-excel-async-");
        try
        {
            staging.WriteStream.Write(new byte[] { 1, 2, 3 });
            staging.WriteStream.Position = 1;
            staging.WriteStream.SetLength(8);

            Assert.Single(Directory.GetFiles(directory));
            Assert.Equal(1, staging.WriteStream.Position);
            Assert.Equal(8, staging.WriteStream.Length);

            staging.WriteStream.WriteByte(9);
            await staging.FlushAsync(CancellationToken.None);
            using var output = new MemoryStream();
            await staging.CopyToAsync(output, CancellationToken.None);

            Assert.Equal(new byte[] { 1, 9, 3, 0, 0, 0, 0, 0 }, output.ToArray());
        }
        finally
        {
            staging.Dispose();
            Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证 Excel 异步文件导出预先取消时会保留目标且不留下暂存文件。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_FileExport_PreCanceled_ShouldPreserveTargetAndLeaveNoStaging()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-excel-target-{Guid.NewGuid():N}.xlsx");
        var original = Encoding.UTF8.GetBytes("before");
        File.WriteAllBytes(path, original);
        var before = GetStagingFiles("bing-offices-excel-async-");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try
        {
            var exporter = new NpoiExcelExporter(new DefaultFileExportCommitter(),
                new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.TempFile));
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
                new[] { new ExcelRow { Code = "A", Count = 1 } }));

            await Assert.ThrowsAsync<OperationCanceledException>(() => exporter.ExportToFileAsync(request, path,
                cancellation.Token));

            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(before, GetStagingFiles("bing-offices-excel-async-"));
            Assert.Empty(GetAtomicTempFiles(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// 验证 Excel 异步文件导出写入期间取消时会保留目标并清理暂存。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_FileExport_MidWriteCancellation_ShouldPreserveTargetAndCleanStaging()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-excel-target-{Guid.NewGuid():N}.xlsx");
        var original = Encoding.UTF8.GetBytes("before");
        File.WriteAllBytes(path, original);
        var before = GetStagingFiles("bing-offices-excel-async-");
        using var cancellation = new CancellationTokenSource();
        var stagingFactory = new BlockingStagingFactory();
        try
        {
            var exporter = new NpoiExcelExporter(new DefaultFileExportCommitter(), stagingFactory);
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
                new[] { new ExcelRow { Code = "A", Count = 1 } }));
            var exportTask = exporter.ExportToFileAsync(request, path, cancellation.Token);
            await stagingFactory.Started.WaitAsync(TimeSpan.FromSeconds(5));
            cancellation.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() => exportTask);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(before, GetStagingFiles("bing-offices-excel-async-"));
            Assert.Empty(GetAtomicTempFiles(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// 验证异步文件导出预先取消时不会创建目标文件。
    /// </summary>
    [Fact]
    public async Task AsyncFileExport_PreCanceled_ShouldNotCreateTarget()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-async-{Guid.NewGuid():N}.csv");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() => exporter.ExportToFileAsync(
            new[] { new CsvRow { Code = "cancelled", Count = 1 } }, path,
            cancellationToken: cancellation.Token));
        Assert.False(File.Exists(path));
    }

    /// <summary>
    /// 验证 CSV 异步操作预先取消时会保留取消异常。
    /// </summary>
    [Fact]
    public async Task CsvAsync_PreCanceled_ShouldPreserveCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() => exporter.ExportAsync(
            new[] { new CsvRow { Code = "cancelled", Count = 1 } }, new MemoryStream(),
            cancellationToken: cancellation.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() => importer.ImportAsync<CsvRow>(
            new MemoryStream(Encoding.UTF8.GetBytes("Code,Count\r\nA,1\r\n")),
            cancellationToken: cancellation.Token));
    }

    /// <summary>
    /// 验证 CSV 异步读取和写入期间取消时会传播取消异常。
    /// </summary>
    [Fact]
    public async Task CsvAsync_MidReadAndMidWriteCancellation_ShouldPropagate()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var bytes = Encoding.UTF8.GetBytes("Code,Count\r\nA,1\r\nB,2\r\n");
        using var readCancellation = new CancellationTokenSource();
        using var input = new AsyncOnlyReadStream(bytes, () => readCancellation.Cancel());
        await Assert.ThrowsAsync<OperationCanceledException>(() => importer.ImportAsync<CsvRow>(input,
            cancellationToken: readCancellation.Token));

        using var writeCancellation = new CancellationTokenSource();
        using var output = new AsyncOnlyWriteStream(() => writeCancellation.Cancel());
        await Assert.ThrowsAsync<OperationCanceledException>(() => exporter.ExportAsync(
            new[] { new CsvRow { Code = "A", Count = 1 }, new CsvRow { Code = "B", Count = 2 } }, output,
            cancellationToken: writeCancellation.Token));
        Assert.Equal(readCancellation.Token, input.LastReadCancellationToken);
        Assert.Equal(writeCancellation.Token, output.LastWriteCancellationToken);
    }

    /// <summary>
    /// 验证 CSV 异步阻塞读取和写入使用调用方的取消令牌。
    /// </summary>
    [Fact]
    public async Task CsvAsync_BlockingReadAndWrite_ShouldUseCallerCancellationToken()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();

        using var readCancellation = new CancellationTokenSource();
        using var input = new BlockingAsyncReadStream();
        var importTask = importer.ImportAsync<CsvRow>(input, cancellationToken: readCancellation.Token);
        await input.Started.WaitAsync(TimeSpan.FromSeconds(5));
        readCancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => importTask);
        Assert.Equal(readCancellation.Token, input.LastReadCancellationToken);

        using var writeCancellation = new CancellationTokenSource();
        using var output = new BlockingAsyncWriteStream();
        var exportTask = exporter.ExportAsync(new[] { new CsvRow { Code = "A", Count = 1 } }, output,
            cancellationToken: writeCancellation.Token);
        await output.Started.WaitAsync(TimeSpan.FromSeconds(5));
        writeCancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => exportTask);
        Assert.Equal(writeCancellation.Token, output.LastWriteCancellationToken);
    }

    /// <summary>
    /// 验证 CSV 异步错误结果与同步结果一致。
    /// </summary>
    [Fact]
    public async Task CsvAsync_ErrorResult_ShouldMatchSyncResult()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var bytes = Encoding.UTF8.GetBytes("Code,Count\r\n,not-a-number\r\n");
        using var syncSource = new MemoryStream(bytes, writable: false);
        using var asyncSource = new MemoryStream(bytes, writable: false);
        var sync = importer.Import<ValidatedCsvRow>(syncSource);
        var async = await importer.ImportAsync<ValidatedCsvRow>(asyncSource);

        Assert.Equal(sync.Items.Count, async.Items.Count);
        Assert.Equal(sync.Errors.Count, async.Errors.Count);
        Assert.Equal(sync.Errors.Select(error => error.Code), async.Errors.Select(error => error.Code));
    }

    /// <summary>
    /// 验证 CSV 异步映射、校验和转换器行为与同步流程一致。
    /// </summary>
    [Fact]
    public async Task CsvAsync_MappingValidationConverterDynamic_ShouldMatchSyncPipeline()
    {
        var converter = new AsyncParityCodeConverter();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddSingleton<IExcelValueConverter>(converter)
            .BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(AsyncParityRow.Code), ConverterName = "async-code" },
                new() { PropertyName = nameof(AsyncParityRow.Name) }
            }
        };
        var options = new CsvExportOptions<AsyncParityRow>
        {
            MappingConfiguration = mapping,
            DynamicColumns = new[] { "Extra" }
        };
        var source = new[]
        {
            new AsyncParityRow
            {
                Code = new AsyncParityCode("42"),
                Name = "valid",
                Values = new Dictionary<string, object> { ["Extra"] = "dynamic" }
            }
        };
        using var syncOutput = new MemoryStream();
        exporter.Export(source, syncOutput, options);
        var syncBytes = syncOutput.ToArray();
        using var asyncOutput = new AsyncOnlyWriteStream();
        await exporter.ExportAsync(source, asyncOutput, options);

        using var syncInput = new MemoryStream(syncBytes, writable: false);
        using var asyncInput = new AsyncOnlyReadStream(asyncOutput.ToArray());
        var sync = importer.Import<AsyncParityRow>(syncInput,
            new CsvImportOptions<AsyncParityRow> { MappingConfiguration = mapping });
        var async = await importer.ImportAsync<AsyncParityRow>(asyncInput,
            new CsvImportOptions<AsyncParityRow> { MappingConfiguration = mapping });

        var syncItem = Assert.Single(sync.Items);
        var asyncItem = Assert.Single(async.Items);
        Assert.Empty(sync.Errors);
        Assert.Empty(async.Errors);
        Assert.Equal(syncItem.Code.Value, asyncItem.Code.Value);
        Assert.Equal(syncItem.Name, asyncItem.Name);
        Assert.Equal(syncItem.Values["Extra"], asyncItem.Values["Extra"]);
    }

    /// <summary>
    /// 验证 Excel 异步文件提交失败时只观察一次相同异常。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_FileCommitFailure_ShouldObserveSameExceptionOnce()
    {
        var observer = new RecordingObserver();
        var exporter = new Bing.Offices.Exports.NpoiExcelExporter(exceptionObservers: new[] { observer });
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new ExcelRow { Code = "A", Count = 1 } }));
        var targetDirectory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MissingParent",
            Guid.NewGuid().ToString("N"), "target.xlsx");

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            exporter.ExportToFileAsync(request, targetDirectory));

        Assert.Single(observer.Exceptions);
        Assert.Same(exception, observer.Exceptions[0]);
    }

    /// <summary>
    /// 验证 Excel 异步失败工作簿会复制到异步目标。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_FailureWorkbook_ShouldCopyToAsyncDestination()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new FailureRow { Code = "BAD" } }));
        using var source = new MemoryStream();
        exporter.Export(exportRequest, source);
        source.Position = 0;
        using var failureDestination = new AsyncOnlyWriteStream();
        var importRequest = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureDestination
            })
            .Sheet("Sheet1", root => root.Rows));

        var result = await importer.ImportAsync(source, importRequest);

        Assert.NotEmpty(result.Errors);
        Assert.NotEmpty(failureDestination.ToArray());
        Assert.True(failureDestination.AsyncWriteCount > 0);
        Assert.Equal(0, failureDestination.SyncWriteCount);
    }

    /// <summary>
    /// 验证 Excel 异步失败工作簿的暂存策略会清理文件。
    /// </summary>
    /// <param name="strategyName">暂存策略名称。</param>
    [Theory]
    [InlineData("Memory")]
    [InlineData("TempFile")]
    [InlineData("Hybrid")]
    public async Task ExcelAsync_FailureWorkbook_StagingStrategies_ShouldCleanFiles(string strategyName)
    {
        var strategy = Enum.Parse<NpoiAsyncStagingStrategy>(strategyName);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new FailureRow { Code = "BAD" } }));
        using var source = new MemoryStream();
        exporter.Export(exportRequest, source);
        source.Position = 0;
        using var failureDestination = new MemoryStream();
        var stagingDirectory = CreateStagingDirectory();
        var importer = new NpoiExcelImporter(new NpoiAsyncStagingFactory(strategy,
            hybridThresholdBytes: 1, directory: stagingDirectory));
        var importRequest = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureDestination
            })
            .Sheet("Sheet1", root => root.Rows));

        try
        {
            var result = await importer.ImportAsync(source, importRequest);

            Assert.NotEmpty(result.Errors);
            Assert.NotEmpty(failureDestination.ToArray());
            Assert.Empty(Directory.GetFiles(stagingDirectory));
        }
        finally
        {
            Directory.Delete(stagingDirectory, true);
        }
    }

    /// <summary>
    /// 验证 Excel 异步暂存清理失败会附加到导出异常。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_StagingCleanupFailure_ShouldAttachToTranslatedExportException()
    {
        var exporter = new NpoiExcelExporter(new DefaultFileExportCommitter(),
            new FailingStagingFactory());
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new ExcelRow { Code = "A", Count = 1 } }));

        var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
            exporter.ExportAsync(request, new MemoryStream()));

        Assert.True(exception.Data.Contains("Bing.Offices.ExcelAsync.StagingCleanupException"));
        Assert.IsType<IOException>(exception.Data["Bing.Offices.ExcelAsync.StagingCleanupException"]);
    }

    /// <summary>
    /// 验证 Excel 异步失败工作簿清理失败会附加到导入异常。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_FailureWorkbookCleanupFailure_ShouldAttachToTranslatedImportException()
    {
        var exporter = new NpoiExcelExporter();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new FailureRow { Code = "BAD" } }));
        using var source = new MemoryStream();
        exporter.Export(exportRequest, source);
        source.Position = 0;
        var importer = new NpoiExcelImporter(new FailingStagingFactory());
        var importRequest = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = new MemoryStream()
            })
            .Sheet("Sheet1", root => root.Rows));

        var exception = await Assert.ThrowsAsync<BingOfficesImportException>(() =>
            importer.ImportAsync(source, importRequest));

        Assert.True(exception.Data.Contains("Bing.Offices.ExcelAsync.FailureStagingCleanupException"));
        Assert.IsType<IOException>(exception.Data["Bing.Offices.ExcelAsync.FailureStagingCleanupException"]);
    }

    /// <summary>
    /// 验证 Excel 异步阻塞读取和写入会传播取消异常。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_BlockingReadAndWrite_ShouldPropagateCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        var rows = new[] { new ExcelRow { Code = "A", Count = 1 } };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1", rows));
        var importRequest = ExcelImport.Workbook<ExcelWorkbook>(workbook => workbook.Sheet("Sheet1",
            root => root.Rows));

        using var source = new MemoryStream();
        exporter.Export(exportRequest, source);
        var sourceBytes = source.ToArray();
        using var readCancellation = new CancellationTokenSource();
        using var input = new BlockingAsyncReadStream(sourceBytes);
        var importTask = importer.ImportAsync(input, importRequest, readCancellation.Token);
        await input.Started.WaitAsync(TimeSpan.FromSeconds(5));
        readCancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => importTask);
        Assert.Equal(readCancellation.Token, input.LastReadCancellationToken);

        using var writeCancellation = new CancellationTokenSource();
        using var output = new BlockingAsyncWriteStream();
        var exportTask = exporter.ExportAsync(exportRequest, output, writeCancellation.Token);
        await output.Started.WaitAsync(TimeSpan.FromSeconds(5));
        writeCancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => exportTask);
        Assert.Equal(writeCancellation.Token, output.LastWriteCancellationToken);
    }

    /// <summary>
    /// 验证 Excel 异步失败工作簿的阻塞目标会传播取消异常。
    /// </summary>
    [Fact]
    public async Task ExcelAsync_FailureWorkbook_BlockingDestination_ShouldPropagateCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new FailureRow { Code = "BAD" } }));
        using var source = new MemoryStream();
        exporter.Export(exportRequest, source);
        source.Position = 0;
        var before = GetStagingFiles("bing-offices-failure-async-");
        using var failureDestination = new BlockingAsyncWriteStream();
        var importRequest = ExcelImport.Workbook<FailureWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureDestination
            })
            .Sheet("Sheet1", root => root.Rows));
        using var cancellation = new CancellationTokenSource();
        var importTask = importer.ImportAsync(source, importRequest, cancellation.Token);
        await failureDestination.Started.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => importTask);
        Assert.Equal(cancellation.Token, failureDestination.LastWriteCancellationToken);
        Assert.Equal(before, GetStagingFiles("bing-offices-failure-async-"));
    }

    /// <summary>
    /// 获取暂存文件。
    /// </summary>
    /// <param name="prefix">文本前缀。</param>
    /// <returns>生成的字符串。</returns>
    private static HashSet<string> GetStagingFiles(string prefix) =>
        Directory.GetFiles(Path.GetTempPath(), prefix + "*.tmp")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 创建暂存目录。
    /// </summary>
    /// <returns>生成的字符串。</returns>
    private static string CreateStagingDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-staging-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    /// <summary>
    /// 获取原子提交临时文件。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    /// <returns>生成的字符串。</returns>
    private static string[] GetAtomicTempFiles(string path)
    {
        var directory = Path.GetDirectoryName(path)!;
        var fileName = Path.GetFileName(path);
        return Directory.GetFiles(directory, fileName + ".*.tmp");
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class CsvRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ValidatedCsvRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [Bing.Offices.Attributes.ExcelRequired]
        public string Code { get; set; }
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class AsyncParityRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public AsyncParityCode Code { get; set; }

        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        [ExcelRequired]
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置值集合。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Values { get; set; } = new();
    }

    /// <summary>
    /// 表示异步往返测试中的编码数据。
    /// </summary>
    private sealed class AsyncParityCode
    {
        /// <summary>
        /// 初始化一个 <see cref="AsyncParityCode" /> 类型的实例。
        /// </summary>
        /// <param name="value">异步一致性测试使用的编码值。</param>
        public AsyncParityCode(string value) => Value = value;
        /// <summary>
         /// 获取异步一致性测试值。
        /// </summary>
        public string Value { get; }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class AsyncParityCodeConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "async-code";
        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(AsyncParityCode);
        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = new AsyncParityCode(((string)context.Value).Substring(3));
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = $"CV-{((AsyncParityCode)context.Value).Value}";
            return true;
        }
    }

    /// <summary>
    /// 记录测试场景中的通知事件。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 获取异常集合。
        /// </summary>
        public List<BingOfficesException> Exceptions { get; } = new();
        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ExcelRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class ExcelWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ExcelRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class FailureRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelRegex("^OK-")]
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class FailureWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<FailureRow> Rows { get; } = new();
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class AsyncOnlyReadStream : Stream
    {
        /// <summary>
        /// 承载异步读取测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 异步读取完成后执行的测试回调。
        /// </summary>
        private readonly Action _afterAsyncRead;

        /// <summary>
        /// 初始化一个 <see cref="AsyncOnlyReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="bytes">作为只读流内容的字节数组。</param>
        /// <param name="afterAsyncRead">异步读取完成后执行的可选回调。</param>
        public AsyncOnlyReadStream(byte[] bytes, Action afterAsyncRead = null)
        {
            _inner = new MemoryStream(bytes, writable: false);
            _afterAsyncRead = afterAsyncRead;
        }

        /// <summary>
        /// 获取或设置同步读取次数。
        /// </summary>
        public int SyncReadCount { get; private set; }
        /// <summary>
        /// 获取或设置异步读取次数。
        /// </summary>
        public int AsyncReadCount { get; private set; }
        /// <summary>
        /// 获取或设置最近一次读取取消令牌。
        /// </summary>
        public CancellationToken LastReadCancellationToken { get; private set; }
        /// <inheritdoc />
        public override bool CanRead => true;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流拒绝同步读取，以验证异步入口不会退回同步 IO。
        /// </remarks>
        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            throw new InvalidOperationException("同步 Read 不允许用于 Async API。");
        }

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            AsyncReadCount++;
            LastReadCancellationToken = cancellationToken;
            return ReadAsyncCore(buffer, offset, count, cancellationToken);
        }

        /// <summary>
        /// 读取异步核心。
        /// </summary>
        /// <param name="buffer">数据缓冲区。</param>
        /// <param name="offset">流偏移量。</param>
        /// <param name="count">数据项数量。</param>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        /// <returns>计算得到的数值。</returns>
        private async Task<int> ReadAsyncCore(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken);
            _afterAsyncRead?.Invoke();
            return read;
        }

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            AsyncReadCount++;
            LastReadCancellationToken = cancellationToken;
            return ReadMemoryAsyncCore(buffer, cancellationToken);
        }

        /// <summary>
        /// 读取内存异步核心。
        /// </summary>
        /// <param name="buffer">数据缓冲区。</param>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        /// <returns>计算得到的数值。</returns>
        private async ValueTask<int> ReadMemoryAsyncCore(Memory<byte> buffer, CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, cancellationToken);
            _afterAsyncRead?.Invoke();
            return read;
        }

        /// <inheritdoc />
        public override void Flush() { }
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class AsyncOnlyWriteStream : Stream
    {
        /// <summary>
        /// 承载异步写入测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 异步写入完成后执行的测试回调。
        /// </summary>
        private readonly Action _afterAsyncWrite;

        /// <summary>
        /// 指示同步刷新是否应抛出异常。
        /// </summary>
        private readonly bool _throwOnSyncFlush;

        /// <summary>
        /// 初始化一个 <see cref="AsyncOnlyWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="afterAsyncWrite">异步写入完成后执行的可选回调。</param>
        /// <param name="throwOnSyncFlush">是否在同步刷新时抛出异常。</param>
        public AsyncOnlyWriteStream(Action afterAsyncWrite = null, bool throwOnSyncFlush = false)
        {
            _afterAsyncWrite = afterAsyncWrite;
            _throwOnSyncFlush = throwOnSyncFlush;
        }

        /// <summary>
        /// 获取或设置同步写入次数。
        /// </summary>
        public int SyncWriteCount { get; private set; }
        /// <summary>
        /// 获取或设置同步刷新次数。
        /// </summary>
        public int SyncFlushCount { get; private set; }
        /// <summary>
        /// 获取或设置异步写入次数。
        /// </summary>
        public int AsyncWriteCount { get; private set; }
        /// <summary>
        /// 获取或设置最近一次写入取消令牌。
        /// </summary>
        public CancellationToken LastWriteCancellationToken { get; private set; }
        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => true;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流不支持读取。
        /// </remarks>
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流不支持读取。
        /// </remarks>
        public override int Read(Span<byte> buffer) => throw new NotSupportedException();
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流拒绝同步写入，以验证异步入口不会退回同步 IO。
        /// </remarks>
        public override void Write(byte[] buffer, int offset, int count)
        {
            SyncWriteCount++;
            throw new InvalidOperationException("同步 Write 不允许用于 Async API。");
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            AsyncWriteCount++;
            LastWriteCancellationToken = cancellationToken;
            return WriteAsyncCore(buffer, offset, count, cancellationToken);
        }

        /// <summary>
        /// 写入异步核心。
        /// </summary>
        /// <param name="buffer">数据缓冲区。</param>
        /// <param name="offset">流偏移量。</param>
        /// <param name="count">数据项数量。</param>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        private async Task WriteAsyncCore(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            await _inner.WriteAsync(buffer, offset, count, cancellationToken);
            _afterAsyncWrite?.Invoke();
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWriteCount++;
            LastWriteCancellationToken = cancellationToken;
            return WriteMemoryAsyncCore(buffer, cancellationToken);
        }

        /// <summary>
        /// 写入内存异步核心。
        /// </summary>
        /// <param name="buffer">数据缓冲区。</param>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        private async ValueTask WriteMemoryAsyncCore(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken)
        {
            await _inner.WriteAsync(buffer, cancellationToken);
            _afterAsyncWrite?.Invoke();
        }

        /// <summary>
        /// 执行到数组。
        /// </summary>
        /// <returns>生成的字节内容。</returns>
        public byte[] ToArray() => _inner.ToArray();
        /// <inheritdoc />
        public override void Flush()
        {
            SyncFlushCount++;
            if (_throwOnSyncFlush)
                throw new InvalidOperationException("同步 Flush 不允许用于 Async API。");
        }
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供测试场景使用的工厂替身。
    /// </summary>
    private sealed class BlockingStagingFactory : INpoiAsyncStagingFactory
    {
        /// <summary>
        /// 通知 staging 复制已开始的完成源。
        /// </summary>
        private readonly TaskCompletionSource<bool> _started =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 获取是否已开始。
        /// </summary>
        public Task Started => _started.Task;

        /// <inheritdoc />
        public INpoiAsyncStaging Create(string prefix) => new BlockingStaging(_started);
    }

    /// <summary>
    /// 提供会阻塞写入以验证异步 staging 边界的替身。
    /// </summary>
    private sealed class BlockingStaging : INpoiAsyncStaging
    {
        /// <summary>
        /// 缓冲暂存写入内容的内存流。
        /// </summary>
        private readonly MemoryStream _buffer = new();

        /// <summary>
        /// 通知 staging 复制已开始的完成源。
        /// </summary>
        private readonly TaskCompletionSource<bool> _started;

        /// <summary>
        /// 初始化一个 <see cref="BlockingStaging" /> 类型的实例。
        /// </summary>
        /// <param name="started">用于通知 staging 复制已开始的完成源。</param>
        public BlockingStaging(TaskCompletionSource<bool> started) => _started = started;

        /// <summary>
        /// 获取写入流。
        /// </summary>
        public Stream WriteStream => _buffer;

        /// <inheritdoc />
        public Task FlushAsync(CancellationToken cancellationToken) => _buffer.FlushAsync(cancellationToken);

        /// <inheritdoc />
        /// <remarks>
        /// 复制开始后保持等待，直到收到取消，以验证调用方能够传播取消信号。
        /// </remarks>
        public Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
        {
            _started.TrySetResult(true);
            return WaitForCancellationAsync(cancellationToken);
        }

        /// <inheritdoc />
        public void Dispose() => _buffer.Dispose();

        /// <summary>
        /// 等待取消令牌触发。
        /// </summary>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        private static async Task WaitForCancellationAsync(CancellationToken cancellationToken)
            => await Task.Delay(Timeout.Infinite, cancellationToken);
    }

    /// <summary>
    /// 提供测试场景使用的工厂替身。
    /// </summary>
    private sealed class FailingStagingFactory : INpoiAsyncStagingFactory
    {
        /// <inheritdoc />
        public INpoiAsyncStaging Create(string prefix) => new FailingStaging();
    }

    /// <summary>
    /// 提供会抛出异常以验证 staging 失败路径的替身。
    /// </summary>
    private sealed class FailingStaging : INpoiAsyncStaging
    {
        /// <summary>
        /// 缓冲失败暂存写入内容的内存流。
        /// </summary>
        private readonly MemoryStream _buffer = new();

        /// <summary>
        /// 获取写入流。
        /// </summary>
        public Stream WriteStream => _buffer;

        /// <inheritdoc />
        public Task FlushAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        /// <remarks>
        /// 始终返回 staging 复制失败异常，用于验证失败传播。
        /// </remarks>
        public Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
            => Task.FromException(new InvalidOperationException("staging 复制失败"));

        /// <inheritdoc />
        /// <remarks>
        /// 始终抛出清理异常，用于验证失败路径。
        /// </remarks>
        public void Dispose() => throw new IOException("staging 清理失败");
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class BlockingAsyncReadStream : Stream
    {
        /// <summary>
        /// 承载阻塞异步读取测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 通知异步读取已开始的完成源。
        /// </summary>
        private readonly TaskCompletionSource<bool> _started = new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 初始化一个 <see cref="BlockingAsyncReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="content">作为只读流内容的可选字节数组。</param>
        public BlockingAsyncReadStream(byte[] content = null)
            => _inner = new MemoryStream(content ?? Array.Empty<byte>(), writable: false);

        /// <summary>
        /// 获取是否已开始。
        /// </summary>
        public Task Started => _started.Task;
        /// <summary>
        /// 获取或设置最近一次读取取消令牌。
        /// </summary>
        public CancellationToken LastReadCancellationToken { get; private set; }
        /// <inheritdoc />
        public override bool CanRead => true;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流拒绝同步读取，以验证异步入口能够传播取消信号。
        /// </remarks>
        public override int Read(byte[] buffer, int offset, int count)
            => throw new InvalidOperationException("Async-only blocking stream received sync Read.");
        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            LastReadCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return WaitForCancellationAsync(cancellationToken);
        }
        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            LastReadCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return new ValueTask<int>(WaitForCancellationAsync(cancellationToken));
        }
        /// <summary>
        /// 等待取消令牌触发。
        /// </summary>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        /// <returns>计算得到的数值。</returns>
        private static async Task<int> WaitForCancellationAsync(CancellationToken cancellationToken)
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return 0;
        }
        /// <inheritdoc />
        public override void Flush() { }
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流不支持同步写入。
        /// </remarks>
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class BlockingAsyncWriteStream : Stream
    {
        /// <summary>
        /// 通知异步写入已开始的完成源。
        /// </summary>
        private readonly TaskCompletionSource<bool> _started = new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 获取是否已开始。
        /// </summary>
        public Task Started => _started.Task;
        /// <summary>
        /// 获取或设置最近一次写入取消令牌。
        /// </summary>
        public CancellationToken LastWriteCancellationToken { get; private set; }
        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流不支持同步读取。
        /// </remarks>
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        /// <remarks>
        /// 测试用流拒绝同步写入，以验证异步入口能够传播取消信号。
        /// </remarks>
        public override void Write(byte[] buffer, int offset, int count)
            => throw new InvalidOperationException("Async-only blocking stream received sync Write.");
        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            LastWriteCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return WaitForCancellationAsync(cancellationToken);
        }
        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            LastWriteCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return new ValueTask(WaitForCancellationAsync(cancellationToken));
        }
        /// <summary>
        /// 等待取消令牌触发。
        /// </summary>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
        private static async Task WaitForCancellationAsync(CancellationToken cancellationToken)
            => await Task.Delay(Timeout.Infinite, cancellationToken);
        /// <inheritdoc />
        public override void Flush() { }
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}

/// <summary>
/// 表示异步边界测试使用的测试夹具。
/// </summary>
[CollectionDefinition("Excel Async staging", DisableParallelization = true)]
public sealed class ExcelAsyncStagingCollection
{
}
