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

/// <summary>验证 Async API 使用真实 Stream 异步入口并保持同步业务合同。</summary>
[Collection("Excel Async staging")]
public sealed class AsyncPipelineTest
{
    [Fact]
    public async Task CsvAsync_ShouldUseAsyncReadWrite_AndMatchSyncResult()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var rows = new[] { new CsvRow { Code = "A", Count = 1 }, new CsvRow { Code = "B", Count = 2 } };

        using var syncOutput = new MemoryStream();
        exporter.Export(rows, syncOutput);
        var source = syncOutput.ToArray();

        using var asyncInput = new AsyncOnlyReadStream(source);
        var asyncResult = await importer.ImportAsync<CsvRow>(asyncInput);

        using var asyncOutput = new AsyncOnlyWriteStream();
        await exporter.ExportAsync(rows, asyncOutput);

        Assert.Equal(rows.Select(row => row.Code), asyncResult.Items.Select(row => row.Code));
        Assert.Equal(rows.Select(row => row.Count), asyncResult.Items.Select(row => row.Count));
        Assert.NotEmpty(asyncOutput.ToArray());
        Assert.True(asyncInput.AsyncReadCount > 0);
        Assert.Equal(0, asyncInput.SyncReadCount);
        Assert.True(asyncOutput.AsyncWriteCount > 0);
        Assert.Equal(0, asyncOutput.SyncWriteCount);
    }

    [Fact]
    public async Task CsvAsync_ShouldNotUseSynchronousFlushOnCallerStream()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        using var output = new AsyncOnlyWriteStream(throwOnSyncFlush: true);

        await exporter.ExportAsync(new[] { new CsvRow { Code = "A", Count = 1 } }, output);

        Assert.NotEmpty(output.ToArray());
        Assert.Equal(0, output.SyncFlushCount);
    }

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

    [Fact]
    public async Task ExcelAsync_FileCommitFailure_ShouldObserveSameExceptionOnce()
    {
        var observer = new RecordingObserver();
        var exporter = new Bing.Offices.Exports.NpoiExcelExporter(exceptionObservers: new[] { observer });
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1",
            new[] { new ExcelRow { Code = "A", Count = 1 } }));
        var targetDirectory = Path.GetTempPath();

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            exporter.ExportToFileAsync(request, targetDirectory));

        Assert.Single(observer.Exceptions);
        Assert.Same(exception, observer.Exceptions[0]);
    }

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

    private static HashSet<string> GetStagingFiles(string prefix) =>
        Directory.GetFiles(Path.GetTempPath(), prefix + "*.tmp")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static string CreateStagingDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-staging-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static string[] GetAtomicTempFiles(string path)
    {
        var directory = Path.GetDirectoryName(path)!;
        var fileName = Path.GetFileName(path);
        return Directory.GetFiles(directory, fileName + ".*.tmp");
    }

    private sealed class CsvRow
    {
        public string Code { get; set; }
        public int Count { get; set; }
    }

    private sealed class ValidatedCsvRow
    {
        [Bing.Offices.Attributes.ExcelRequired]
        public string Code { get; set; }
        public int Count { get; set; }
    }

    private sealed class AsyncParityRow
    {
        public AsyncParityCode Code { get; set; }

        [ExcelRequired]
        public string Name { get; set; }

        [DynamicColumn]
        public Dictionary<string, object> Values { get; set; } = new();
    }

    private sealed class AsyncParityCode
    {
        public AsyncParityCode(string value) => Value = value;
        public string Value { get; }
    }

    private sealed class AsyncParityCodeConverter : INamedExcelValueConverter
    {
        public string Name => "async-code";
        public bool CanConvert(Type propertyType) => propertyType == typeof(AsyncParityCode);
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = new AsyncParityCode(((string)context.Value).Substring(3));
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = $"CV-{((AsyncParityCode)context.Value).Value}";
            return true;
        }
    }

    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        public List<BingOfficesException> Exceptions { get; } = new();
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }

    private sealed class ExcelRow
    {
        public string Code { get; set; }
        public int Count { get; set; }
    }

    private sealed class ExcelWorkbook
    {
        public List<ExcelRow> Rows { get; } = new();
    }

    private sealed class FailureRow
    {
        [ExcelRegex("^OK-")]
        public string Code { get; set; }
    }

    private sealed class FailureWorkbook
    {
        public List<FailureRow> Rows { get; } = new();
    }

    private sealed class AsyncOnlyReadStream : Stream
    {
        private readonly MemoryStream _inner;

        private readonly Action _afterAsyncRead;

        public AsyncOnlyReadStream(byte[] bytes, Action afterAsyncRead = null)
        {
            _inner = new MemoryStream(bytes, writable: false);
            _afterAsyncRead = afterAsyncRead;
        }

        public int SyncReadCount { get; private set; }
        public int AsyncReadCount { get; private set; }
        public CancellationToken LastReadCancellationToken { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            throw new InvalidOperationException("同步 Read 不允许用于 Async API。");
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            AsyncReadCount++;
            LastReadCancellationToken = cancellationToken;
            return ReadAsyncCore(buffer, offset, count, cancellationToken);
        }

        private async Task<int> ReadAsyncCore(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken);
            _afterAsyncRead?.Invoke();
            return read;
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            AsyncReadCount++;
            LastReadCancellationToken = cancellationToken;
            return ReadMemoryAsyncCore(buffer, cancellationToken);
        }

        private async ValueTask<int> ReadMemoryAsyncCore(Memory<byte> buffer, CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, cancellationToken);
            _afterAsyncRead?.Invoke();
            return read;
        }

        public override void Flush() { }
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class AsyncOnlyWriteStream : Stream
    {
        private readonly MemoryStream _inner = new();
        private readonly Action _afterAsyncWrite;
        private readonly bool _throwOnSyncFlush;

        public AsyncOnlyWriteStream(Action afterAsyncWrite = null, bool throwOnSyncFlush = false)
        {
            _afterAsyncWrite = afterAsyncWrite;
            _throwOnSyncFlush = throwOnSyncFlush;
        }

        public int SyncWriteCount { get; private set; }
        public int SyncFlushCount { get; private set; }
        public int AsyncWriteCount { get; private set; }
        public CancellationToken LastWriteCancellationToken { get; private set; }
        public override bool CanRead => false;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override int Read(Span<byte> buffer) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count)
        {
            SyncWriteCount++;
            throw new InvalidOperationException("同步 Write 不允许用于 Async API。");
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            AsyncWriteCount++;
            LastWriteCancellationToken = cancellationToken;
            return WriteAsyncCore(buffer, offset, count, cancellationToken);
        }

        private async Task WriteAsyncCore(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            await _inner.WriteAsync(buffer, offset, count, cancellationToken);
            _afterAsyncWrite?.Invoke();
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWriteCount++;
            LastWriteCancellationToken = cancellationToken;
            return WriteMemoryAsyncCore(buffer, cancellationToken);
        }

        private async ValueTask WriteMemoryAsyncCore(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken)
        {
            await _inner.WriteAsync(buffer, cancellationToken);
            _afterAsyncWrite?.Invoke();
        }

        public byte[] ToArray() => _inner.ToArray();
        public override void Flush()
        {
            SyncFlushCount++;
            if (_throwOnSyncFlush)
                throw new InvalidOperationException("同步 Flush 不允许用于 Async API。");
        }
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => _inner.SetLength(value);
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class BlockingStagingFactory : INpoiAsyncStagingFactory
    {
        private readonly TaskCompletionSource<bool> _started =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task Started => _started.Task;

        public INpoiAsyncStaging Create(string prefix) => new BlockingStaging(_started);
    }

    private sealed class BlockingStaging : INpoiAsyncStaging
    {
        private readonly MemoryStream _buffer = new();
        private readonly TaskCompletionSource<bool> _started;

        public BlockingStaging(TaskCompletionSource<bool> started) => _started = started;

        public Stream WriteStream => _buffer;

        public Task FlushAsync(CancellationToken cancellationToken) => _buffer.FlushAsync(cancellationToken);

        public Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
        {
            _started.TrySetResult(true);
            return WaitForCancellationAsync(cancellationToken);
        }

        public void Dispose() => _buffer.Dispose();

        private static async Task WaitForCancellationAsync(CancellationToken cancellationToken)
            => await Task.Delay(Timeout.Infinite, cancellationToken);
    }

    private sealed class FailingStagingFactory : INpoiAsyncStagingFactory
    {
        public INpoiAsyncStaging Create(string prefix) => new FailingStaging();
    }

    private sealed class FailingStaging : INpoiAsyncStaging
    {
        private readonly MemoryStream _buffer = new();

        public Stream WriteStream => _buffer;

        public Task FlushAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.CompletedTask;
        }

        public Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
            => Task.FromException(new InvalidOperationException("staging 复制失败"));

        public void Dispose() => throw new IOException("staging 清理失败");
    }

    private sealed class BlockingAsyncReadStream : Stream
    {
        private readonly MemoryStream _inner;
        private readonly TaskCompletionSource<bool> _started = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public BlockingAsyncReadStream(byte[] content = null)
            => _inner = new MemoryStream(content ?? Array.Empty<byte>(), writable: false);

        public Task Started => _started.Task;
        public CancellationToken LastReadCancellationToken { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
            => throw new InvalidOperationException("Async-only blocking stream received sync Read.");
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            LastReadCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return WaitForCancellationAsync(cancellationToken);
        }
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            LastReadCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return new ValueTask<int>(WaitForCancellationAsync(cancellationToken));
        }
        private static async Task<int> WaitForCancellationAsync(CancellationToken cancellationToken)
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return 0;
        }
        public override void Flush() { }
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class BlockingAsyncWriteStream : Stream
    {
        private readonly TaskCompletionSource<bool> _started = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task Started => _started.Task;
        public CancellationToken LastWriteCancellationToken { get; private set; }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count)
            => throw new InvalidOperationException("Async-only blocking stream received sync Write.");
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            LastWriteCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return WaitForCancellationAsync(cancellationToken);
        }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            LastWriteCancellationToken = cancellationToken;
            _started.TrySetResult(true);
            return new ValueTask(WaitForCancellationAsync(cancellationToken));
        }
        private static async Task WaitForCancellationAsync(CancellationToken cancellationToken)
            => await Task.Delay(Timeout.Infinite, cancellationToken);
        public override void Flush() { }
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}

[CollectionDefinition("Excel Async staging", DisableParallelization = true)]
public sealed class ExcelAsyncStagingCollection
{
}
