using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Excel 流扩展的参数转发、异常传播和内部流生命周期测试。
/// </summary>
public class ExcelStreamExtensionsTest
{
    [Fact]
    public async Task NullAndWhitespaceArguments_ShouldBeRejected()
    {
        var exporter = new TrackingExcelExporter();
        var importer = new TrackingExcelImporter();
        var exportRequest = CreateExportRequest();
        var importRequest = CreateImportRequest();
        var content = new byte[] { 1, 2, 3 };

        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToBytes(null, exportRequest));
        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToFile(null, exportRequest, "target.xlsx"));
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToBytesAsync(null, exportRequest));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = ExcelStreamExtensions.ExportToFileAsync(null, exportRequest, "target.xlsx");
        });

        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromBytes<ExcelWorkbook>(null, content, importRequest));
        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromFile<ExcelWorkbook>(null, "source.xlsx", importRequest));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = ExcelStreamExtensions.ImportFromBytesAsync<ExcelWorkbook>(null, content, importRequest);
        });
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromFileAsync<ExcelWorkbook>(null, "source.xlsx", importRequest));

        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToBytes(exporter, null));
        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToFile(exporter, null, "target.xlsx"));
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToBytesAsync(exporter, null));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = ExcelStreamExtensions.ExportToFileAsync(exporter, null, "target.xlsx");
        });

        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromBytes<ExcelWorkbook>(importer, null, importRequest));
        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromBytes<ExcelWorkbook>(importer, content, null));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = ExcelStreamExtensions.ImportFromBytesAsync<ExcelWorkbook>(importer, null, importRequest);
        });
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = ExcelStreamExtensions.ImportFromBytesAsync<ExcelWorkbook>(importer, content, null);
        });
        Assert.Throws<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromFile<ExcelWorkbook>(importer, "source.xlsx", null));
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            ExcelStreamExtensions.ImportFromFileAsync<ExcelWorkbook>(importer, "source.xlsx", null));

        Assert.Throws<ArgumentException>(() =>
            ExcelStreamExtensions.ExportToFile(exporter, exportRequest, " \t"));
        Assert.Throws<ArgumentException>(() =>
        {
            _ = ExcelStreamExtensions.ExportToFileAsync(exporter, exportRequest, "\r\n");
        });
        Assert.Throws<ArgumentException>(() =>
            ExcelStreamExtensions.ImportFromFile<ExcelWorkbook>(importer, " ", importRequest));
        await Assert.ThrowsAsync<ArgumentException>(() =>
            ExcelStreamExtensions.ImportFromFileAsync<ExcelWorkbook>(importer, "\t", importRequest));
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ExportToBytes_ShouldForwardRequestAndDisposeDestination(TrackingBehavior behavior)
    {
        var exporter = new TrackingExcelExporter { Behavior = behavior };
        var request = CreateExportRequest();
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = ExcelStreamExtensions.ExportToBytes(exporter, request, cancellation.Token);
            Assert.Equal(exporter.Payload, actual);
        }
        else
        {
            var exception = Record.Exception(() =>
                ExcelStreamExtensions.ExportToBytes(exporter, request, cancellation.Token));
            AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
        }

        Assert.Same(request, exporter.LastRequest);
        Assert.Equal(cancellation.Token, exporter.LastCancellationToken);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Fact]
    public void ExportToBytes_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 字节导出失败");
        var exporter = new TrackingExcelExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = Record.Exception(() =>
            ExcelStreamExtensions.ExportToBytes(exporter, CreateExportRequest()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public async Task ExportToBytesAsync_ShouldForwardRequestAndDisposeDestination(TrackingBehavior behavior)
    {
        var exporter = new TrackingExcelExporter { Behavior = behavior };
        var request = CreateExportRequest();
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = await ExcelStreamExtensions.ExportToBytesAsync(exporter, request, cancellation.Token);
            Assert.Equal(exporter.Payload, actual);
        }
        else
        {
            var exception = await Record.ExceptionAsync(() =>
                ExcelStreamExtensions.ExportToBytesAsync(exporter, request, cancellation.Token));
            AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
        }

        Assert.Same(request, exporter.LastRequest);
        Assert.Equal(cancellation.Token, exporter.LastCancellationToken);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Fact]
    public async Task ExportToBytesAsync_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 异步字节导出失败");
        var exporter = new TrackingExcelExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = await Record.ExceptionAsync(() =>
            ExcelStreamExtensions.ExportToBytesAsync(exporter, CreateExportRequest()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ExportToFile_ShouldForwardRequestAndReleaseTarget(TrackingBehavior behavior)
    {
        var exporter = new TrackingExcelExporter { Behavior = behavior };
        var request = CreateExportRequest();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            if (behavior == TrackingBehavior.Success)
            {
                ExcelStreamExtensions.ExportToFile(exporter, request, path, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = Record.Exception(() =>
                    ExcelStreamExtensions.ExportToFile(exporter, request, path, cancellation.Token));
                AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
            }

            Assert.Same(request, exporter.LastRequest);
            Assert.Equal(path, exporter.LastPath);
            Assert.Equal(cancellation.Token, exporter.LastCancellationToken);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Fact]
    public void ExportToFile_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 文件导出失败");
        var exporter = new TrackingExcelExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            var actual = Record.Exception(() =>
                ExcelStreamExtensions.ExportToFile(exporter, CreateExportRequest(), path));
            Assert.Same(expected, actual);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public async Task ExportToFileAsync_ShouldForwardRequestAndReleaseTarget(TrackingBehavior behavior)
    {
        var exporter = new TrackingExcelExporter { Behavior = behavior };
        var request = CreateExportRequest();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            if (behavior == TrackingBehavior.Success)
            {
                await ExcelStreamExtensions.ExportToFileAsync(exporter, request, path, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = await Record.ExceptionAsync(() =>
                    ExcelStreamExtensions.ExportToFileAsync(exporter, request, path, cancellation.Token));
                AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
            }

            Assert.Same(request, exporter.LastRequest);
            Assert.Equal(path, exporter.LastPath);
            Assert.Equal(cancellation.Token, exporter.LastCancellationToken);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Fact]
    public async Task ExportToFileAsync_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 异步文件导出失败");
        var exporter = new TrackingExcelExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            var actual = await Record.ExceptionAsync(() => ExcelStreamExtensions.ExportToFileAsync(
                exporter, CreateExportRequest(), path));
            Assert.Same(expected, actual);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ImportFromBytes_ShouldForwardRequestAndDisposeSource(TrackingBehavior behavior)
    {
        var importer = new TrackingExcelImporter { Behavior = behavior };
        var content = new byte[] { 7, 8, 9 };
        var request = CreateImportRequest();
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = ExcelStreamExtensions.ImportFromBytes(importer, content, request, cancellation.Token);
            Assert.Same(importer.LastResult, actual);
        }
        else
        {
            var exception = Record.Exception(() =>
                ExcelStreamExtensions.ImportFromBytes(importer, content, request, cancellation.Token));
            AssertExpectedBehaviorException(importer, exception, cancellation.Token);
        }

        Assert.Same(request, importer.LastRequest);
        Assert.Equal(cancellation.Token, importer.LastCancellationToken);
        Assert.Equal(content[0], importer.FirstByte);
        Assert.False(importer.SourceWasWritable);
        AssertStreamDisposed(importer.LastSource);
    }

    [Fact]
    public void ImportFromBytes_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 字节导入失败");
        var importer = new TrackingExcelImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = Record.Exception(() => ExcelStreamExtensions.ImportFromBytes(
            importer, new byte[] { 1 }, CreateImportRequest()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(importer.LastSource);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public async Task ImportFromBytesAsync_ShouldForwardRequestAndDisposeSource(TrackingBehavior behavior)
    {
        var importer = new TrackingExcelImporter { Behavior = behavior };
        var content = new byte[] { 10, 11, 12 };
        var request = CreateImportRequest();
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = await ExcelStreamExtensions.ImportFromBytesAsync(importer, content, request,
                cancellation.Token);
            Assert.Same(importer.LastResult, actual);
        }
        else
        {
            var exception = await Record.ExceptionAsync(() =>
                ExcelStreamExtensions.ImportFromBytesAsync(importer, content, request, cancellation.Token));
            AssertExpectedBehaviorException(importer, exception, cancellation.Token);
        }

        Assert.Same(request, importer.LastRequest);
        Assert.Equal(cancellation.Token, importer.LastCancellationToken);
        Assert.Equal(content[0], importer.FirstByte);
        Assert.False(importer.SourceWasWritable);
        AssertStreamDisposed(importer.LastSource);
    }

    [Fact]
    public async Task ImportFromBytesAsync_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 异步字节导入失败");
        var importer = new TrackingExcelImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = await Record.ExceptionAsync(() => ExcelStreamExtensions.ImportFromBytesAsync(
            importer, new byte[] { 1 }, CreateImportRequest()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(importer.LastSource);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ImportFromFile_ShouldForwardRequestAndReleaseSource(TrackingBehavior behavior)
    {
        var importer = new TrackingExcelImporter { Behavior = behavior };
        var request = CreateImportRequest();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            File.WriteAllBytes(path, new byte[] { 13, 14, 15 });
            if (behavior == TrackingBehavior.Success)
            {
                var actual = ExcelStreamExtensions.ImportFromFile(importer, path, request, cancellation.Token);
                Assert.Same(importer.LastResult, actual);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = Record.Exception(() =>
                    ExcelStreamExtensions.ImportFromFile(importer, path, request, cancellation.Token));
                AssertExpectedBehaviorException(importer, exception, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }

            Assert.Same(request, importer.LastRequest);
            Assert.Equal(path, importer.LastPath);
            Assert.Equal(cancellation.Token, importer.LastCancellationToken);
            Assert.Equal(13, importer.FirstByte);
            AssertStreamDisposed(importer.LastSource);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Fact]
    public void ImportFromFile_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 文件导入失败");
        var importer = new TrackingExcelImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            File.WriteAllBytes(path, new byte[] { 1 });
            var actual = Record.Exception(() => ExcelStreamExtensions.ImportFromFile(
                importer, path, CreateImportRequest()));
            Assert.Same(expected, actual);
            AssertStreamDisposed(importer.LastSource);
            AssertFileCanBeOpenedExclusivelyAndDeleted(path);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public async Task ImportFromFileAsync_ShouldForwardRequestAndReleaseSource(TrackingBehavior behavior)
    {
        var importer = new TrackingExcelImporter { Behavior = behavior };
        var request = CreateImportRequest();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            File.WriteAllBytes(path, new byte[] { 16, 17, 18 });
            if (behavior == TrackingBehavior.Success)
            {
                var actual = await ExcelStreamExtensions.ImportFromFileAsync(importer, path, request,
                    cancellation.Token);
                Assert.Same(importer.LastResult, actual);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = await Record.ExceptionAsync(() =>
                    ExcelStreamExtensions.ImportFromFileAsync(importer, path, request, cancellation.Token));
                AssertExpectedBehaviorException(importer, exception, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }

            Assert.Same(request, importer.LastRequest);
            Assert.Equal(path, importer.LastPath);
            Assert.Equal(cancellation.Token, importer.LastCancellationToken);
            Assert.Equal(16, importer.FirstByte);
            AssertStreamDisposed(importer.LastSource);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    [Fact]
    public async Task ImportFromFileAsync_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("Excel 异步文件导入失败");
        var importer = new TrackingExcelImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".xlsx");

        try
        {
            File.WriteAllBytes(path, new byte[] { 1 });
            var actual = await Record.ExceptionAsync(() => ExcelStreamExtensions.ImportFromFileAsync(
                importer, path, CreateImportRequest()));
            Assert.Same(expected, actual);
            AssertStreamDisposed(importer.LastSource);
            AssertFileCanBeOpenedExclusivelyAndDeleted(path);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    private static ExcelWorkbookExportRequest CreateExportRequest() =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new ExcelRow { Name = "测试数据" }
        }));

    private static ExcelWorkbookImportRequest<ExcelWorkbook> CreateImportRequest() =>
        ExcelImport.Workbook<ExcelWorkbook>(builder => builder.Sheet("Data", workbook => workbook.Rows));

    private static CancellationTokenSource CreateCancellation(TrackingBehavior behavior)
    {
        var cancellation = new CancellationTokenSource();
        if (behavior == TrackingBehavior.Cancellation)
            cancellation.Cancel();
        return cancellation;
    }

    private static string CreateTemporaryPath(string extension)
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, "stream" + extension);
    }

    private static void DeleteTemporaryPath(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory)
            && Directory.GetFileSystemEntries(directory).Length == 0)
            Directory.Delete(directory);
    }

    private static void AssertFileCanBeOpenedExclusivelyAndDeleted(string path)
    {
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            Assert.True(stream.CanRead);
        }
        File.Delete(path);
        Assert.False(File.Exists(path));
    }

    private static void AssertStreamDisposed(Stream stream)
    {
        Assert.NotNull(stream);
        Assert.False(stream.CanRead);
        Assert.False(stream.CanWrite);
    }

    private static void AssertExpectedBehaviorException(TrackingExcelExporter exporter, Exception exception,
        CancellationToken token)
    {
        if (exporter.Behavior == TrackingBehavior.Failure)
            Assert.Same(exporter.ExceptionToThrow, exception);
        else
        {
            var cancellation = Assert.IsType<OperationCanceledException>(exception);
            Assert.Equal(token, cancellation.CancellationToken);
        }
    }

    private static void AssertExpectedBehaviorException(TrackingExcelImporter importer, Exception exception,
        CancellationToken token)
    {
        if (importer.Behavior == TrackingBehavior.Failure)
            Assert.Same(importer.ExceptionToThrow, exception);
        else
        {
            var cancellation = Assert.IsType<OperationCanceledException>(exception);
            Assert.Equal(token, cancellation.CancellationToken);
        }
    }

    public enum TrackingBehavior
    {
        Success,
        Failure,
        Cancellation
    }

    private sealed class ExcelWorkbook
    {
        public List<ExcelRow> Rows { get; } = new List<ExcelRow>();
    }

    private sealed class ExcelRow
    {
        public string Name { get; set; }
    }

    private sealed class TrackingExcelExporter : IExcelExporter
    {
        public TrackingBehavior Behavior { get; set; }

        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("Excel fake failure");

        public byte[] Payload { get; } = Encoding.UTF8.GetBytes("excel-payload");

        public ExcelWorkbookExportRequest LastRequest { get; private set; }

        public string LastPath { get; private set; }

        public Stream LastDestination { get; private set; }

        public CancellationToken LastCancellationToken { get; private set; }

        public void Export(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            LastDestination = destination;
            LastCancellationToken = cancellationToken;
            ApplyBehavior(cancellationToken);
            destination.Write(Payload, 0, Payload.Length);
        }

        public async Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            LastDestination = destination;
            LastCancellationToken = cancellationToken;
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            await destination.WriteAsync(Payload, 0, Payload.Length, cancellationToken);
        }

        public void ExportToFile(ExcelWorkbookExportRequest request, string path,
            CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            LastPath = path;
            LastCancellationToken = cancellationToken;
            ApplyBehavior(cancellationToken);
            File.WriteAllBytes(path, Payload);
        }

        public async Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path,
            CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            LastPath = path;
            LastCancellationToken = cancellationToken;
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            await File.WriteAllBytesAsync(path, Payload, cancellationToken);
        }

        private void ApplyBehavior(CancellationToken cancellationToken)
        {
            if (Behavior == TrackingBehavior.Failure)
                throw ExceptionToThrow;
            if (Behavior == TrackingBehavior.Cancellation)
            {
                if (!cancellationToken.IsCancellationRequested)
                    throw new InvalidOperationException("fake 未收到预取消令牌");
                throw new OperationCanceledException(cancellationToken);
            }
        }
    }

    private sealed class TrackingExcelImporter : IExcelImporter
    {
        public TrackingBehavior Behavior { get; set; }

        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("Excel fake failure");

        public ExcelWorkbookImportResult<ExcelWorkbook> LastResult { get; private set; }

        public object LastRequest { get; private set; }

        public string LastPath { get; private set; }

        public Stream LastSource { get; private set; }

        public bool SourceWasWritable { get; private set; }

        public int FirstByte { get; private set; }

        public CancellationToken LastCancellationToken { get; private set; }

        public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
            where TWorkbook : class, new()
        {
            Record(source, request, cancellationToken);
            ApplyBehavior(cancellationToken);
            return CreateResult<TWorkbook>();
        }

        public async Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
            where TWorkbook : class, new()
        {
            Record(source, request, cancellationToken);
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            return CreateResult<TWorkbook>();
        }

        private ExcelWorkbookImportResult<TWorkbook> CreateResult<TWorkbook>() where TWorkbook : class, new()
        {
            var result = new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(),
                Array.Empty<ExcelSheetImportResult>(), Array.Empty<ExcelImportError>(), false, null);
            if (typeof(TWorkbook) == typeof(ExcelWorkbook))
                LastResult = (ExcelWorkbookImportResult<ExcelWorkbook>)(object)result;
            return result;
        }

        private void Record<TWorkbook>(Stream source, ExcelWorkbookImportRequest<TWorkbook> request,
            CancellationToken cancellationToken) where TWorkbook : class, new()
        {
            LastSource = source;
            LastRequest = request;
            LastPath = source is FileStream file ? file.Name : null;
            LastCancellationToken = cancellationToken;
            SourceWasWritable = source.CanWrite;
            var position = source.Position;
            FirstByte = source.ReadByte();
            source.Position = position;
        }

        private void ApplyBehavior(CancellationToken cancellationToken)
        {
            if (Behavior == TrackingBehavior.Failure)
                throw ExceptionToThrow;
            if (Behavior == TrackingBehavior.Cancellation)
            {
                if (!cancellationToken.IsCancellationRequested)
                    throw new InvalidOperationException("fake 未收到预取消令牌");
                throw new OperationCanceledException(cancellationToken);
            }
        }
    }
}
