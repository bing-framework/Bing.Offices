using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Bing.Offices.Extensions;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// CSV 流扩展的参数转发、异常传播和内部流生命周期测试。
/// </summary>
public class CsvStreamExtensionsTest
{
    [Fact]
    public async Task NullAndWhitespaceArguments_ShouldBeRejected()
    {
        var data = CreateData();
        var exporter = new TrackingCsvExporter();
        var importer = new TrackingCsvImporter();
        var options = new CsvExportOptions<CsvRow>();
        var importOptions = new CsvImportOptions<CsvRow>();
        var content = new byte[] { 1, 2, 3 };

        Assert.Throws<ArgumentNullException>(() =>
            CsvStreamExtensions.ExportToBytes<CsvRow>(null, data, options));
        Assert.Throws<ArgumentNullException>(() =>
            CsvStreamExtensions.ExportToFile<CsvRow>(null, data, "target.csv", options));
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            CsvStreamExtensions.ExportToBytesAsync<CsvRow>(null, data, options));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = CsvStreamExtensions.ExportToFileAsync<CsvRow>(null, data, "target.csv", options);
        });

        Assert.Throws<ArgumentNullException>(() =>
            CsvStreamExtensions.ImportFromBytes<CsvRow>(null, content, importOptions));
        Assert.Throws<ArgumentNullException>(() =>
            CsvStreamExtensions.ImportFromFile<CsvRow>(null, "source.csv", importOptions));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = CsvStreamExtensions.ImportFromBytesAsync<CsvRow>(null, content, importOptions);
        });
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            CsvStreamExtensions.ImportFromFileAsync<CsvRow>(null, "source.csv", importOptions));

        Assert.Throws<ArgumentNullException>(() =>
            CsvStreamExtensions.ImportFromBytes<CsvRow>(importer, null, importOptions));
        Assert.Throws<ArgumentNullException>(() =>
        {
            _ = CsvStreamExtensions.ImportFromBytesAsync<CsvRow>(importer, null, importOptions);
        });

        Assert.Throws<ArgumentException>(() =>
            CsvStreamExtensions.ExportToFile<CsvRow>(exporter, data, " \t", options));
        Assert.Throws<ArgumentException>(() =>
        {
            _ = CsvStreamExtensions.ExportToFileAsync<CsvRow>(exporter, data, "\r\n", options);
        });
        Assert.Throws<ArgumentException>(() =>
            CsvStreamExtensions.ImportFromFile<CsvRow>(importer, " ", importOptions));
        await Assert.ThrowsAsync<ArgumentException>(() =>
            CsvStreamExtensions.ImportFromFileAsync<CsvRow>(importer, "\t", importOptions));
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ExportToBytes_ShouldForwardArgumentsAndDisposeDestination(TrackingBehavior behavior)
    {
        var exporter = new TrackingCsvExporter { Behavior = behavior };
        var data = CreateData();
        var options = new CsvExportOptions<CsvRow> { IncludeHeader = false, Delimiter = ';' };
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = CsvStreamExtensions.ExportToBytes(exporter, data, options, cancellation.Token);
            Assert.Equal(exporter.Payload, actual);
        }
        else
        {
            var exception = Record.Exception(() =>
                CsvStreamExtensions.ExportToBytes(exporter, data, options, cancellation.Token));
            AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
        }

        Assert.Same(data, exporter.LastData);
        Assert.Same(options, exporter.LastOptions);
        Assert.Equal(cancellation.Token, exporter.LastCancellationToken);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Fact]
    public void ExportToBytes_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("CSV 字节导出失败");
        var exporter = new TrackingCsvExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = Record.Exception(() =>
            CsvStreamExtensions.ExportToBytes(exporter, CreateData(), new CsvExportOptions<CsvRow>()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public async Task ExportToBytesAsync_ShouldForwardArgumentsAndDisposeDestination(TrackingBehavior behavior)
    {
        var exporter = new TrackingCsvExporter { Behavior = behavior };
        var data = CreateData();
        var options = new CsvExportOptions<CsvRow> { IncludeHeader = false, Delimiter = ';' };
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = await CsvStreamExtensions.ExportToBytesAsync(exporter, data, options,
                cancellation.Token);
            Assert.Equal(exporter.Payload, actual);
        }
        else
        {
            var exception = await Record.ExceptionAsync(() =>
                CsvStreamExtensions.ExportToBytesAsync(exporter, data, options, cancellation.Token));
            AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
        }

        Assert.Same(data, exporter.LastData);
        Assert.Same(options, exporter.LastOptions);
        Assert.Equal(cancellation.Token, exporter.LastCancellationToken);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Fact]
    public async Task ExportToBytesAsync_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("CSV 异步字节导出失败");
        var exporter = new TrackingCsvExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = await Record.ExceptionAsync(() =>
            CsvStreamExtensions.ExportToBytesAsync(exporter, CreateData(), new CsvExportOptions<CsvRow>()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(exporter.LastDestination);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ExportToFile_ShouldForwardArgumentsAndReleaseTarget(TrackingBehavior behavior)
    {
        var exporter = new TrackingCsvExporter { Behavior = behavior };
        var data = CreateData();
        var options = new CsvExportOptions<CsvRow>();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".csv");

        try
        {
            if (behavior == TrackingBehavior.Success)
            {
                CsvStreamExtensions.ExportToFile(exporter, data, path, options, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = Record.Exception(() =>
                    CsvStreamExtensions.ExportToFile(exporter, data, path, options, cancellation.Token));
                AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
            }

            Assert.Same(data, exporter.LastData);
            Assert.Same(options, exporter.LastOptions);
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
        var expected = new InvalidOperationException("CSV 文件导出失败");
        var exporter = new TrackingCsvExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".csv");

        try
        {
            var actual = Record.Exception(() => CsvStreamExtensions.ExportToFile(exporter, CreateData(), path));
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
    public async Task ExportToFileAsync_ShouldForwardArgumentsAndReleaseTarget(TrackingBehavior behavior)
    {
        var exporter = new TrackingCsvExporter { Behavior = behavior };
        var data = CreateData();
        var options = new CsvExportOptions<CsvRow>();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".csv");

        try
        {
            if (behavior == TrackingBehavior.Success)
            {
                await CsvStreamExtensions.ExportToFileAsync(exporter, data, path, options,
                    cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = await Record.ExceptionAsync(() =>
                    CsvStreamExtensions.ExportToFileAsync(exporter, data, path, options,
                        cancellation.Token));
                AssertExpectedBehaviorException(exporter, exception, cancellation.Token);
            }

            Assert.Same(data, exporter.LastData);
            Assert.Same(options, exporter.LastOptions);
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
        var expected = new InvalidOperationException("CSV 异步文件导出失败");
        var exporter = new TrackingCsvExporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".csv");

        try
        {
            var actual = await Record.ExceptionAsync(() =>
                CsvStreamExtensions.ExportToFileAsync(exporter, CreateData(), path));
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
    public void ImportFromBytes_ShouldForwardArgumentsAndDisposeSource(TrackingBehavior behavior)
    {
        var importer = new TrackingCsvImporter { Behavior = behavior };
        var content = new byte[] { 7, 8, 9 };
        var options = new CsvImportOptions<CsvRow> { HasHeader = false, Delimiter = ';' };
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = CsvStreamExtensions.ImportFromBytes(importer, content, options, cancellation.Token);
            Assert.Same(importer.LastResult, actual);
        }
        else
        {
            var exception = Record.Exception(() =>
                CsvStreamExtensions.ImportFromBytes(importer, content, options, cancellation.Token));
            AssertExpectedBehaviorException(importer, exception, cancellation.Token);
        }

        Assert.Same(options, importer.LastOptions);
        Assert.Equal(cancellation.Token, importer.LastCancellationToken);
        Assert.Equal(content[0], importer.FirstByte);
        Assert.False(importer.SourceWasWritable);
        AssertStreamDisposed(importer.LastSource);
    }

    [Fact]
    public void ImportFromBytes_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("CSV 字节导入失败");
        var importer = new TrackingCsvImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = Record.Exception(() => CsvStreamExtensions.ImportFromBytes(
            importer, new byte[] { 1 }, new CsvImportOptions<CsvRow>()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(importer.LastSource);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public async Task ImportFromBytesAsync_ShouldForwardArgumentsAndDisposeSource(TrackingBehavior behavior)
    {
        var importer = new TrackingCsvImporter { Behavior = behavior };
        var content = new byte[] { 10, 11, 12 };
        var options = new CsvImportOptions<CsvRow> { HasHeader = false, Delimiter = ';' };
        using var cancellation = CreateCancellation(behavior);

        if (behavior == TrackingBehavior.Success)
        {
            var actual = await CsvStreamExtensions.ImportFromBytesAsync(importer, content, options,
                cancellation.Token);
            Assert.Same(importer.LastResult, actual);
        }
        else
        {
            var exception = await Record.ExceptionAsync(() =>
                CsvStreamExtensions.ImportFromBytesAsync(importer, content, options, cancellation.Token));
            AssertExpectedBehaviorException(importer, exception, cancellation.Token);
        }

        Assert.Same(options, importer.LastOptions);
        Assert.Equal(cancellation.Token, importer.LastCancellationToken);
        Assert.Equal(content[0], importer.FirstByte);
        Assert.False(importer.SourceWasWritable);
        AssertStreamDisposed(importer.LastSource);
    }

    [Fact]
    public async Task ImportFromBytesAsync_ShouldPreserveOriginalFailure()
    {
        var expected = new InvalidOperationException("CSV 异步字节导入失败");
        var importer = new TrackingCsvImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };

        var actual = await Record.ExceptionAsync(() => CsvStreamExtensions.ImportFromBytesAsync(
            importer, new byte[] { 1 }, new CsvImportOptions<CsvRow>()));

        Assert.Same(expected, actual);
        AssertStreamDisposed(importer.LastSource);
    }

    [Theory]
    [InlineData(TrackingBehavior.Success)]
    [InlineData(TrackingBehavior.Failure)]
    [InlineData(TrackingBehavior.Cancellation)]
    public void ImportFromFile_ShouldForwardArgumentsAndReleaseSource(TrackingBehavior behavior)
    {
        var importer = new TrackingCsvImporter { Behavior = behavior };
        var options = new CsvImportOptions<CsvRow>();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".csv");

        try
        {
            File.WriteAllBytes(path, new byte[] { 13, 14, 15 });
            if (behavior == TrackingBehavior.Success)
            {
                var actual = CsvStreamExtensions.ImportFromFile(importer, path, options, cancellation.Token);
                Assert.Same(importer.LastResult, actual);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = Record.Exception(() =>
                    CsvStreamExtensions.ImportFromFile(importer, path, options, cancellation.Token));
                AssertExpectedBehaviorException(importer, exception, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }

            Assert.Equal(path, importer.LastPath);
            Assert.Same(options, importer.LastOptions);
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
        var expected = new InvalidOperationException("CSV 文件导入失败");
        var importer = new TrackingCsvImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".csv");

        try
        {
            File.WriteAllBytes(path, new byte[] { 1 });
            var actual = Record.Exception(() => CsvStreamExtensions.ImportFromFile(
                importer, path, new CsvImportOptions<CsvRow>()));
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
    public async Task ImportFromFileAsync_ShouldForwardArgumentsAndReleaseSource(TrackingBehavior behavior)
    {
        var importer = new TrackingCsvImporter { Behavior = behavior };
        var options = new CsvImportOptions<CsvRow>();
        using var cancellation = CreateCancellation(behavior);
        var path = CreateTemporaryPath(".csv");

        try
        {
            File.WriteAllBytes(path, new byte[] { 16, 17, 18 });
            if (behavior == TrackingBehavior.Success)
            {
                var actual = await CsvStreamExtensions.ImportFromFileAsync(importer, path, options,
                    cancellation.Token);
                Assert.Same(importer.LastResult, actual);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }
            else
            {
                var exception = await Record.ExceptionAsync(() =>
                    CsvStreamExtensions.ImportFromFileAsync(importer, path, options,
                        cancellation.Token));
                AssertExpectedBehaviorException(importer, exception, cancellation.Token);
                AssertFileCanBeOpenedExclusivelyAndDeleted(path);
            }

            Assert.Equal(path, importer.LastPath);
            Assert.Same(options, importer.LastOptions);
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
        var expected = new InvalidOperationException("CSV 异步文件导入失败");
        var importer = new TrackingCsvImporter
        {
            Behavior = TrackingBehavior.Failure,
            ExceptionToThrow = expected
        };
        var path = CreateTemporaryPath(".csv");

        try
        {
            File.WriteAllBytes(path, new byte[] { 1 });
            var actual = await Record.ExceptionAsync(() => CsvStreamExtensions.ImportFromFileAsync(
                importer, path, new CsvImportOptions<CsvRow>()));
            Assert.Same(expected, actual);
            AssertStreamDisposed(importer.LastSource);
            AssertFileCanBeOpenedExclusivelyAndDeleted(path);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    private static CsvRow[] CreateData() => new[] { new CsvRow { Name = "测试数据" } };

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

    private static void AssertExpectedBehaviorException(TrackingCsvExporter exporter, Exception exception,
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

    private static void AssertExpectedBehaviorException(TrackingCsvImporter importer, Exception exception,
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

    private sealed class CsvRow
    {
        public string Name { get; set; }
    }

    private sealed class TrackingCsvExporter : ICsvExporter
    {
        public TrackingBehavior Behavior { get; set; }

        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("CSV fake failure");

        public byte[] Payload { get; } = Encoding.UTF8.GetBytes("csv-payload");

        public object LastData { get; private set; }

        public object LastOptions { get; private set; }

        public string LastPath { get; private set; }

        public Stream LastDestination { get; private set; }

        public CancellationToken LastCancellationToken { get; private set; }

        public void Export<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options = null,
            CancellationToken cancellationToken = default) where T : class, new()
        {
            LastData = data;
            LastOptions = options;
            LastDestination = destination;
            LastCancellationToken = cancellationToken;
            ApplyBehavior(cancellationToken);
            destination.Write(Payload, 0, Payload.Length);
        }

        public async Task ExportAsync<T>(IEnumerable<T> data, Stream destination,
            CsvExportOptions<T> options = null, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            LastData = data;
            LastOptions = options;
            LastDestination = destination;
            LastCancellationToken = cancellationToken;
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            await destination.WriteAsync(Payload, 0, Payload.Length, cancellationToken);
        }

        public void ExportToFile<T>(IEnumerable<T> data, string path, CsvExportOptions<T> options = null,
            CancellationToken cancellationToken = default) where T : class, new()
        {
            LastData = data;
            LastOptions = options;
            LastPath = path;
            LastCancellationToken = cancellationToken;
            ApplyBehavior(cancellationToken);
            File.WriteAllBytes(path, Payload);
        }

        public async Task ExportToFileAsync<T>(IEnumerable<T> data, string path,
            CsvExportOptions<T> options = null, CancellationToken cancellationToken = default)
            where T : class, new()
        {
            LastData = data;
            LastOptions = options;
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

    private sealed class TrackingCsvImporter : ICsvImporter
    {
        public TrackingBehavior Behavior { get; set; }

        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("CSV fake failure");

        public object LastOptions { get; private set; }

        public string LastPath { get; private set; }

        public Stream LastSource { get; private set; }

        public object LastResult { get; private set; }

        public bool SourceWasWritable { get; private set; }

        public int FirstByte { get; private set; }

        public CancellationToken LastCancellationToken { get; private set; }

        public CsvImportResult<T> Import<T>(Stream source, CsvImportOptions<T> options = null,
            CancellationToken cancellationToken = default) where T : class, new()
        {
            Record(source, options, cancellationToken);
            ApplyBehavior(cancellationToken);
            return CreateResult<T>();
        }

        public async Task<CsvImportResult<T>> ImportAsync<T>(Stream source, CsvImportOptions<T> options = null,
            CancellationToken cancellationToken = default) where T : class, new()
        {
            Record(source, options, cancellationToken);
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            return CreateResult<T>();
        }

        private CsvImportResult<T> CreateResult<T>() where T : class, new()
        {
            var result = new CsvImportResult<T>(Array.Empty<T>(), Array.Empty<CsvImportError>());
            LastResult = result;
            return result;
        }

        private void Record<T>(Stream source, CsvImportOptions<T> options, CancellationToken cancellationToken)
            where T : class, new()
        {
            LastSource = source;
            LastOptions = options;
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
