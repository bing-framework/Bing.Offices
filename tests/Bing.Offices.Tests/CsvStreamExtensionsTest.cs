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
    /// <summary>
    /// 验证空值和空白参数应被拒绝。
    /// </summary>
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
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            CsvStreamExtensions.ExportToBytesAsync<CsvRow>(null, data, options));

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
            CsvStreamExtensions.ImportFromFile<CsvRow>(importer, " ", importOptions));
        await Assert.ThrowsAsync<ArgumentException>(() =>
            CsvStreamExtensions.ImportFromFileAsync<CsvRow>(importer, "\t", importOptions));
    }

    /// <summary>
    /// 验证导出到字节应转发参数并释放目标。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导出到字节应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导出到字节异步应转发参数并释放目标。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导出到字节异步应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从字节应转发参数并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从字节应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从字节异步应转发参数并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从字节异步应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从文件应转发参数并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从文件应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从文件异步应转发参数并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从文件异步应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 创建测试数据。
    /// </summary>
    /// <returns>包含一行固定名称数据的 CSV 测试数组。</returns>
    private static CsvRow[] CreateData() => new[] { new CsvRow { Name = "测试数据" } };

    /// <summary>
    /// 创建测试用取消源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
    /// <returns>测试取消源；取消场景中已取消，其他场景中未取消。</returns>
    private static CancellationTokenSource CreateCancellation(TrackingBehavior behavior)
    {
        var cancellation = new CancellationTokenSource();
        if (behavior == TrackingBehavior.Cancellation)
            cancellation.Cancel();
        return cancellation;
    }

    /// <summary>
    /// 创建临时文件路径。
    /// </summary>
    /// <param name="extension">文件扩展名。</param>
    /// <returns>已创建的独立临时目录内、带指定扩展名的文件路径。</returns>
    private static string CreateTemporaryPath(string extension)
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, "stream" + extension);
    }

    /// <summary>
    /// 删除临时文件。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    private static void DeleteTemporaryPath(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory)
            && Directory.GetFileSystemEntries(directory).Length == 0)
            Directory.Delete(directory);
    }

    /// <summary>
    /// 验证文件可独占打开并删除。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    private static void AssertFileCanBeOpenedExclusivelyAndDeleted(string path)
    {
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            Assert.True(stream.CanRead);
        }
        File.Delete(path);
        Assert.False(File.Exists(path));
    }

    /// <summary>
    /// 验证流已被释放。
    /// </summary>
    /// <param name="stream">参与操作的流。</param>
    private static void AssertStreamDisposed(Stream stream)
    {
        Assert.NotNull(stream);
        Assert.False(stream.CanRead);
        Assert.False(stream.CanWrite);
    }

    /// <summary>
    /// 断言测试替身按预期抛出异常或取消。
    /// </summary>
    /// <param name="exporter">Excel 或 CSV 导出器。</param>
    /// <param name="exception">测试期间要传播的异常。</param>
    /// <param name="token">操作令牌。</param>
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

    /// <summary>
    /// 断言测试替身按预期抛出异常或取消。
    /// </summary>
    /// <param name="importer">Excel 或 CSV 导入器。</param>
    /// <param name="exception">测试期间要传播的异常。</param>
    /// <param name="token">操作令牌。</param>
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

    /// <summary>
    /// 表示测试替身记录的行为状态。
    /// </summary>
    public enum TrackingBehavior
    {
        /// <summary>
        /// 表示操作成功完成。
        /// </summary>
        Success,
        /// <summary>
        /// 表示操作以失败结束。
        /// </summary>
        Failure,
        /// <summary>
        /// 表示操作因取消而结束。
        /// </summary>
        Cancellation
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class CsvRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 记录测试过程中的状态或调用次数。
    /// </summary>
    private sealed class TrackingCsvExporter : ICsvExporter
    {
        /// <summary>
        /// 获取或设置行为。
        /// </summary>
        public TrackingBehavior Behavior { get; set; }

        /// <summary>
        /// 获取或设置要抛出的异常。
        /// </summary>
        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("CSV fake failure");

        /// <summary>
        /// 获取负载。
        /// </summary>
        public byte[] Payload { get; } = Encoding.UTF8.GetBytes("csv-payload");

        /// <summary>
        /// 获取或设置最近一次数据。
        /// </summary>
        public object LastData { get; private set; }

        /// <summary>
        /// 获取或设置最近一次选项。
        /// </summary>
        public object LastOptions { get; private set; }

        /// <summary>
        /// 获取或设置最近一次路径。
        /// </summary>
        public string LastPath { get; private set; }

        /// <summary>
        /// 获取或设置最近一次目标流。
        /// </summary>
        public Stream LastDestination { get; private set; }

        /// <summary>
        /// 获取或设置最近一次取消令牌。
        /// </summary>
        public CancellationToken LastCancellationToken { get; private set; }

        /// <inheritdoc />
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

        /// <inheritdoc />
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

        /// <inheritdoc />
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

        /// <inheritdoc />
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

        /// <summary>
        /// 应用测试替身配置的行为。
        /// </summary>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
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

    /// <summary>
    /// 记录测试过程中的状态或调用次数。
    /// </summary>
    private sealed class TrackingCsvImporter : ICsvImporter
    {
        /// <summary>
        /// 获取或设置行为。
        /// </summary>
        public TrackingBehavior Behavior { get; set; }

        /// <summary>
        /// 获取或设置要抛出的异常。
        /// </summary>
        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("CSV fake failure");

        /// <summary>
        /// 获取或设置最近一次选项。
        /// </summary>
        public object LastOptions { get; private set; }

        /// <summary>
        /// 获取或设置最近一次路径。
        /// </summary>
        public string LastPath { get; private set; }

        /// <summary>
        /// 获取或设置最近一次源流。
        /// </summary>
        public Stream LastSource { get; private set; }

        /// <summary>
        /// 获取或设置最近一次结果。
        /// </summary>
        public object LastResult { get; private set; }

        /// <summary>
        /// 获取或设置源是否可写。
        /// </summary>
        public bool SourceWasWritable { get; private set; }

        /// <summary>
        /// 获取或设置首字节。
        /// </summary>
        public int FirstByte { get; private set; }

        /// <summary>
        /// 获取或设置最近一次取消令牌。
        /// </summary>
        public CancellationToken LastCancellationToken { get; private set; }

        /// <inheritdoc />
        public CsvImportResult<T> Import<T>(Stream source, CsvImportOptions<T> options = null,
            CancellationToken cancellationToken = default) where T : class, new()
        {
            Record(source, options, cancellationToken);
            ApplyBehavior(cancellationToken);
            return CreateResult<T>();
        }

        /// <inheritdoc />
        public async Task<CsvImportResult<T>> ImportAsync<T>(Stream source, CsvImportOptions<T> options = null,
            CancellationToken cancellationToken = default) where T : class, new()
        {
            Record(source, options, cancellationToken);
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            return CreateResult<T>();
        }

        /// <summary>
        /// 创建测试导入结果。
        /// </summary>
        /// <typeparam name="T">泛型参数 T 表示方法处理的数据类型。</typeparam>
        /// <returns>数据和错误集合均为空的 CSV 导入结果。</returns>
        private CsvImportResult<T> CreateResult<T>() where T : class, new()
        {
            var result = new CsvImportResult<T>(Array.Empty<T>(), Array.Empty<CsvImportError>());
            LastResult = result;
            return result;
        }

        /// <summary>
        /// 记录测试替身调用。
        /// </summary>
        /// <typeparam name="T">泛型参数 T 表示方法处理的数据类型。</typeparam>
        /// <param name="source">输入流。</param>
        /// <param name="options">导入或导出选项。</param>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
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

        /// <summary>
        /// 应用测试替身配置的行为。
        /// </summary>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
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
