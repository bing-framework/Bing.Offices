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
    /// <summary>
    /// 验证空值和空白参数应被拒绝。
    /// </summary>
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
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToBytesAsync(null, exportRequest));

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
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            ExcelStreamExtensions.ExportToBytesAsync(exporter, null));

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
            ExcelStreamExtensions.ImportFromFile<ExcelWorkbook>(importer, " ", importRequest));
        await Assert.ThrowsAsync<ArgumentException>(() =>
            ExcelStreamExtensions.ImportFromFileAsync<ExcelWorkbook>(importer, "\t", importRequest));
    }

    /// <summary>
    /// 验证导出到字节应转发请求并释放目标。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导出到字节应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导出到字节异步应转发请求并释放目标。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导出到字节异步应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从字节应转发请求并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从字节应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从字节异步应转发请求并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从字节异步应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从文件应转发请求并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从文件应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 验证导入从文件异步应转发请求并释放源。
    /// </summary>
    /// <param name="behavior">测试替身要模拟的行为。</param>
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

    /// <summary>
    /// 验证导入从文件异步应保留原始失败。
    /// </summary>
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

    /// <summary>
    /// 创建 Excel 导出请求。
    /// </summary>
    /// <returns>包含 Data 工作表和一行测试数据的导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateExportRequest() =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new ExcelRow { Name = "测试数据" }
        }));

    /// <summary>
    /// 创建 Excel 导入请求。
    /// </summary>
    /// <returns>将 Data 工作表绑定到 Rows 集合的导入请求。</returns>
    private static ExcelWorkbookImportRequest<ExcelWorkbook> CreateImportRequest() =>
        ExcelImport.Workbook<ExcelWorkbook>(builder => builder.Sheet("Data", workbook => workbook.Rows));

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

    /// <summary>
    /// 断言测试替身按预期抛出异常或取消。
    /// </summary>
    /// <param name="importer">Excel 或 CSV 导入器。</param>
    /// <param name="exception">测试期间要传播的异常。</param>
    /// <param name="token">操作令牌。</param>
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
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class ExcelWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ExcelRow> Rows { get; } = new List<ExcelRow>();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ExcelRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 记录测试过程中的状态或调用次数。
    /// </summary>
    private sealed class TrackingExcelExporter : IExcelExporter
    {
        /// <summary>
        /// 获取或设置行为。
        /// </summary>
        public TrackingBehavior Behavior { get; set; }

        /// <summary>
        /// 获取或设置要抛出的异常。
        /// </summary>
        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("Excel fake failure");

        /// <summary>
        /// 获取负载。
        /// </summary>
        public byte[] Payload { get; } = Encoding.UTF8.GetBytes("excel-payload");

        /// <summary>
        /// 获取或设置最近一次请求。
        /// </summary>
        public ExcelWorkbookExportRequest LastRequest { get; private set; }

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
        public void Export(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            LastDestination = destination;
            LastCancellationToken = cancellationToken;
            ApplyBehavior(cancellationToken);
            destination.Write(Payload, 0, Payload.Length);
        }

        /// <inheritdoc />
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

        /// <inheritdoc />
        public void ExportToFile(ExcelWorkbookExportRequest request, string path,
            CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            LastPath = path;
            LastCancellationToken = cancellationToken;
            ApplyBehavior(cancellationToken);
            File.WriteAllBytes(path, Payload);
        }

        /// <inheritdoc />
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
    private sealed class TrackingExcelImporter : IExcelImporter
    {
        /// <summary>
        /// 获取或设置行为。
        /// </summary>
        public TrackingBehavior Behavior { get; set; }

        /// <summary>
        /// 获取或设置要抛出的异常。
        /// </summary>
        public Exception ExceptionToThrow { get; set; } = new InvalidOperationException("Excel fake failure");

        /// <summary>
        /// 获取或设置最近一次结果。
        /// </summary>
        public ExcelWorkbookImportResult<ExcelWorkbook> LastResult { get; private set; }

        /// <summary>
        /// 获取或设置最近一次请求。
        /// </summary>
        public object LastRequest { get; private set; }

        /// <summary>
        /// 获取或设置最近一次路径。
        /// </summary>
        public string LastPath { get; private set; }

        /// <summary>
        /// 获取或设置最近一次源流。
        /// </summary>
        public Stream LastSource { get; private set; }

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
        public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
            where TWorkbook : class, new()
        {
            Record(source, request, cancellationToken);
            ApplyBehavior(cancellationToken);
            return CreateResult<TWorkbook>();
        }

        /// <inheritdoc />
        public async Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
            where TWorkbook : class, new()
        {
            Record(source, request, cancellationToken);
            await Task.Yield();
            ApplyBehavior(cancellationToken);
            return CreateResult<TWorkbook>();
        }

        /// <summary>
        /// 创建测试导入结果。
        /// </summary>
        /// <typeparam name="TWorkbook">泛型参数 TWorkbook 表示方法处理的数据类型。</typeparam>
        /// <returns>包含新工作簿、空工作表结果及空错误集合的导入结果。</returns>
        private ExcelWorkbookImportResult<TWorkbook> CreateResult<TWorkbook>() where TWorkbook : class, new()
        {
            var result = new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(),
                Array.Empty<ExcelSheetImportResult>(), Array.Empty<ExcelImportError>(), false, null);
            if (typeof(TWorkbook) == typeof(ExcelWorkbook))
                LastResult = (ExcelWorkbookImportResult<ExcelWorkbook>)(object)result;
            return result;
        }

        /// <summary>
        /// 记录测试替身调用。
        /// </summary>
        /// <typeparam name="TWorkbook">泛型参数 TWorkbook 表示方法处理的数据类型。</typeparam>
        /// <param name="source">输入流。</param>
        /// <param name="request">导入或导出请求。</param>
        /// <param name="cancellationToken">用于取消异步操作的令牌。</param>
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
