using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Mappings;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// CSV 异步文件入口集成测试。
/// </summary>
public sealed class CsvAsyncFileIntegrationTest
{
    /// <summary>
    /// 测试 - 真实 UTF-8 CSV 文件应支持特殊字段往返，并原子替换已有目标。
    /// </summary>
    /// <param name="newLine">写入 CSV 文件时使用的换行序列。</param>
    [Theory]
    [InlineData("\r\n")]
    [InlineData("\n")]
    public async Task ExportToFileAsync_ImportFromFileAsync_ShouldRoundTripUtf8AndReplaceTarget(
        string newLine)
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "data.csv");
        try
        {
            using var provider = BuildProvider();
            var exporter = provider.GetRequiredService<ICsvExporter>();
            var importer = provider.GetRequiredService<ICsvImporter>();
            var options = CreateExportOptions(newLine);

            await exporter.ExportToFileAsync(new[]
            {
                new CsvFileRow { Name = "旧目标", Description = "旧描述", Count = 1 }
            }, path, options);
            var originalBytes = File.ReadAllBytes(path);

            var expected = new CsvFileRow
            {
                Name = "客户, \"小王\"",
                Description = $"第一行{newLine}第二行",
                Count = 42
            };
            await exporter.ExportToFileAsync(new[] { expected }, path, options);

            var actualBytes = File.ReadAllBytes(path);
            var content = new UTF8Encoding(false, true).GetString(actualBytes);
            Assert.False(originalBytes.SequenceEqual(actualBytes));
            Assert.Contains(actualBytes, value => value > 0x7F);
            Assert.Contains("\"客户, \"\"小王\"\"\"", content);
            Assert.Contains($"\"第一行{newLine}第二行\"", content);

            var result = await importer.ImportFromFileAsync<CsvFileRow>(path,
                new CsvImportOptions<CsvFileRow>
                {
                    Encoding = new UTF8Encoding(false, true)
                });

            Assert.Empty(result.Errors);
            var actual = Assert.Single(result.Items);
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Description, actual.Description);
            Assert.Equal(expected.Count, actual.Count);
            Assert.Empty(GetAtomicTemporaryFiles(path));
            AssertFileCanBeOpenedExclusively(path);
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 导出枚举首次推进时取消应保留原文件并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ExportToFileAsync_CancellationDuringEnumeration_ShouldPreserveTargetAndCleanupTemporaryFile()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "data.csv");
        var originalBytes = Encoding.UTF8.GetBytes("原始目标\r\n保留内容");
        File.WriteAllBytes(path, originalBytes);
        using var cancellation = new CancellationTokenSource();
        try
        {
            using var provider = BuildProvider();
            var exporter = provider.GetRequiredService<ICsvExporter>();
            var data = CancelOnFirstMoveNext(new[]
            {
                new CsvFileRow { Name = "不会提交", Description = "取消", Count = 7 }
            }, cancellation);

            var exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
                exporter.ExportToFileAsync(data, path, CreateExportOptions("\r\n"), cancellation.Token));

            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.True(originalBytes.SequenceEqual(File.ReadAllBytes(path)));
            Assert.Empty(GetAtomicTemporaryFiles(path));
            AssertFileCanBeOpenedExclusively(path);
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 导出枚举写入一行后抛异常应保留原文件并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ExportToFileAsync_EnumerationFailure_ShouldPreserveTargetAndCleanupTemporaryFile()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "data.csv");
        var originalBytes = Encoding.UTF8.GetBytes("原始目标\r\n保留内容");
        File.WriteAllBytes(path, originalBytes);
        var expectedFailure = new InvalidOperationException("测试 CSV 枚举失败");
        try
        {
            using var provider = BuildProvider();
            var exporter = provider.GetRequiredService<ICsvExporter>();
            var data = YieldThenThrow(new CsvFileRow
            {
                Name = "不会提交",
                Description = "异常",
                Count = 8
            }, expectedFailure);

            var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
                exporter.ExportToFileAsync(data, path, CreateExportOptions("\r\n")));

            Assert.Same(expectedFailure, exception.InnerException);
            Assert.True(originalBytes.SequenceEqual(File.ReadAllBytes(path)));
            Assert.Empty(GetAtomicTemporaryFiles(path));
            AssertFileCanBeOpenedExclusively(path);
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 异步 CSV 的实际输出写入必须走异步流边界，不得回退到同步 Write。
    /// </summary>
    [Fact]
    public async Task ExportAsync_ShouldUseAsyncWriteBoundary()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        using var destination = new AsyncOnlyWriteStream();

        await exporter.ExportAsync(new[]
        {
            new CsvFileRow { Name = "异步写", Description = "边界", Count = 9 }
        }, destination, CreateExportOptions("\r\n"));

        Assert.Equal(0, destination.SyncWriteCount);
        Assert.True(destination.AsyncWriteCount > 0);
        Assert.NotEmpty(destination.ToArray());
    }

    /// <summary>
    /// 测试 - BOM 编码的空 CSV 异步输出应与同步 StreamWriter 产生完整相同字节。
    /// </summary>
    [Fact]
    public async Task ExportAsync_BomEmptyCollection_ShouldMatchSyncStreamWriterBytes()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var encoding = new UTF8Encoding(true, true);
        var options = CreateExportOptions("\r\n", encoding, includeHeader: false);
        using var syncDestination = new MemoryStream();
        using var asyncDestination = new AsyncOnlyWriteStream();

        exporter.Export(Array.Empty<CsvFileRow>(), syncDestination, options);
        await exporter.ExportAsync(Array.Empty<CsvFileRow>(), asyncDestination, options);

        var expected = encoding.GetPreamble();
        Assert.Equal(expected, syncDestination.ToArray());
        Assert.Equal(expected, asyncDestination.ToArray());
        Assert.Equal(0, asyncDestination.SyncWriteCount);
        Assert.Equal(0, asyncDestination.SyncFlushCount);
        Assert.True(asyncDestination.AsyncWriteCount > 0);
        Assert.True(asyncDestination.AsyncFlushCount > 0);
    }

    /// <summary>
    /// 测试 - BOM 编码的正常记录异步输出应与同步 StreamWriter 产生完整相同字节。
    /// </summary>
    [Fact]
    public async Task ExportAsync_BomRecords_ShouldMatchSyncStreamWriterBytes()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var encoding = new UTF8Encoding(true, true);
        var options = CreateExportOptions("\r\n", encoding);
        var rows = new[]
        {
            new CsvFileRow { Name = "Alice", Description = "Report", Count = 42 }
        };
        using var syncDestination = new MemoryStream();
        using var asyncDestination = new AsyncOnlyWriteStream();

        exporter.Export(rows, syncDestination, options);
        await exporter.ExportAsync(rows, asyncDestination, options);

        var expected = encoding.GetPreamble()
            .Concat(encoding.GetBytes("Name,Description,Count\r\nAlice,Report,42\r\n"))
            .ToArray();
        Assert.Equal(expected, syncDestination.ToArray());
        Assert.Equal(expected, asyncDestination.ToArray());
        Assert.Equal(0, asyncDestination.SyncWriteCount);
        Assert.Equal(0, asyncDestination.SyncFlushCount);
        Assert.True(asyncDestination.AsyncWriteCount > 0);
        Assert.True(asyncDestination.AsyncFlushCount > 0);
    }

    /// <summary>
    /// 测试 - BOM 编码写入起始位置非零时不得在已有字节中间插入 BOM。
    /// </summary>
    [Fact]
    public async Task ExportAsync_BomWithNonZeroPosition_ShouldMatchSyncStreamWriterBytes()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var encoding = new UTF8Encoding(true, true);
        var options = CreateExportOptions("\r\n", encoding, includeHeader: false);
        var rows = new[]
        {
            new CsvFileRow { Name = "Alice", Description = "Report", Count = 42 }
        };
        var prefix = new byte[] { 0x41, 0x42 };
        using var syncDestination = new MemoryStream();
        using var asyncDestination = new AsyncOnlyWriteStream(initialBytes: prefix);
        syncDestination.Write(prefix, 0, prefix.Length);

        exporter.Export(rows, syncDestination, options);
        await exporter.ExportAsync(rows, asyncDestination, options);

        var expected = prefix
            .Concat(encoding.GetBytes("Alice,Report,42\r\n"))
            .ToArray();
        Assert.Equal(expected, syncDestination.ToArray());
        Assert.Equal(expected, asyncDestination.ToArray());
        Assert.Equal(0, asyncDestination.SyncWriteCount);
        Assert.Equal(0, asyncDestination.SyncFlushCount);
        Assert.True(asyncDestination.AsyncWriteCount > 0);
        Assert.True(asyncDestination.AsyncFlushCount > 0);
    }

    /// <summary>
    /// 测试 - 超过 StreamWriter 缓冲区的单字段仍必须只走异步 Write 和 Flush。
    /// </summary>
    [Fact]
    public async Task ExportAsync_LargeSingleField_ShouldUseOnlyAsyncWriteAndFlush()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        using var destination = new AsyncOnlyWriteStream();
        var value = new string('X', 20_000);

        await exporter.ExportAsync(new[]
        {
            new CsvFileRow { Name = value, Description = "单字段", Count = 9 }
        }, destination, CreateExportOptions("\r\n"));

        var expected = "Name,Description,Count\r\n"
            + value + ",单字段,9\r\n";
        Assert.Equal(expected, new UTF8Encoding(false, true).GetString(destination.ToArray()));
        Assert.Equal(0, destination.SyncWriteCount);
        Assert.Equal(0, destination.SyncFlushCount);
        Assert.True(destination.AsyncWriteCount > 0);
        Assert.True(destination.AsyncFlushCount > 0);
    }

    /// <summary>
    /// 测试 - 累计超过内部缓冲区的宽记录不得隐式回退到同步 Write。
    /// </summary>
    [Fact]
    public async Task ExportAsync_WideRecord_ShouldUseOnlyAsyncWriteBoundary()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        using var destination = new AsyncOnlyWriteStream();
        var name = new string('N', 2_000);
        var description = new string('D', 2_000);

        await exporter.ExportAsync(new[]
        {
            new CsvFileRow { Name = name, Description = description, Count = 10 }
        }, destination, CreateExportOptions("\r\n"));

        var expected = "Name,Description,Count\r\n"
            + name + "," + description + ",10\r\n";
        Assert.Equal(expected, new UTF8Encoding(false, true).GetString(destination.ToArray()));
        Assert.Equal(0, destination.SyncWriteCount);
        Assert.Equal(0, destination.SyncFlushCount);
        Assert.True(destination.AsyncWriteCount > 0);
        Assert.True(destination.AsyncFlushCount > 0);
    }

    /// <summary>
    /// 测试 - 超大记录在异步写边界取消时应保留原始取消异常，且不调用同步边界。
    /// </summary>
    [Fact]
    public async Task ExportAsync_LargeRecordCancellation_ShouldUseOnlyAsyncBoundary()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        using var cancellation = new CancellationTokenSource();
        using var destination = new AsyncOnlyWriteStream(cancellation);
        var value = new string('C', 20_000);

        var exception = await Assert.ThrowsAsync<OperationCanceledException>(() => exporter.ExportAsync(new[]
        {
            new CsvFileRow { Name = value, Description = "取消", Count = 11 }
        }, destination, CreateExportOptions("\r\n"), cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(0, destination.SyncWriteCount);
        Assert.Equal(0, destination.SyncFlushCount);
        Assert.True(destination.AsyncWriteCount > 0);
    }

    /// <summary>
    /// 测试 - 异步目标流异常应原样传播，且不得回退到同步 Write/Flush。
    /// </summary>
    [Fact]
    public async Task ExportAsync_LargeRecordWriteFailure_ShouldPropagateAsyncExceptionOnly()
    {
        using var provider = BuildProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var expected = new IOException("测试异步写入失败");
        using var destination = new AsyncOnlyWriteStream(asyncWriteException: expected);
        var value = new string('E', 20_000);

        var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() => exporter.ExportAsync(new[]
        {
            new CsvFileRow { Name = value, Description = "异常", Count = 12 }
        }, destination, CreateExportOptions("\r\n")));

        Assert.Same(expected, exception.InnerException);
        Assert.Equal(0, destination.SyncWriteCount);
        Assert.Equal(0, destination.SyncFlushCount);
        Assert.True(destination.AsyncWriteCount > 0);
    }

    /// <summary>
    /// 测试 - 已写入第一条记录后取消仍应保留旧目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ExportToFileAsync_CancellationAfterFirstRecord_ShouldPreserveTargetAndCleanupTemporaryFile()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "data.csv");
        var originalBytes = Encoding.UTF8.GetBytes("原始目标\r\n保留内容");
        File.WriteAllBytes(path, originalBytes);
        using var cancellation = new CancellationTokenSource();
        try
        {
            using var provider = BuildProvider();
            var exporter = provider.GetRequiredService<ICsvExporter>();
            var data = YieldThenCancel(
                new CsvFileRow { Name = "第一条", Description = "已写入", Count = 1 },
                new CsvFileRow { Name = "第二条", Description = "不会提交", Count = 2 },
                cancellation);

            var exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
                exporter.ExportToFileAsync(data, path, CreateExportOptions("\r\n"), cancellation.Token));

            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.Equal(originalBytes, File.ReadAllBytes(path));
            Assert.Empty(GetAtomicTemporaryFiles(path));
            AssertFileCanBeOpenedExclusively(path);
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 构建测试使用的服务提供程序。
    /// </summary>
    /// <returns>已注册默认映射计划工厂的服务提供程序。</returns>
    private static ServiceProvider BuildProvider() =>
        ExcelMappingPlanFactoryProvider.RegisterDefault(new ServiceCollection())
            .BuildServiceProvider();

    /// <summary>
    /// 创建 CSV 导出选项。
    /// </summary>
    /// <param name="newLine">CSV 换行序列。</param>
    /// <param name="encoding">文本编码。</param>
    /// <param name="includeHeader">是否包含表头。</param>
    /// <returns>使用指定换行、编码和表头设置的 CSV 导出选项。</returns>
    private static CsvExportOptions<CsvFileRow> CreateExportOptions(string newLine,
        Encoding encoding = null, bool includeHeader = true) =>
        new CsvExportOptions<CsvFileRow>
        {
            Encoding = encoding ?? new UTF8Encoding(false, true),
            IncludeHeader = includeHeader,
            NewLine = newLine
        };

    /// <summary>
    /// 创建首次移动到下一行即取消的数据序列。
    /// </summary>
    /// <param name="rows">数据行或行集合。</param>
    /// <param name="cancellation">取消令牌。</param>
    /// <returns>每次产出数据行前触发取消的延迟序列。</returns>
    private static IEnumerable<CsvFileRow> CancelOnFirstMoveNext(IEnumerable<CsvFileRow> rows,
        CancellationTokenSource cancellation)
    {
        foreach (var row in rows)
        {
            cancellation.Cancel();
            yield return row;
        }
    }

    /// <summary>
    /// 创建先产生数据再抛出异常的序列。
    /// </summary>
    /// <param name="row">数据行或行集合。</param>
    /// <param name="exception">测试期间要传播的异常。</param>
    /// <returns>先产出指定行、继续枚举时抛出指定异常的序列。</returns>
    private static IEnumerable<CsvFileRow> YieldThenThrow(CsvFileRow row, Exception exception)
    {
        yield return row;
        throw exception;
    }

    /// <summary>
    /// 创建先产生数据再取消的序列。
    /// </summary>
    /// <param name="first">第一个数据项。</param>
    /// <param name="second">第二个数据项。</param>
    /// <param name="cancellation">取消令牌。</param>
    /// <returns>产出第一行后触发取消、再产出第二行的序列。</returns>
    private static IEnumerable<CsvFileRow> YieldThenCancel(CsvFileRow first, CsvFileRow second,
        CancellationTokenSource cancellation)
    {
        yield return first;
        cancellation.Cancel();
        yield return second;
    }

    /// <summary>
    /// 创建临时目录。
    /// </summary>
    /// <returns>新创建的独立测试临时目录路径。</returns>
    private static string CreateTemporaryDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests.Integration",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    /// <summary>
    /// 获取原子提交临时文件。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    /// <returns>目标目录中匹配目标文件临时命名模式的路径数组。</returns>
    private static string[] GetAtomicTemporaryFiles(string path)
    {
        var directory = Path.GetDirectoryName(path);
        return Directory.GetFiles(directory, Path.GetFileName(path) + ".*.tmp");
    }

    /// <summary>
    /// 验证文件可以独占打开。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    private static void AssertFileCanBeOpenedExclusively(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(stream.Length > 0);
    }

    /// <summary>
    /// 删除临时目录。
    /// </summary>
    /// <param name="directory">目标目录路径。</param>
    private static void DeleteTemporaryDirectory(string directory)
    {
        if (Directory.Exists(directory))
            Directory.Delete(directory, true);
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class AsyncOnlyWriteStream : MemoryStream
    {
        /// <summary>
        /// 用于在首次异步写入后取消操作的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancelAfterFirstAsyncWrite;

        /// <summary>
        /// 异步写入时应抛出的测试异常。
        /// </summary>
        private readonly Exception _asyncWriteException;

        /// <summary>
        /// 初始化一个 <see cref="AsyncOnlyWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="cancelAfterFirstAsyncWrite">首次异步写入后触发的取消源；null 表示不取消。</param>
        /// <param name="asyncWriteException">异步写入时抛出的异常；null 表示正常写入。</param>
        /// <param name="initialBytes">预先写入流的字节；null 表示从空流开始。</param>
        public AsyncOnlyWriteStream(CancellationTokenSource cancelAfterFirstAsyncWrite = null,
            Exception asyncWriteException = null, byte[] initialBytes = null)
        {
            _cancelAfterFirstAsyncWrite = cancelAfterFirstAsyncWrite;
            _asyncWriteException = asyncWriteException;
            if (initialBytes != null)
                base.Write(initialBytes, 0, initialBytes.Length);
        }

        /// <summary>
        /// 获取或设置同步写入次数。
        /// </summary>
        public int SyncWriteCount { get; private set; }

        /// <summary>
        /// 获取或设置异步写入次数。
        /// </summary>
        public int AsyncWriteCount { get; private set; }

        /// <summary>
        /// 获取或设置同步刷新次数。
        /// </summary>
        public int SyncFlushCount { get; private set; }

        /// <summary>
        /// 获取或设置异步刷新次数。
        /// </summary>
        public int AsyncFlushCount { get; private set; }

        /// <inheritdoc />
        public override void Flush()
        {
            SyncFlushCount++;
            throw new InvalidOperationException("异步 CSV 不应调用同步 Flush。");
        }

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken)
        {
            AsyncFlushCount++;
            cancellationToken.ThrowIfCancellationRequested();
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            SyncWriteCount++;
            throw new InvalidOperationException("异步 CSV 不应调用同步 Write。");
        }

        /// <inheritdoc />
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            SyncWriteCount++;
            throw new InvalidOperationException("异步 CSV 不应调用同步 Write。");
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncWriteCount++;
            cancellationToken.ThrowIfCancellationRequested();
            if (_asyncWriteException != null)
                throw _asyncWriteException;
            base.Write(buffer, offset, count);
            if (AsyncWriteCount == 1)
                _cancelAfterFirstAsyncWrite?.Cancel();
            cancellationToken.ThrowIfCancellationRequested();
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWriteCount++;
            cancellationToken.ThrowIfCancellationRequested();
            if (_asyncWriteException != null)
                throw _asyncWriteException;
            base.Write(buffer.Span);
            if (AsyncWriteCount == 1)
                _cancelAfterFirstAsyncWrite?.Cancel();
            cancellationToken.ThrowIfCancellationRequested();
            return ValueTask.CompletedTask;
        }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class CsvFileRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置描述。
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }
}
