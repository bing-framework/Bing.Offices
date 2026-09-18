using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Bing.Offices.Mappings;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 使用 Core CSV 实现验证异步读写边界。
/// </summary>
public sealed class AsyncPipelineCsvTest
{
    /// <summary>
    /// 验证 CSV 异步读写使用异步流路径并产生与同步流程一致的结果。
    /// </summary>
    [Fact]
    public async Task CsvAsync_ShouldUseAsyncReadWrite_AndMatchSyncResult()
    {
        using var provider = ExcelMappingPlanFactoryProvider.RegisterDefault(new ServiceCollection())
            .BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var rows = new[] { new CsvRow { Code = "A", Count = 1 }, new CsvRow { Code = "B", Count = 2 } };

        using var syncOutput = new MemoryStream();
        exporter.Export(rows, syncOutput);
        using var asyncInput = new AsyncOnlyReadStream(syncOutput.ToArray());
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

    /// <summary>
    /// 验证 CSV 异步导出不会调用调用方流的同步刷新。
    /// </summary>
    [Fact]
    public async Task CsvAsync_ShouldNotUseSynchronousFlushOnCallerStream()
    {
        using var provider = ExcelMappingPlanFactoryProvider.RegisterDefault(new ServiceCollection())
            .BuildServiceProvider();
        var exporter = provider.GetRequiredService<ICsvExporter>();
        using var output = new AsyncOnlyWriteStream(throwOnSyncFlush: true);

        await exporter.ExportAsync(new[] { new CsvRow { Code = "A", Count = 1 } }, output);

        Assert.NotEmpty(output.ToArray());
        Assert.Equal(0, output.SyncFlushCount);
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
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class AsyncOnlyReadStream : Stream
    {
        /// <summary>
        /// 承载异步读取测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个 <see cref="AsyncOnlyReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="bytes">作为只读输入的字节数据。</param>
        public AsyncOnlyReadStream(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);

        /// <summary>
        /// 获取或设置同步读取次数。
        /// </summary>
        public int SyncReadCount { get; private set; }
        /// <summary>
        /// 获取或设置异步读取次数。
        /// </summary>
        public int AsyncReadCount { get; private set; }
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
        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            throw new InvalidOperationException("同步 Read 不允许用于 Async API。");
        }

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncReadCount++;
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncReadCount++;
            return _inner.ReadAsync(buffer, cancellationToken);
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
        /// 指示同步刷新是否应抛出异常。
        /// </summary>
        private readonly bool _throwOnSyncFlush;

        /// <summary>
        /// 初始化一个 <see cref="AsyncOnlyWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="throwOnSyncFlush">是否在同步刷新时抛出异常，默认为 false。</param>
        public AsyncOnlyWriteStream(bool throwOnSyncFlush = false) => _throwOnSyncFlush = throwOnSyncFlush;

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
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            SyncWriteCount++;
            throw new InvalidOperationException("同步 Write 不允许用于 Async API。");
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncWriteCount++;
            return _inner.WriteAsync(buffer, offset, count, cancellationToken);
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWriteCount++;
            return _inner.WriteAsync(buffer, cancellationToken);
        }

        /// <summary>
        /// 获取测试流当前缓冲区中的全部字节。
        /// </summary>
        /// <returns>当前缓冲区内容的副本。</returns>
        public byte[] ToArray() => _inner.ToArray();
        /// <inheritdoc />
        public override void Flush()
        {
            SyncFlushCount++;
            if (_throwOnSyncFlush)
                throw new InvalidOperationException("同步 Flush 不允许用于 Async API。");
        }

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) =>
            _inner.FlushAsync(cancellationToken);
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
}
