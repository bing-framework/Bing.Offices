using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 记录 Excel ExportAsync 当前模板读取边界的特征测试。
/// </summary>
[Collection("Excel Async staging")]
public sealed class TemplateAsyncBoundaryTest
{
    /// <summary>
    /// 测试 - XLS/XLSX 模板在 ExportAsync 中当前通过同步 Read 读取。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportAsync_TemplateRead_ShouldRemainSynchronous(ExcelFormat format)
    {
        using var template = new AsyncOnlyReadStream(CreateTemplateBytes(format));
        var request = CreateRequest(format, template, leaveOpen: true);
        using var destination = new MemoryStream();

        var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
            new NpoiExcelExporter().ExportAsync(request, destination));

        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.True(template.SyncReadCount > 0);
        Assert.Equal(0, template.AsyncReadCount);
        Assert.False(template.IsDisposed);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 测试 - 模板读取失败时，ExportAsync 应按 leaveOpen 语义处理模板流。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="leaveOpen">是否保留模板流打开。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xls, false)]
    [InlineData(ExcelFormat.Xlsx, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    public async Task ExportAsync_TemplateReadFailure_ShouldHonorLeaveOpen(ExcelFormat format, bool leaveOpen)
    {
        using var template = new AsyncOnlyReadStream(CreateTemplateBytes(format));
        var request = CreateRequest(format, template, leaveOpen);
        using var destination = new MemoryStream();

        var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
            new NpoiExcelExporter().ExportAsync(request, destination));

        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal(!leaveOpen, template.IsDisposed);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 测试 - 模板读取期间取消时，XLS/XLSX ExportAsync 应保留取消异常并按 leaveOpen 释放模板。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="leaveOpen">是否保留模板流打开。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xls, false)]
    [InlineData(ExcelFormat.Xlsx, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    public async Task ExportAsync_TemplateReadCancellation_ShouldHonorLeaveOpen(
        ExcelFormat format, bool leaveOpen)
    {
        using var cancellation = new CancellationTokenSource();
        using var template = new CancellationOnSyncReadStream(CreateTemplateBytes(format), cancellation);
        var request = CreateRequest(format, template, leaveOpen);
        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new NpoiExcelExporter().ExportAsync(request, destination, cancellation.Token));

        Assert.True(template.SyncReadCount > 0);
        Assert.Equal(0, template.AsyncReadCount);
        Assert.Equal(!leaveOpen, template.IsDisposed);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 测试 - 调用前已取消时，ExportAsync 应不读取模板且不进入请求级释放逻辑。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportAsync_PreCancelled_ShouldNotReadOrDisposeTemplate(ExcelFormat format)
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var template = new AsyncOnlyReadStream(CreateTemplateBytes(format));
        var request = CreateRequest(format, template, leaveOpen: false);
        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new NpoiExcelExporter().ExportAsync(request, destination, cancellation.Token));

        Assert.Equal(0, template.SyncReadCount);
        Assert.Equal(0, template.AsyncReadCount);
        Assert.False(template.IsDisposed);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 测试 - 未取消的模板导出成功后，模板流仅在 leaveOpen 为 false 时释放。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="leaveOpen">是否保留模板流打开。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xls, false)]
    [InlineData(ExcelFormat.Xlsx, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    public async Task ExportAsync_TemplateSuccess_ShouldHonorLeaveOpen(ExcelFormat format, bool leaveOpen)
    {
        using var template = new MemoryStream(CreateTemplateBytes(format));
        var request = CreateRequest(format, template, leaveOpen);
        using var destination = new MemoryStream();

        await new NpoiExcelExporter().ExportAsync(request, destination);

        Assert.True(destination.Length > 0);
        Assert.Equal(leaveOpen, template.CanRead);
    }

    /// <summary>
    /// 创建 Excel 导出请求。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
    /// <param name="template">模板对象。</param>
    /// <param name="leaveOpen">是否保持流打开。</param>
    /// <returns>生成的 ExcelWorkbookExportRequest 结果。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(ExcelFormat format, Stream template,
        bool leaveOpen)
    {
        return ExcelExport.Workbook(builder => builder
            .Format(format)
            .UseTemplate(template, leaveOpen)
            .AddSheet("Data", new[] { new TemplateRow { Name = "模板数据", Count = 1 } }));
    }

    /// <summary>
    /// 创建模板文件字节内容。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
    /// <returns>生成的字节内容。</returns>
    private static byte[] CreateTemplateBytes(ExcelFormat format)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        workbook.CreateSheet("Data");
        using var destination = new MemoryStream();
        workbook.Write(destination, false);
        return destination.ToArray();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class TemplateRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
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
        /// 指示流是否已释放。
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// 初始化一个 <see cref="AsyncOnlyReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="content">作为只读输入的模板字节。</param>
        public AsyncOnlyReadStream(byte[] content)
        {
            _inner = new MemoryStream(content, writable: false);
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
        /// 获取是否已释放。
        /// </summary>
        public bool IsDisposed => _disposed;
        /// <inheritdoc />
        public override bool CanRead => !_disposed;
        /// <inheritdoc />
        public override bool CanSeek => !_disposed;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            throw new InvalidOperationException("模板读取边界当前使用了同步 Read。");
        }

        /// <inheritdoc />
        public override int Read(Span<byte> buffer)
        {
            SyncReadCount++;
            throw new InvalidOperationException("模板读取边界当前使用了同步 Read。");
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
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        /// <inheritdoc />
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _disposed = true;
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class CancellationOnSyncReadStream : Stream
    {
        /// <summary>
        /// 承载同步读取测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 用于在同步读取时取消操作的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 指示流是否已释放。
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// 初始化一个 <see cref="CancellationOnSyncReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="content">作为只读输入的模板字节。</param>
        /// <param name="cancellation">同步读取期间触发取消的令牌源。</param>
        public CancellationOnSyncReadStream(byte[] content, CancellationTokenSource cancellation)
        {
            _inner = new MemoryStream(content, writable: false);
            _cancellation = cancellation;
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
        /// 获取是否已释放。
        /// </summary>
        public bool IsDisposed => _disposed;
        /// <inheritdoc />
        public override bool CanRead => !_disposed;
        /// <inheritdoc />
        public override bool CanSeek => !_disposed;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            var read = _inner.Read(buffer, offset, count);
            _cancellation.Cancel();
            return read;
        }

        /// <inheritdoc />
        public override int Read(Span<byte> buffer)
        {
            SyncReadCount++;
            var read = _inner.Read(buffer);
            _cancellation.Cancel();
            return read;
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
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _disposed = true;
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
