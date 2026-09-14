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

    private static ExcelWorkbookExportRequest CreateRequest(ExcelFormat format, Stream template,
        bool leaveOpen)
    {
        return ExcelExport.Workbook(builder => builder
            .Format(format)
            .UseTemplate(template, leaveOpen)
            .AddSheet("Data", new[] { new TemplateRow { Name = "模板数据", Count = 1 } }));
    }

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

    private sealed class TemplateRow
    {
        public string Name { get; set; }
        public int Count { get; set; }
    }

    private sealed class AsyncOnlyReadStream : Stream
    {
        private readonly MemoryStream _inner;
        private bool _disposed;

        public AsyncOnlyReadStream(byte[] content)
        {
            _inner = new MemoryStream(content, writable: false);
        }

        public int SyncReadCount { get; private set; }
        public int AsyncReadCount { get; private set; }
        public bool IsDisposed => _disposed;
        public override bool CanRead => !_disposed;
        public override bool CanSeek => !_disposed;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            throw new InvalidOperationException("模板读取边界当前使用了同步 Read。");
        }

        public override int Read(Span<byte> buffer)
        {
            SyncReadCount++;
            throw new InvalidOperationException("模板读取边界当前使用了同步 Read。");
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncReadCount++;
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncReadCount++;
            return _inner.ReadAsync(buffer, cancellationToken);
        }

        public override void Flush() { }
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

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

    private sealed class CancellationOnSyncReadStream : Stream
    {
        private readonly MemoryStream _inner;
        private readonly CancellationTokenSource _cancellation;
        private bool _disposed;

        public CancellationOnSyncReadStream(byte[] content, CancellationTokenSource cancellation)
        {
            _inner = new MemoryStream(content, writable: false);
            _cancellation = cancellation;
        }

        public int SyncReadCount { get; private set; }
        public int AsyncReadCount { get; private set; }
        public bool IsDisposed => _disposed;
        public override bool CanRead => !_disposed;
        public override bool CanSeek => !_disposed;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCount++;
            var read = _inner.Read(buffer, offset, count);
            _cancellation.Cancel();
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            SyncReadCount++;
            var read = _inner.Read(buffer);
            _cancellation.Cancel();
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncReadCount++;
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncReadCount++;
            return _inner.ReadAsync(buffer, cancellationToken);
        }

        public override void Flush() { }
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

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
