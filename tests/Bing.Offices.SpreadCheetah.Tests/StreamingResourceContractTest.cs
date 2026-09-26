using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Mappings;
using Xunit;

namespace Bing.Offices.SpreadCheetah.Tests;

/// <summary>
/// 验证前向流式导出的批次、背压、取消、故障和文件提交边界。
/// </summary>
public sealed class StreamingResourceContractTest
{
    /// <summary>
    /// 验证批次边界数据量只触发一次源枚举。
    /// </summary>
    /// <param name="count">测试数据行数。</param>
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(999)]
    [InlineData(1000)]
    [InlineData(1001)]
    public async Task ExportBatchesAsync_ShouldHandleBoundaryCountsWithSingleEnumeration(int count)
    {
        var source = new ProbeEnumerable<ResourceRow>(CreateRows(count));
        var request = CreateRequest(source);
        using var destination = new MemoryStream();

        await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination,
            new ExcelStreamingExportOptions { BatchSize = 1000 });

        Assert.Equal(1, source.EnumeratorCount);
        Assert.Equal(count + 1, source.MoveNextCount);
        using var workbook = Open(destination);
        Assert.Equal(count == 0 ? 0 : count, workbook.GetSheet("Rows").LastRowNum);
    }

    /// <summary>
    /// 验证慢目标流不引发重复数据枚举。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_WithSlowDestination_ShouldCompleteSingleEnumeration()
    {
        var source = new ProbeEnumerable<ResourceRow>(CreateRows(3));
        var request = CreateRequest(source);
        using var destination = new BackpressureStream(() => source.MoveNextCount > 0);
        var exportTask = new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination,
            new ExcelStreamingExportOptions { BatchSize = 1 });

        await destination.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(10));
        // SpreadCheetah 在 Zip 条目收尾前可能把 Sheet XML 留在内部缓冲区；
        // 此处验证慢目标下不会重新枚举或泄漏资源，不宣称目标流能对每行形成即时背压。
        Assert.Equal(1, source.EnumeratorCount);
        Assert.Equal(4, source.MoveNextCount);
        destination.Release.TrySetResult(true);
        await exportTask;
        Assert.Equal(4, source.MoveNextCount);
    }

    /// <summary>
    /// 验证异步写入阻塞时暂停大数据源枚举。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_WithLargeSourceAndAsyncWriteGate_ShouldPauseBeforeSourceCompletes()
    {
        const int totalRows = 100_000;
        const int minimumRowsBeforeGate = 1_000;
        var source = new ProbeEnumerable<ResourceRow>(Enumerable.Range(1, totalRows)
            .Select(index => new ResourceRow
            {
                Id = index,
                Name = $"{new string('x', 192)}-{index}"
            }).ToArray());
        using var destination = new AsyncWriteGateStream(() =>
            source.MoveNextCount > minimumRowsBeforeGate && source.MoveNextCount < totalRows);
        var exportTask = new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(
            CreateRequest(source), destination, new ExcelStreamingExportOptions { BatchSize = 1000 });

        var completed = await Task.WhenAny(destination.Blocked.Task, exportTask,
            Task.Delay(TimeSpan.FromSeconds(30)));
        Assert.Same(destination.Blocked.Task, completed);

        var observedCount = source.MoveNextCount;
        Assert.InRange(observedCount, minimumRowsBeforeGate + 1, totalRows - 1);
        await Task.Delay(100);
        Assert.Equal(observedCount, source.MoveNextCount);

        destination.Release.TrySetResult(true);
        await exportTask;
        Assert.Equal(totalRows + 1, source.MoveNextCount);
        Assert.True(destination.AsyncWriteCount > 0);
    }

    /// <summary>
    /// 验证写入期间取消向上传播且目标流保持打开。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_WhenDestinationCancelsDuringWrite_ShouldPropagateAndLeaveStreamOpen()
    {
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelOnWriteStream(cancellation, triggerAtWrite: 6);
        var request = CreateRequest(CreateRows(50));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination,
                new ExcelStreamingExportOptions { BatchSize = 1 }, cancellation.Token));

        Assert.False(destination.IsDisposed);
    }

    /// <summary>
    /// 验证写入失败向上传播且目标流保持打开。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_WhenDestinationFails_ShouldPropagateWriteFailureAndLeaveStreamOpen()
    {
        using var destination = new FailingWriteStream(failAtWrite: 4);
        var request = CreateRequest(CreateRows(10));

        await Assert.ThrowsAsync<IOException>(() =>
            new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination,
                new ExcelStreamingExportOptions { BatchSize = 1 }));

        Assert.False(destination.IsDisposed);
    }

    /// <summary>
    /// 验证写入中取消保留原文件并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ExportBatchesToFileAsync_WhenCanceledMidWrite_ShouldPreserveSentinelAndCleanTemporaryFile()
    {
        var cancellation = new CancellationTokenSource();
        var converter = new CancelAfterFirstConverter(cancellation);
        var factory = ExcelMappingPlanFactoryProvider.CreateDefault(new[] { converter });
        var exporter = new SpreadCheetahStreamingExcelExporter(factory);
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "result.xlsx");
            var sentinel = new byte[] { 0x53, 0x45, 0x4E, 0x54, 0x49, 0x4E, 0x45, 0x4C };
            await File.WriteAllBytesAsync(path, sentinel);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                exporter.ExportBatchesToFileAsync(CreateRequest(CreateRows(10), converter.Name), path,
                    new ExcelStreamingExportOptions { BatchSize = 1 }, cancellation.Token));

            Assert.Equal(sentinel, await File.ReadAllBytesAsync(path));
            Assert.Empty(Directory.EnumerateFiles(directory).Where(file => !string.Equals(file, path,
                StringComparison.OrdinalIgnoreCase)));
        }
        finally
        {
            cancellation.Dispose();
            Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证枚举失败保留原文件并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ExportBatchesToFileAsync_WhenEnumeratorFails_ShouldPreserveSentinelAndCleanTemporaryFile()
    {
        var exporter = new SpreadCheetahStreamingExcelExporter();
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "result.xlsx");
            var sentinel = new byte[] { 0x53, 0x45, 0x4E, 0x54, 0x49, 0x4E, 0x45, 0x4C };
            await File.WriteAllBytesAsync(path, sentinel);

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                exporter.ExportBatchesToFileAsync(CreateRequest(new ThrowingEnumerable<ResourceRow>()), path,
                    new ExcelStreamingExportOptions { BatchSize = 1 }));

            Assert.Equal(sentinel, await File.ReadAllBytesAsync(path));
            Assert.Empty(Directory.EnumerateFiles(directory).Where(file => !string.Equals(file, path,
                StringComparison.OrdinalIgnoreCase)));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 创建用于导出验证的工作簿请求。
    /// </summary>
    /// <param name="rows">待导出的资源测试数据。</param>
    /// <param name="converterName">名称列使用的转换器名称；null 表示不配置转换器。</param>
    /// <returns>包含测试数据和所需映射的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(IEnumerable<ResourceRow> rows,
        string converterName = null)
    {
        return ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", rows, sheet =>
        {
            if (converterName != null)
            {
                sheet.Mapping(new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new()
                        {
                            PropertyName = nameof(ResourceRow.Name),
                            ConverterName = converterName
                        }
                    }
                });
            }
        }));
    }

    /// <summary>
    /// 创建连续编号的资源测试数据。
    /// </summary>
    /// <param name="count">测试数据行数。</param>
    /// <returns>按编号递增的测试数据集合。</returns>
    private static IReadOnlyList<ResourceRow> CreateRows(int count) => Enumerable.Range(1, count)
        .Select(index => new ResourceRow { Id = index, Name = $"Name {index}" })
        .ToArray();

    /// <summary>
    /// 从导出流的副本打开工作簿。
    /// </summary>
    /// <param name="stream">包含导出工作簿的内存流。</param>
    /// <returns>从独立流副本加载的工作簿。</returns>
    private static NPOI.XSSF.UserModel.XSSFWorkbook Open(MemoryStream stream) =>
        new(new MemoryStream(stream.ToArray()));

    /// <summary>
    /// 用于流式资源边界测试的数据行。
    /// </summary>
    private sealed class ResourceRow
    {
        /// <summary>
        /// 获取或设置测试数据编号。
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 获取或设置测试数据名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 记录创建和推进次数的资源测试序列。
    /// </summary>
    /// <typeparam name="T">测试序列的元素类型。</typeparam>
    private sealed class ProbeEnumerable<T> : IEnumerable<T>
    {
        /// <summary>
        /// 固定保存供测试枚举的数据快照。
        /// </summary>
        private readonly IReadOnlyList<T> _items;

        /// <summary>
        /// 初始化一个 <see cref="ProbeEnumerable{T}"/> 类型的实例。
        /// </summary>
        /// <param name="items">待枚举的测试数据。</param>
        public ProbeEnumerable(IEnumerable<T> items) => _items = items.ToArray();

        /// <summary>
        /// 获取或设置枚举器创建次数。
        /// </summary>
        public int EnumeratorCount { get; private set; }
        /// <summary>
        /// 获取或设置枚举推进次数。
        /// </summary>
        public int MoveNextCount { get; private set; }

        /// <inheritdoc />
        public IEnumerator<T> GetEnumerator()
        {
            EnumeratorCount++;
            for (var index = 0; index < _items.Count; index++)
            {
                MoveNextCount++;
                yield return _items[index];
            }
            MoveNextCount++;
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// 创建枚举器时抛出异常的测试序列。
    /// </summary>
    /// <typeparam name="T">测试序列的元素类型。</typeparam>
    private sealed class ThrowingEnumerable<T> : IEnumerable<T>
    {
        /// <inheritdoc />
        public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException("测试枚举器故障。");

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// 首次导出转换时触发取消的测试转换器。
    /// </summary>
    private sealed class CancelAfterFirstConverter : INamedExcelValueConverter
    {
        /// <summary>
        /// 用于从测试扩展中触发导出取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;
        /// <summary>
        /// 记录首次转换是否已触发取消的原子状态。
        /// </summary>
        private int _called;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstConverter"/> 类型的实例。
        /// </summary>
        /// <param name="cancellation">用于触发取消的令牌源。</param>
        public CancelAfterFirstConverter(CancellationTokenSource cancellation) => _cancellation = cancellation;

        /// <inheritdoc />
        public string Name => "cancel-after-first";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return false;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            if (Interlocked.Exchange(ref _called, 1) == 0)
                _cancellation.Cancel();
            return true;
        }
    }

    /// <summary>
    /// 可阻塞一次写入的慢目标测试流。
    /// </summary>
    private sealed class BackpressureStream : Stream
    {
        /// <summary>
        /// 保存测试输出并随包装器释放的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();
        /// <summary>
        /// 判断当前写入是否应进入阻塞的测试条件。
        /// </summary>
        private readonly Func<bool> _shouldBlock;
        /// <summary>
        /// 确保阻塞仅触发一次的原子状态。
        /// </summary>
        private int _blocked;

        /// <summary>
        /// 初始化一个 <see cref="BackpressureStream"/> 类型的实例。
        /// </summary>
        /// <param name="shouldBlock">判断写入是否应阻塞的委托。</param>
        public BackpressureStream(Func<bool> shouldBlock) => _shouldBlock = shouldBlock;

        /// <summary>
        /// 获取通知写入已进入阻塞的完成源。
        /// </summary>
        public TaskCompletionSource<bool> Blocked { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        /// <summary>
        /// 获取允许被阻塞写入继续执行的完成源。
        /// </summary>
        public TaskCompletionSource<bool> Release { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

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
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            WaitIfBlocked();
            _inner.Write(buffer, offset, count);
        }

        /// <inheritdoc />
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            await WaitIfBlockedAsync(cancellationToken).ConfigureAwait(false);
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>
        /// 条件满足时等待测试解除写入阻塞。
        /// </summary>
        private void WaitIfBlocked()
        {
            if (!_shouldBlock() || Interlocked.CompareExchange(ref _blocked, 1, 0) != 0)
                return;
            Blocked.TrySetResult(true);
            Release.Task.GetAwaiter().GetResult();
        }

        /// <summary>
        /// 条件满足时异步等待测试解除写入阻塞。
        /// </summary>
        /// <param name="cancellationToken">取消等待或写入的令牌。</param>
        private async Task WaitIfBlockedAsync(CancellationToken cancellationToken)
        {
            if (!_shouldBlock() || Interlocked.CompareExchange(ref _blocked, 1, 0) != 0)
                return;
            Blocked.TrySetResult(true);
            await Release.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// 通过异步写入闸门验证数据源暂停的测试流。
    /// </summary>
    private sealed class AsyncWriteGateStream : Stream
    {
        /// <summary>
        /// 保存测试输出并随包装器释放的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();
        /// <summary>
        /// 判断当前写入是否应进入阻塞的测试条件。
        /// </summary>
        private readonly Func<bool> _shouldBlock;
        /// <summary>
        /// 确保阻塞仅触发一次的原子状态。
        /// </summary>
        private int _blocked;
        /// <summary>
        /// 已执行异步写入调用的原子计数。
        /// </summary>
        private int _asyncWriteCount;

        /// <summary>
        /// 初始化一个 <see cref="AsyncWriteGateStream"/> 类型的实例。
        /// </summary>
        /// <param name="shouldBlock">判断写入是否应阻塞的委托。</param>
        public AsyncWriteGateStream(Func<bool> shouldBlock) => _shouldBlock = shouldBlock;

        /// <summary>
        /// 获取通知写入已进入阻塞的完成源。
        /// </summary>
        public TaskCompletionSource<bool> Blocked { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        /// <summary>
        /// 获取允许被阻塞写入继续执行的完成源。
        /// </summary>
        public TaskCompletionSource<bool> Release { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        /// <summary>
        /// 获取异步写入次数。
        /// </summary>
        public int AsyncWriteCount => Volatile.Read(ref _asyncWriteCount);

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
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            Interlocked.Increment(ref _asyncWriteCount);
            if (_shouldBlock() && Interlocked.CompareExchange(ref _blocked, 1, 0) == 0)
                return WaitThenWriteAsync(buffer, cancellationToken);
            return _inner.WriteAsync(buffer, cancellationToken);
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
            => WriteAsync(new ReadOnlyMemory<byte>(buffer, offset, count), cancellationToken).AsTask();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>
        /// 异步等待闸门放行后写入数据。
        /// </summary>
        /// <param name="buffer">等待结束后写入的数据。</param>
        /// <param name="cancellationToken">取消等待或写入的令牌。</param>
        private async ValueTask WaitThenWriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken)
        {
            Blocked.TrySetResult(true);
            await Release.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// 在指定写入次数触发取消的测试流。
    /// </summary>
    private sealed class CancelOnWriteStream : Stream
    {
        /// <summary>
        /// 保存测试输出并随包装器释放的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();
        /// <summary>
        /// 用于从测试扩展中触发导出取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;
        /// <summary>
        /// 首次触发取消的写入调用序号。
        /// </summary>
        private readonly int _triggerAtWrite;
        /// <summary>
        /// 已执行写入调用的原子计数。
        /// </summary>
        private int _writeCount;

        /// <summary>
        /// 初始化一个 <see cref="CancelOnWriteStream"/> 类型的实例。
        /// </summary>
        /// <param name="cancellation">用于触发取消的令牌源。</param>
        /// <param name="triggerAtWrite">触发取消的写入调用序号。</param>
        public CancelOnWriteStream(CancellationTokenSource cancellation, int triggerAtWrite)
        {
            _cancellation = cancellation;
            _triggerAtWrite = triggerAtWrite;
        }

        /// <summary>
        /// 获取或设置流是否已被释放。
        /// </summary>
        public bool IsDisposed { get; private set; }
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
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            CancelIfRequired();
            _inner.Write(buffer, offset, count);
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            CancelIfRequired();
            return _inner.WriteAsync(buffer, cancellationToken);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>
        /// 达到指定写入次数时触发取消。
        /// </summary>
        private void CancelIfRequired()
        {
            if (Interlocked.Increment(ref _writeCount) == _triggerAtWrite)
                _cancellation.Cancel();
        }
    }

    /// <summary>
    /// 从指定写入次数起抛出故障的测试流。
    /// </summary>
    private sealed class FailingWriteStream : Stream
    {
        /// <summary>
        /// 保存测试输出并随包装器释放的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();
        /// <summary>
        /// 开始抛出写入故障的调用序号。
        /// </summary>
        private readonly int _failAtWrite;
        /// <summary>
        /// 已执行写入调用的原子计数。
        /// </summary>
        private int _writeCount;

        /// <summary>
        /// 初始化一个 <see cref="FailingWriteStream"/> 类型的实例。
        /// </summary>
        /// <param name="failAtWrite">开始抛出故障的写入调用序号。</param>
        public FailingWriteStream(int failAtWrite) => _failAtWrite = failAtWrite;

        /// <summary>
        /// 获取或设置流是否已被释放。
        /// </summary>
        public bool IsDisposed { get; private set; }
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
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            ThrowIfRequired();
            _inner.Write(buffer, offset, count);
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            ThrowIfRequired();
            return _inner.WriteAsync(buffer, cancellationToken);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>
        /// 达到指定写入次数时抛出模拟故障。
        /// </summary>
        private void ThrowIfRequired()
        {
            if (Interlocked.Increment(ref _writeCount) >= _failAtWrite)
                throw new IOException("测试输出流故障。");
        }
    }
}
