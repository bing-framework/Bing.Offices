using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML Forward 实体批次在暂存不可定位输入时执行资源限制和取消清理。
/// </summary>
[Collection("ClosedXml Forward temporary files")]
public sealed class EntityForwardResourceTest
{
    /// <summary>
    /// 缓存一次生成的较大真实 XLSX 输入，供同步和异步缓冲用例复用。
    /// </summary>
    private static readonly byte[] LargeWorkbookBytes = CreateLargeWorkbook();

    /// <summary>
    /// 验证同步和异步不可定位流超过输入字节上限时保留流所有权并清理暂存。
    /// </summary>
    /// <param name="asynchronous">是否使用异步 Forward 入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MaxInputBytes_ShouldRejectBufferedNonSeekableInput(bool asynchronous)
    {
        Assert.True(LargeWorkbookBytes.Length > 81_920);
        var before = ForwardTemporaryFiles();
        using var source = new NonSeekableInputStream(LargeWorkbookBytes);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        const int maximum = 1024;
        var options = BatchOptions(maximum);
        var callbackCount = 0;

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            Import(importer, source, options, asynchronous, _ => callbackCount++));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(source.BytesRead > maximum);
        Assert.True(source.BytesRead < LargeWorkbookBytes.Length);
        Assert.Equal(0, callbackCount);
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证同步和异步输入暂存期间取消会释放临时文件且不关闭调用方流。
    /// </summary>
    /// <param name="asynchronous">是否使用异步 Forward 入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDuringBuffer_ShouldCleanupAndLeaveSourceOpen(bool asynchronous)
    {
        var before = ForwardTemporaryFiles();
        using var cancellation = new CancellationTokenSource();
        using var source = new NonSeekableInputStream(LargeWorkbookBytes, cancellation);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var callbackCount = 0;

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            Import(importer, source, BatchOptions(null), asynchronous, _ => callbackCount++,
                cancellation.Token));

        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.True(source.BytesRead > 0);
        Assert.Equal(0, callbackCount);
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证 Buffered 摘要模式可从不可定位输入流完成导入并清理暂存。
    /// </summary>
    /// <param name="asynchronous">是否使用异步实体批次入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummary_ShouldReadNonSeekableInput(bool asynchronous)
    {
        var before = ForwardTemporaryFiles();
        using var source = new NonSeekableInputStream(LargeWorkbookBytes);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var callbackCount = 0;

        var summary = await Import(importer, source, BufferedSummaryOptions(null), asynchronous,
            _ => callbackCount++);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(4000, summary.RowsRead);
        Assert.Equal(4000, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        Assert.Empty(summary.Result.Entity.Rows);
        Assert.True(callbackCount > 0);
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证 Buffered 摘要模式在不可定位输入超过字节上限时停止读取并清理暂存。
    /// </summary>
    /// <param name="asynchronous">是否使用异步实体批次入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummary_ShouldRejectNonSeekableInputAboveLimit(bool asynchronous)
    {
        Assert.True(LargeWorkbookBytes.Length > 1024);
        var before = ForwardTemporaryFiles();
        using var source = new NonSeekableInputStream(LargeWorkbookBytes);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            Import(importer, source, BufferedSummaryOptions(1024), asynchronous, _ => { }));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(source.BytesRead > 1024);
        Assert.True(source.BytesRead < LargeWorkbookBytes.Length);
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证 Buffered 摘要模式在不可定位输入暂存期间取消时不交付批次。
    /// </summary>
    /// <param name="asynchronous">是否使用异步实体批次入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummary_ShouldCleanupAfterNonSeekableBufferCancellation(bool asynchronous)
    {
        var before = ForwardTemporaryFiles();
        using var cancellation = new CancellationTokenSource();
        using var source = new NonSeekableInputStream(LargeWorkbookBytes, cancellation);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var callbackCount = 0;

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            Import(importer, source, BufferedSummaryOptions(null), asynchronous,
                _ => callbackCount++, cancellation.Token));

        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.True(source.BytesRead > 0);
        Assert.Equal(0, callbackCount);
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 创建含多行唯一文本的大型 XLSX 工作簿。
    /// </summary>
    /// <returns>可复用于导入测试的 XLSX 文件字节。</returns>
    private static byte[] CreateLargeWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Note";
        var random = new Random(7319);
        const string alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        for (var row = 2; row <= 4001; row++)
        {
            sheet.Cell(row, 1).Value = $"C{row:D5}";
            var characters = new char[96];
            for (var index = 0; index < characters.Length; index++)
                characters[index] = alphabet[random.Next(alphabet.Length)];
            sheet.Cell(row, 2).Value = new string(characters);
        }
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// 创建仅使用单列表 Forward 解析并应用输入字节上限的选项。
    /// </summary>
    /// <param name="maximumInputBytes">最大输入字节数；为空时不限制。</param>
    /// <returns>配置好的前向批次导入选项。</returns>
    private static ExcelEntityBatchImportOptions BatchOptions(long? maximumInputBytes) =>
        new(128, new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = maximumInputBytes
        }), ExcelEntityBatchParsingMode.Forward);

    /// <summary>
    /// 创建使用 Buffered 入口但只保留摘要结果的批次选项。
    /// </summary>
    /// <param name="maximumInputBytes">最大输入字节数；为空时不限制。</param>
    /// <returns>配置好的 Buffered 摘要批次选项。</returns>
    private static ExcelEntityBatchImportOptions BufferedSummaryOptions(long? maximumInputBytes) =>
        new(128, new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = maximumInputBytes
        }), ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

    /// <summary>
    /// 按同步或异步入口执行前向实体批次导入。
    /// </summary>
    /// <param name="importer">实体批次导入器。</param>
    /// <param name="source">不可定位的大型 XLSX 输入流。</param>
    /// <param name="options">输入缓冲资源选项。</param>
    /// <param name="asynchronous">是否执行异步入口。</param>
    /// <param name="onBatch">接收批次的同步回调。</param>
    /// <param name="token">导入取消令牌。</param>
    /// <returns>实体批次摘要。</returns>
    private static async Task<ExcelEntityBatchImportSummary<Root>> Import(
        IExcelEntityBatchImporter importer, Stream source, ExcelEntityBatchImportOptions options,
        bool asynchronous, Action<ExcelEntityImportBatch<Root>> onBatch,
        CancellationToken token = default)
    {
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
        if (asynchronous)
        {
            return await importer.ImportEntityBatchesAsync(source, layout, options,
                (batch, _) =>
                {
                    onBatch(batch);
                    return Task.CompletedTask;
                }, token).ConfigureAwait(false);
        }

        return importer.ImportEntityBatches(source, layout, options, onBatch, token);
    }

    /// <summary>
    /// 获取当前 Forward 输入暂存文件清单，用于确认失败路径已清理新文件。
    /// </summary>
    /// <returns>按路径序排序的 Forward 暂存文件路径。</returns>
    private static string[] ForwardTemporaryFiles() =>
        Directory.GetFiles(Path.GetTempPath(), "bing-offices-forward-*.xlsx")
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();

    /// <summary>
    /// 保存前向资源测试中的根对象。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取输入明细集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 保存前向资源测试中的明细行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置唯一编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置随机文本。
        /// </summary>
        public string Note { get; set; }
    }

    /// <summary>
    /// 包装非定位输入流，并可在首次读取后取消关联令牌。
    /// </summary>
    private sealed class NonSeekableInputStream : Stream
    {
        /// <summary>
        /// 保存实际 XLSX 数据的只读流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 首次读取完成后触发取消的可选令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancelAfterRead;

        /// <summary>
        /// 标记首次读取取消回调是否已执行。
        /// </summary>
        private bool _canceled;

        /// <summary>
        /// 初始化一个 <see cref="NonSeekableInputStream" /> 类型的实例。
        /// </summary>
        /// <param name="bytes">XLSX 输入数据。</param>
        /// <param name="cancelAfterRead">读取一块数据后触发取消的令牌源。</param>
        internal NonSeekableInputStream(byte[] bytes,
            CancellationTokenSource cancelAfterRead = null)
        {
            _inner = new MemoryStream(bytes, writable: false);
            _cancelAfterRead = cancelAfterRead;
        }

        /// <summary>
        /// 获取已从调用方输入流读取的字节数。
        /// </summary>
        public long BytesRead { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;

        /// <inheritdoc />
        public override bool CanSeek => false;

        /// <inheritdoc />
        public override bool CanWrite => false;

        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => throw new NotSupportedException();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            AfterRead(read);
            return read;
        }

        /// <inheritdoc />
        public override int Read(Span<byte> buffer)
        {
            var read = _inner.Read(buffer);
            AfterRead(read);
            return read;
        }

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = Read(buffer, offset, count);
            return Task.FromResult(read);
        }

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = Read(buffer.Span);
            return ValueTask.FromResult(read);
        }

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <summary>
        /// 记录读取进度并按需触发第一次读取后的取消。
        /// </summary>
        /// <param name="count">本次实际读取的字节数。</param>
        private void AfterRead(int count)
        {
            BytesRead += count;
            if (count > 0 && !_canceled && _cancelAfterRead != null)
            {
                _canceled = true;
                _cancelAfterRead.Cancel();
            }
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
