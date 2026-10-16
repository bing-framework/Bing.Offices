using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 对真实 Entity 批次入口验证临时文件聚合预算。
/// </summary>
[Collection("Entity temporary storage")]
public abstract class EntityTemporaryStorageBudgetContractTest
{
    private static readonly string[] SharedValues = { "Code", "Note", "A01", "中文", "B02", "second" };

    protected abstract string ProviderName { get; }
    protected abstract ServiceProvider CreateProvider();

    /// <summary>
    /// 同步与异步、Complete 与两种 SummaryOnly 入口均按相同字节边界计费。
    /// </summary>
    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public async Task Xlsx_ShouldShareInputAndStringStorageBudget(bool asynchronous, int mode)
    {
        var bytes = CreateXlsx();
        var required = bytes.LongLength + SharedValues.Sum(value => Encoding.UTF8.GetByteCount(value) + 12L);
        var before = TemporaryFiles();
        using var provider = CreateProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using (var source = new Input(bytes, seekable: false))
        {
            var rows = new List<Row>();
            var result = await Import(importer, source, Options(required, mode), asynchronous,
                batch => rows.AddRange(batch.Items.Cast<Row>()));
            Assert.True(result.IsSuccess);
            Assert.Equal(new[] { "A01", "B02" }, rows.Select(row => row.Code));
            Assert.Equal(new[] { "中文", "second" }, rows.Select(row => row.Note));
            Assert.Equal(2, result.RowsRead);
            Assert.Equal(mode == 0 ? 2 : 0, result.Result.Entity.Rows.Count);
            Assert.Equal(bytes.Length, source.BytesRead);
            Assert.True(source.CanRead);
            if (asynchronous)
                Assert.True(source.AsyncReads > 0);
        }
        Assert.Equal(before, TemporaryFiles());

        using (var source = new Input(bytes, seekable: false))
        {
            var calls = 0;
            var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                Import(importer, source, Options(required - 1, mode), asynchronous, _ => calls++));
            AssertResourceException(exception);
            Assert.Equal(0, calls);
            Assert.True(source.CanRead);
        }
        Assert.Equal(before, TemporaryFiles());
    }

    /// <summary>
    /// 可定位源由调用方持有，不计入受管临时存储。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SeekableInput_ShouldChargeOnlySharedStrings(bool asynchronous)
    {
        var required = SharedValues.Sum(value => Encoding.UTF8.GetByteCount(value) + 12L);
        using var provider = CreateProvider();
        using var source = new Input(CreateXlsx(), seekable: true);
        var result = await Import(provider.GetRequiredService<IExcelEntityBatchImporter>(), source,
            Options(required, 1), asynchronous, _ => { });
        Assert.True(result.IsSuccess);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 配置不能被忽略；未纳入预算的路径在读取前拒绝。
    /// </summary>
    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public async Task UnsupportedCombination_ShouldRejectBeforeReading(bool asynchronous, int combination)
    {
        using var provider = CreateProvider();
        using var source = new Input(CreateXlsx(), seekable: false);
        using var templateStream = new Input(CreateXlsx(), seekable: false);
        using var output = new MemoryStream();
        var importOptions = combination == 1
            ? new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
                ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
                new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = output
            })
            : null;
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            combination == 0 ? ExcelEntityBatchParsingMode.Buffered : ExcelEntityBatchParsingMode.Forward,
            ExcelEntityBatchResultMode.Complete, 1_000_000);
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(async () =>
        {
            if (combination == 2)
            {
                var template = new ExcelEntityTemplateOptions(templateStream, leaveOpen: true);
                if (asynchronous)
                    await importer.ImportForTemplateBatchesAsync(source, Layout(), template, options,
                        (_, _) => Task.CompletedTask);
                else
                    importer.ImportForTemplateBatches(source, Layout(), template, options, _ => { });
            }
            else
                await Import(importer, source, options, asynchronous, _ => { });
        });
        Assert.Equal(ProviderName, exception.Provider);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(0, source.BytesRead);
        Assert.Equal(0, templateStream.BytesRead);
        Assert.Equal(0, output.Length);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 回调异常或取消后释放输入与字符串文件，保持原始异常和调用方所有权。
    /// </summary>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task CallbackFailure_ShouldCleanup(bool asynchronous, bool cancel)
    {
        var before = TemporaryFiles();
        using var provider = CreateProvider();
        using var source = new Input(CreateXlsx(), seekable: false);
        using var cancellation = new CancellationTokenSource();
        var expected = new InvalidOperationException("callback failure");
        Action<ExcelEntityImportBatch<Root>> callback = _ =>
        {
            if (cancel)
            {
                cancellation.Cancel();
                cancellation.Token.ThrowIfCancellationRequested();
            }
            throw expected;
        };
        var error = await Record.ExceptionAsync(() => Import(
            provider.GetRequiredService<IExcelEntityBatchImporter>(), source, Options(1_000_000, 1),
            asynchronous, callback, cancellation.Token));
        if (cancel)
            Assert.Equal(cancellation.Token, Assert.IsAssignableFrom<OperationCanceledException>(error).CancellationToken);
        else
            Assert.Same(expected, error);
        Assert.True(source.CanRead);
        Assert.Equal(before, TemporaryFiles());
    }

    /// <summary>
    /// 输入 IO 已开始后取消时不交付批次，释放预算文件并保持源流所有权。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InputCancellation_ShouldCleanupBudgetedFiles(bool asynchronous)
    {
        var before = TemporaryFiles();
        using var provider = CreateProvider();
        using var cancellation = new CancellationTokenSource();
        using var source = new Input(CreateXlsx(), seekable: false, cancellation);
        var calls = 0;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Import(
            provider.GetRequiredService<IExcelEntityBatchImporter>(), source, Options(1_000_000),
            asynchronous, _ => calls++, cancellation.Token));
        Assert.True(source.BytesRead > 0);
        Assert.Equal(0, calls);
        Assert.True(source.CanRead);
        Assert.Equal(before, TemporaryFiles());
    }

    /// <summary>
    /// 输入已经写入暂存文件后发生 IO 错误时，原始错误传播且文件被清理。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InputFailure_ShouldCleanupAndPreserveException(bool asynchronous)
    {
        var before = TemporaryFiles();
        using var provider = CreateProvider();
        var expected = new IOException("source read failed");
        using var source = new Input(CreateXlsx(), seekable: false, readFailure: expected);
        var calls = 0;
        var actual = await Record.ExceptionAsync(() => Import(
            provider.GetRequiredService<IExcelEntityBatchImporter>(), source, Options(1_000_000),
            asynchronous, _ => calls++));
        Assert.Same(expected, actual);
        Assert.True(source.BytesRead > 0);
        Assert.Equal(0, calls);
        Assert.True(source.CanRead);
        Assert.Equal(before, TemporaryFiles());
    }

    /// <summary>
    /// 零、负数不构成可用预算；既有构造函数保持不限额。
    /// </summary>
    [Fact]
    public void Options_ShouldRejectNonPositiveBudgetAndPreserveDefaults()
    {
        Assert.Null(new ExcelEntityBatchImportOptions().MaxTemporaryStorageBytes);
        foreach (var invalid in new long[] { 0, -1 })
            Assert.Equal("maxTemporaryStorageBytes", Assert.Throws<ArgumentOutOfRangeException>(() =>
                Options(invalid, 1)).ParamName);
    }

    protected void AssertResourceException(BingOfficesResourceLimitException exception)
    {
        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(ProviderName, exception.Provider);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Open, exception.Stage);
    }

    protected static ExcelEntityBatchImportOptions Options(long maximum, int mode = 1) =>
        new(1, null, mode == 2 ? ExcelEntityBatchParsingMode.Buffered : ExcelEntityBatchParsingMode.Forward,
            mode == 0 ? ExcelEntityBatchResultMode.Complete : ExcelEntityBatchResultMode.SummaryOnly, maximum);

    protected static ExcelEntityLayout<Root> Layout() => ExcelEntity.Layout<Root>(builder =>
        builder.ListRegion("Data", "A1", root => root.Rows));

    protected static async Task<ExcelEntityBatchImportSummary<Root>> Import(
        IExcelEntityBatchImporter importer, Stream source, ExcelEntityBatchImportOptions options,
        bool asynchronous, Action<ExcelEntityImportBatch<Root>> callback, CancellationToken token = default)
    {
        if (asynchronous)
            return await importer.ImportEntityBatchesAsync(source, Layout(), options, async (batch, _) =>
            {
                await Task.Yield();
                callback(batch);
            }, token);
        return importer.ImportEntityBatches(source, Layout(), options, callback, token);
    }

    protected static string[] TemporaryFiles() => Directory.GetFiles(Path.GetTempPath(), "bing-offices-forward-*")
        .OrderBy(path => path, StringComparer.Ordinal).ToArray();

    private static byte[] CreateXlsx()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            Write(archive, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
            Write(archive, "xl/sharedStrings.xml", "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + string.Concat(SharedValues.Select(value => "<si><t>" + value + "</t></si>")) + "</sst>");
            Write(archive, "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\" t=\"s\"><v>2</v></c><c r=\"B2\" t=\"s\"><v>3</v></c></row><row r=\"3\"><c r=\"A3\" t=\"s\"><v>4</v></c><c r=\"B3\" t=\"s\"><v>5</v></c></row></sheetData></worksheet>");
        }
        return stream.ToArray();
    }

    private static void Write(ZipArchive archive, string path, string text)
    {
        using var writer = new StreamWriter(archive.CreateEntry(path).Open(), new UTF8Encoding(false));
        writer.Write(text);
    }

    public sealed class Root { public List<Row> Rows { get; } = new(); }
    public sealed class Row { public string Code { get; set; } public string Note { get; set; } }

    protected sealed class Input : Stream
    {
        private readonly MemoryStream _inner;
        private readonly bool _seekable;
        private readonly CancellationTokenSource _cancelAfterRead;
        private readonly Exception _readFailure;
        internal Input(byte[] bytes, bool seekable, CancellationTokenSource cancelAfterRead = null,
            Exception readFailure = null)
        { _inner = new MemoryStream(bytes, false); _seekable = seekable; _cancelAfterRead = cancelAfterRead; _readFailure = readFailure; }
        internal int BytesRead { get; private set; }
        internal int AsyncReads { get; private set; }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _seekable;
        public override bool CanWrite => false;
        public override long Length => _seekable ? _inner.Length : throw new NotSupportedException();
        public override long Position { get => _seekable ? _inner.Position : throw new NotSupportedException(); set { if (!_seekable) throw new NotSupportedException(); _inner.Position = value; } }
        public override int Read(byte[] buffer, int offset, int count) { BeforeRead(); var read = _inner.Read(buffer, offset, count); AfterRead(read); return read; }
        public override int Read(Span<byte> buffer) { BeforeRead(); var read = _inner.Read(buffer); AfterRead(read); return read; }
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) { BeforeRead(); AsyncReads++; var read = await _inner.ReadAsync(buffer, cancellationToken); AfterRead(read); return read; }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        public override long Seek(long offset, SeekOrigin origin) => _seekable ? _inner.Seek(offset, origin) : throw new NotSupportedException();
        public override void Flush() => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
        private void AfterRead(int count) { BytesRead += count; if (count > 0) _cancelAfterRead?.Cancel(); }
        private void BeforeRead() { if (BytesRead > 0 && _readFailure != null) throw _readFailure; }
    }
}

/// <summary>
/// 临时文件全局快照用例不与其他文件资源测试并行。
/// </summary>
[CollectionDefinition("Entity temporary storage", DisableParallelization = true)]
public sealed class EntityTemporaryStorageCollection { }
