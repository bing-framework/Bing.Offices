using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>串行验证实际输入预算，避免多个 128 MiB 缓冲测试同时占用内存。</summary>
[CollectionDefinition("Excel default input budget", DisableParallelization = true)]
public sealed class DefaultInputLimitCollection { }

/// <summary>通过真实 Provider 复制路径验证默认、显式限额与显式关闭三种配置。</summary>
[Collection("Excel default input budget")]
public sealed class DefaultInputLimitTest
{
    private const long DefaultBudget = 128L * 1024 * 1024;

    public static IEnumerable<object[]> BudgetCases()
    {
        foreach (var asynchronous in new[] { false, true })
        foreach (var singleSheet in new[] { false, true })
        foreach (var scenario in new[] { "overflow", "exact", "disabled" })
            yield return new object[] { asynchronous, singleSheet, scenario };
    }

    /// <summary>实际默认阈值前允许复制、超出一字节拒绝，显式 null 则继续读取而不进入 DOM。</summary>
    [Theory]
    [MemberData(nameof(BudgetCases))]
    public async Task Import_DefaultBudget_ShouldEnforceBoundaryAndHonorExplicitNull(
        bool asynchronous, bool singleSheet, string scenario)
    {
        var length = DefaultBudget + (scenario == "exact" ? 0 : 1);
        using var source = new GeneratedInputStream(length, asynchronous);
        var limits = scenario == "disabled" ? new ExcelResourceLimits { MaxInputBytes = null } : null;
        if (scenario == "overflow")
        {
            var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(
                () => Import(source, asynchronous, singleSheet, limits));
            AssertLimitException(exception, DefaultBudget);
            Assert.False(source.TerminalReadAttempted);
        }
        else
        {
            // 到达哨兵证明限额未提前拒绝；哨兵在输入结束前抛出，避免解析巨大/无效 DOM。
            var exception = await Assert.ThrowsAsync<BingOfficesImportException>(
                () => Import(source, asynchronous, singleSheet, limits));
            Assert.Same(source.TerminalException, exception.InnerException);
            Assert.Equal(BingOfficesStage.Read, exception.Stage);
            Assert.True(source.TerminalReadAttempted);
        }
        Assert.Equal(length, source.BytesRead);
        Assert.True(source.CanRead);
        Assert.Equal(asynchronous, source.AsyncReads > 0);
        Assert.Equal(!asynchronous, source.SyncReads > 0);
    }

    public static IEnumerable<object[]> WorkbookCases()
    {
        foreach (var format in new[] { "xls", "xlsx" })
        foreach (var asynchronous in new[] { false, true })
        foreach (var singleSheet in new[] { false, true })
            yield return new object[] { format, asynchronous, singleSheet };
    }

    /// <summary>真实文件 IO 与两种格式验证成功、恰好上限及超限，不改变调用方文件句柄所有权。</summary>
    [Theory]
    [MemberData(nameof(WorkbookCases))]
    public async Task Import_RealFile_ShouldSupportDefaultExactDisabledAndRejectOverflow(
        string format, bool asynchronous, bool singleSheet)
    {
        using IWorkbook workbook = format == "xls" ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Count");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(7);
        using var serialized = new MemoryStream();
        workbook.Write(serialized, true);
        var bytes = serialized.ToArray();
        var path = Path.Combine(Path.GetTempPath(), $"bing-budget-{Guid.NewGuid():N}.{format}");
        await File.WriteAllBytesAsync(path, bytes);
        try
        {
            foreach (var limits in new[] { null, new ExcelResourceLimits { MaxInputBytes = bytes.Length },
                         new ExcelResourceLimits { MaxInputBytes = null } })
            {
                using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                    4096, asynchronous ? FileOptions.Asynchronous : FileOptions.None);
                var result = await Import(source, asynchronous, singleSheet, limits);
                Assert.Equal(new[] { 7 }, result.Items.Select(row => row.Count));
                Assert.Equal(new[] { 1 }, result.SourceRows);
                Assert.Empty(result.Errors);
                Assert.True(result.IsSuccess);
                Assert.True(result.IsComplete);
                Assert.Equal(ImportStopReason.None, result.StopReason);
                Assert.False(result.ErrorsTruncated);
                Assert.Null(result.MaxErrors);
                Assert.True(source.CanRead);
            }
            using var rejected = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, asynchronous ? FileOptions.Asynchronous : FileOptions.None);
            var error = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                Import(rejected, asynchronous, singleSheet,
                    new ExcelResourceLimits { MaxInputBytes = bytes.Length - 1 }));
            AssertLimitException(error, bytes.Length - 1);
            Assert.True(rejected.CanRead);
        }
        finally { File.Delete(path); }
    }

    private static void AssertLimitException(BingOfficesResourceLimitException exception, long limit)
    {
        Assert.Equal($"输入工作簿超过最大字节数: {limit}", exception.Message);
        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal("NPOI", exception.Provider);
        Assert.Equal(BingOfficesStage.Open, exception.Stage);
        Assert.Null(exception.InnerException);
    }

    private static async Task<ExcelSingleSheetImportResult<Row>> Import(Stream source, bool asynchronous,
        bool singleSheet, ExcelResourceLimits limits)
    {
        var importer = new NpoiExcelImporter();
        if (singleSheet)
            return asynchronous ? await importer.ImportSheetAsync<Row>(source, resourceLimits: limits)
                : importer.ImportSheet<Row>(source, resourceLimits: limits);
        var request = ExcelImport.Workbook<Root>(builder =>
        {
            builder.Sheet("Data", root => root.Rows);
            if (limits != null) builder.ResourceLimits(limits);
        });
        var result = asynchronous ? await importer.ImportAsync(source, request) : importer.Import(source, request);
        return new ExcelSingleSheetImportResult<Row>(result.Workbook.Rows, result.Errors,
            Assert.Single(result.Sheets).SourceRows, result.ErrorsTruncated, result.MaxErrors);
    }

    private sealed class Row { public int Count { get; set; } }
    private sealed class Root { public List<Row> Rows { get; } = new(); }

    /// <summary>不保存大型输入；按块供给真实字节，终点哨兵保证失败测试不进入 Office 解析器。</summary>
    private sealed class GeneratedInputStream : Stream
    {
        private readonly long _length;
        private readonly bool _asynchronous;
        private bool _disposed;
        public GeneratedInputStream(long length, bool asynchronous)
            => (_length, _asynchronous) = (length, asynchronous);
        public IOException TerminalException { get; } = new("budget-test input boundary");
        public long BytesRead { get; private set; }
        public int SyncReads { get; private set; }
        public int AsyncReads { get; private set; }
        public bool TerminalReadAttempted { get; private set; }
        public override bool CanRead => !_disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_asynchronous) throw new InvalidOperationException("异步输入不得回退到同步读取。");
            SyncReads++;
            return ReadBlock(buffer, offset, count);
        }
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            if (!_asynchronous) throw new InvalidOperationException("同步测试不应调用异步入口。");
            cancellationToken.ThrowIfCancellationRequested();
            await Task.Yield();
            AsyncReads++;
            return ReadBlock(buffer, offset, count);
        }
        private int ReadBlock(byte[] buffer, int offset, int count)
        {
            if (BytesRead == _length)
            {
                TerminalReadAttempted = true;
                throw TerminalException;
            }
            var read = (int)Math.Min(count, _length - BytesRead);
            Array.Clear(buffer, offset, read);
            BytesRead += read;
            return read;
        }
        protected override void Dispose(bool disposing) { _disposed = true; base.Dispose(disposing); }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
