using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Data;
using Bing.Offices.Testing.Models;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 通过公共导入入口验证失败工作簿提交与流复制的失败边界。
/// </summary>
public sealed class FailureWorkbookFailureBoundaryContractTest
{
    /// <summary>
    /// 获取受支持格式、调用方式和失败工作簿模式的组合。
    /// </summary>
    public static IEnumerable<object[]> Cases
    {
        get
        {
            foreach (var provider in new[] { "NPOI", "ClosedXML" })
            foreach (var format in new[] { ExcelFormat.Xls, ExcelFormat.Xlsx })
            foreach (var asynchronous in new[] { false, true })
            foreach (var mode in new[] { ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                         ExcelImportFailureWorkbookMode.ErrorRowsOnly })
                if (provider == "NPOI" || format == ExcelFormat.Xlsx)
                    yield return new object[] { provider, format, asynchronous, mode };
        }
    }

    /// <summary>
    /// 两种模式均能创建或替换目标，输出可重新打开且保留错误位置和摘要。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="format">测试工作簿格式。</param>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    /// <param name="mode">失败工作簿生成模式。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task PathSuccess_ShouldCreateAndReplaceCompleteWorkbook(string provider, ExcelFormat format,
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, "failure" + (format == ExcelFormat.Xls ? ".xls" : ".xlsx"));
        using var source = await CreateSource(provider, format);
        var request = Request(new ExcelImportFailureOptions { Mode = mode, DestinationPath = path });
        foreach (var existing in new[] { false, true })
        {
            if (existing)
                File.WriteAllBytes(path, new byte[] { 7, 8, 9 });
            source.Position = 0;
            var result = await Import(provider, source, request, asynchronous);
            Assert.False(result.IsSuccess);
            Assert.Empty(result.Workbook.Rows);
            using (var output = File.OpenRead(path))
            using (var workbook = WorkbookFactory.Create(output))
            {
                Assert.Equal(new[] { "Data", "_ImportErrors" },
                    Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName));
                var sheet = workbook.GetSheet("Data");
                Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
                Assert.Equal("Quantity", sheet.GetRow(0).GetCell(1).StringCellValue);
                Assert.Equal(1, sheet.LastRowNum);
                Assert.Equal(0, sheet.GetRow(1).GetCell(1).NumericCellValue);
                if (mode == ExcelImportFailureWorkbookMode.AnnotatedOriginal)
                    Assert.NotNull(sheet.GetRow(1).GetCell(0).CellComment);
                else
                {
                    Assert.Equal("Data", sheet.GetRow(1).GetCell(2).StringCellValue);
                    Assert.Equal(2, sheet.GetRow(1).GetCell(3).NumericCellValue);
                }
                var summary = workbook.GetSheet("_ImportErrors");
                Assert.Equal(result.Errors.Count, summary.LastRowNum);
                Assert.Equal("Data", summary.GetRow(1).GetCell(2).StringCellValue);
                Assert.Equal(2, summary.GetRow(1).GetCell(3).NumericCellValue);
            }
            Assert.True(source.CanRead);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        }
    }

    /// <summary>
    /// 真实文件系统提交失败不得破坏旧目标，且同目录提交临时文件必须清理。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="format">测试工作簿格式。</param>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    /// <param name="mode">失败工作簿生成模式。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task CommitFailure_ShouldPreserveTargetAndCleanTemporaryFiles(string provider, ExcelFormat format,
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, "target.xlsx");
        var sentinel = Encoding.UTF8.GetBytes("existing failure output");
        // Windows 使用拒绝删除共享的真实旧文件；其他系统用目录目标稳定拒绝文件提交。
        var sentinelPath = OperatingSystem.IsWindows() ? path : Path.Combine(path, "sentinel");
        if (!OperatingSystem.IsWindows())
            Directory.CreateDirectory(path);
        File.WriteAllBytes(sentinelPath, sentinel);
        using var locked = File.Open(sentinelPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var source = await CreateSource(provider, format);
        var request = Request(new ExcelImportFailureOptions
        {
            Mode = mode, DestinationPath = path, TemporaryDirectory = directory.Path
        });

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            Import(provider, source, request, asynchronous));

        Assert.Equal(BingOfficesStage.Commit, exception.Stage);
        Assert.NotNull(exception.InnerException);
        Assert.Equal(sentinel, File.ReadAllBytes(sentinelPath));
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 序列化预算失败发生在目标提交之前，必须保留旧文件与调用方输入。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="format">测试工作簿格式。</param>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    /// <param name="mode">失败工作簿生成模式。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task SerializationFailure_ShouldPreserveTargetAndCleanTemporaryFiles(string provider,
        ExcelFormat format, bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, "target.xlsx");
        var sentinel = new byte[] { 1, 2, 3, 4 };
        File.WriteAllBytes(path, sentinel);
        using var source = await CreateSource(provider, format);
        var request = Request(new ExcelImportFailureOptions
        {
            Mode = mode, DestinationPath = path, TemporaryDirectory = directory.Path, MaxSerializedBytes = 1
        });

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            Import(provider, source, request, asynchronous));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(provider, exception.Provider);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 通用目标流部分写入后失败属于 best-effort，仍须传播失败并保留调用方所有权。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="format">测试工作簿格式。</param>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    /// <param name="mode">失败工作簿生成模式。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task StreamPartialWriteFailure_ShouldPreserveOwnership(string provider, ExcelFormat format,
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = await CreateSource(provider, format);
        using var destination = new PartialFailureStream(asynchronous);
        var request = Request(new ExcelImportFailureOptions
        {
            Mode = mode, Destination = destination, TemporaryDirectory = directory.Path
        });

        var exception = await Assert.ThrowsAnyAsync<Exception>(() =>
            Import(provider, source, request, asynchronous));

        Assert.True(ContainsCause(exception, destination.Failure), exception.ToString());
        Assert.Equal(4, destination.Length);
        Assert.Equal(asynchronous, destination.AsyncWriteObserved);
        Assert.False(destination.Disposed);
        Assert.True(destination.CanWrite);
        Assert.True(source.CanRead);
        Assert.Empty(Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 判断异常链是否包含指定异常实例。
    /// </summary>
    /// <param name="exception">待遍历的异常链起点。</param>
    /// <param name="cause">要查找的原始异常实例。</param>
    /// <returns>异常链包含指定实例时返回 true；否则返回 false。</returns>
    private static bool ContainsCause(Exception exception, Exception cause) =>
        ReferenceEquals(exception, cause) || (exception.InnerException != null && ContainsCause(exception.InnerException, cause));

    /// <summary>
    /// 异步创建包含校验失败数据的源工作簿。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="format">测试工作簿格式。</param>
    /// <returns>包含测试工作簿内容且位于起始位置的内存流。</returns>
    private static async Task<MemoryStream> CreateSource(string provider, ExcelFormat format)
    {
        var source = new MemoryStream();
        await ProviderDrivers.Get(provider).CreateExporter().ExportAsync(ExcelExport.Workbook(builder =>
            builder.Format(format).AddSheet("Data", ContractData.InvalidValidationRows())), source);
        source.Position = 0;
        return source;
    }

    /// <summary>
    /// 创建生成失败工作簿的导入请求。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <returns>启用失败工作簿并继续收集校验错误的导入请求。</returns>
    private static ExcelWorkbookImportRequest<ValidationContractWorkbook> Request(ExcelImportFailureOptions options) =>
        ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder.FailureWorkbook(options)
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 通过指定调用方式执行工作簿导入。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="source">包含测试工作簿的输入流。</param>
    /// <param name="request">工作簿操作请求。</param>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    /// <returns>指定调用方式完成后的工作簿导入结果。</returns>
    private static async Task<ExcelWorkbookImportResult<ValidationContractWorkbook>> Import(string provider,
        Stream source, ExcelWorkbookImportRequest<ValidationContractWorkbook> request, bool asynchronous)
    {
        var importer = ProviderDrivers.Get(provider).CreateImporter();
        return asynchronous ? await importer.ImportAsync(source, request) : importer.Import(source, request);
    }

    /// <summary>
    /// 负责创建和清理独立临时目录的测试资源。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取当前测试独占的临时目录路径。
        /// </summary>
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "bing-contract-failure-" + Guid.NewGuid().ToString("N"));
        /// <summary>
        /// 初始化一个 <see cref="TestDirectory"/> 类型的实例。
        /// </summary>
        internal TestDirectory() => Directory.CreateDirectory(Path);
        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }

    /// <summary>
    /// 注入部分写入失败的测试内存流。
    /// </summary>
    /// <remarks>写入最多四个字节后抛出固定异常；可配置为拒绝同步写入。</remarks>
    private sealed class PartialFailureStream : MemoryStream
    {
        /// <summary>
        /// 标记是否禁止同步写入，以检测异步出口的同步回退。
        /// </summary>
        private readonly bool _requireAsync;
        /// <summary>
        /// 初始化一个 <see cref="PartialFailureStream"/> 类型的实例。
        /// </summary>
        /// <param name="requireAsync">是否禁止通过同步入口写入。</param>
        internal PartialFailureStream(bool requireAsync) => _requireAsync = requireAsync;
        /// <summary>
        /// 获取部分写入后抛出的固定异常实例。
        /// </summary>
        internal IOException Failure { get; } = new("injected partial destination write failure");
        /// <summary>
        /// 获取是否调用过异步写入入口。
        /// </summary>
        internal bool AsyncWriteObserved { get; private set; }
        /// <summary>
        /// 获取目标流是否已被释放。
        /// </summary>
        internal bool Disposed { get; private set; }
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_requireAsync)
                throw new InvalidOperationException("Synchronous destination write is forbidden.");
            Fail(buffer, offset, count);
        }
        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            AsyncWriteObserved = true;
            try { Fail(buffer, offset, count); return Task.CompletedTask; }
            catch (Exception exception) { return Task.FromException(exception); }
        }
        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default) =>
            new(WriteAsync(buffer.ToArray(), 0, buffer.Length, token));
        /// <summary>
        /// 写入最多四个字节后抛出固定异常。
        /// </summary>
        /// <param name="buffer">待写入的字节缓冲区。</param>
        /// <param name="offset">写入数据在缓冲区中的起始位置。</param>
        /// <param name="count">请求写入的字节数。</param>
        private void Fail(byte[] buffer, int offset, int count)
        {
            base.Write(buffer, offset, Math.Min(count, 4));
            throw Failure;
        }
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }
}
