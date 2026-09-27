using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 失败工作簿公开提交扩展点的职责级回归测试。
/// </summary>
public sealed class NpoiFailureWorkbookCommitterTest
{
    /// <summary>
    /// 枚举 XLS、XLSX 与同步、异步导入组合。
    /// </summary>
    /// <returns>测试参数集合。</returns>
    public static IEnumerable<object[]> FormatsAndCalls()
    {
        foreach (var xls in new[] { false, true })
        foreach (var asynchronous in new[] { false, true })
            yield return new object[] { xls, asynchronous };
    }

    /// <summary>
    /// 验证注入的提交器仅提交一次完整失败工作簿。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task InjectedCommitter_ShouldCommitCompleteWorkbookExactlyOnce(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);
        var path = Path.Combine(directory.Path, xls ? "failure.xls" : "failure.xlsx");
        File.WriteAllBytes(path, new byte[] { 7, 8, 9 });

        var result = await Import(importer, source, Request(directory.Path, path), asynchronous, cancellation.Token);

        AssertCalls(committer, asynchronous);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.Equal(cancellation.Token, committer.Token);
        AssertWorkbook(path, result);
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证空提交器和旧构造函数均回退到默认原子提交实现。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task NullCommitter_ShouldUseDefaultWithoutChangingOldConstructor(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, xls ? "failure.xls" : "failure.xlsx");
        foreach (var importer in new[] { new NpoiExcelImporter(),
                     new NpoiExcelImporter(null, null, null, null, null, (IFileExportCommitter)null) })
        {
            using var source = Source(xls);
            var result = await Import(importer, source, Request(directory.Path, path), asynchronous);
            AssertWorkbook(path, result);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        }
    }

    /// <summary>
    /// 验证宿主注册的提交器由 NPOI 导入器使用且不会被默认注册覆盖。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Registration_ShouldPreserveHostCommitter(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(false);
        var committer = new RecordingCommitter();
        var services = new ServiceCollection();
        services.AddSingleton<IFileExportCommitter>(committer);
        services.AddBingOfficesNpoi().AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        var path = Path.Combine(directory.Path, "failure.xlsx");

        var result = await Import(provider.GetRequiredService<IExcelImporter>(), source,
            Request(directory.Path, path), asynchronous);

        Assert.Same(committer, provider.GetRequiredService<IFileExportCommitter>());
        Assert.Single(services.Where(item => item.ServiceType == typeof(IFileExportCommitter)));
        AssertCalls(committer, asynchronous);
        AssertWorkbook(path, result);
    }

    /// <summary>
    /// 验证无文件输出或无失败时不会调用提交器。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="scenario">测试场景名称。</param>
    [Theory]
    [InlineData(false, "valid")]
    [InlineData(true, "valid")]
    [InlineData(false, "disabled")]
    [InlineData(true, "disabled")]
    [InlineData(false, "stream")]
    [InlineData(true, "stream")]
    public async Task NonFileOrNoFailure_ShouldNotCallCommitter(bool asynchronous, string scenario)
    {
        using var directory = new TestDirectory();
        using var source = Source(false, scenario == "valid");
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);
        var options = new ExcelImportFailureOptions
        {
            Mode = scenario == "disabled" ? ExcelImportFailureWorkbookMode.None : ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = scenario == "stream" ? output : null,
            DestinationPath = scenario == "stream" ? null : Path.Combine(directory.Path, "unused.xlsx"),
            TemporaryDirectory = directory.Path
        };

        var result = await Import(importer, source, Request(options), asynchronous);

        Assert.Equal(scenario == "valid", result.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(Directory.GetFiles(directory.Path));
        Assert.True(output.CanWrite);
        if (scenario == "stream")
        {
            output.Position = 0;
            using var workbook = WorkbookFactory.Create(output);
            AssertWorkbook(workbook, result);
        }
        else
            Assert.Equal(0, output.Length);
    }

    /// <summary>
    /// 验证提交失败时保留原异常、目标文件和单次异常观察。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task CommitFailure_ShouldPreserveExceptionAndObserveOnce(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"), "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new NpoiExcelImporter(null, null, null, null, new[] { observer }, committer);

        var actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            Import(importer, source, Request(directory.Path, path), asynchronous));

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(BingOfficesStage.Commit, actual.Stage);
        Assert.Equal(BingOfficesErrorCode.FileCommitFailed, actual.Code);
        AssertCalls(committer, asynchronous);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证写入期间取消时保留目标文件并清理临时文件。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task CancellationDuringWrite_ShouldPreserveTargetAndCleanup(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 1, 3, 5 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var observer = new RecordingObserver();
        var importer = new NpoiExcelImporter(null, null, null, null, new[] { observer }, committer);

        var exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Import(importer, source, Request(directory.Path, path), asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        AssertCalls(committer, asynchronous);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证写入失败时保留目标文件并只观察一次异常。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task WriteFailure_ShouldPreserveTargetAndObserveOnce(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 5, 4, 3 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesImportException("write failure", new IOException("write"), "NPOI", BingOfficesStage.Write);
        var committer = new RecordingCommitter { AfterWrite = () => throw expected };
        var observer = new RecordingObserver();
        var importer = new NpoiExcelImporter(null, null, null, null, new[] { observer }, committer);

        var actual = await Assert.ThrowsAsync<BingOfficesImportException>(() =>
            Import(importer, source, Request(directory.Path, path), asynchronous));

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(BingOfficesStage.Write, actual.Stage);
        Assert.Equal(BingOfficesOperation.Import, actual.Operation);
        AssertCalls(committer, asynchronous);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证预取消时不会进入文件提交扩展点。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task PreCanceled_ShouldNotCallCommitter(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);

        await Assert.ThrowsAsync<OperationCanceledException>(() => Import(importer, source,
            Request(directory.Path, Path.Combine(directory.Path, "unused.xlsx")), asynchronous, cancellation.Token));

        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 断言同步和异步提交调用次数与入口一致。
    /// </summary>
    /// <param name="committer">记录调用信息的提交器。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    private static void AssertCalls(RecordingCommitter committer, bool asynchronous)
    {
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
    }

    /// <summary>
    /// 创建用于导入测试的 XLS 或 XLSX 工作簿流。
    /// </summary>
    /// <param name="xls">是否创建 XLS 工作簿。</param>
    /// <param name="valid">是否写入有效编码。</param>
    /// <returns>定位到起始位置的工作簿流。</returns>
    private static MemoryStream Source(bool xls, bool valid = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(valid ? "valid" : "");
        // 第二个有效列确保空 Code 所在的数据行不会被视为空白行跳过。
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.GetRow(1).CreateCell(1).SetCellValue(1);
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建指定文件路径的失败工作簿请求。
    /// </summary>
    /// <param name="directory">临时文件目录。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <returns>配置完成的工作簿导入请求。</returns>
    private static ExcelWorkbookImportRequest<Root> Request(string directory, string path) => Request(
        new ExcelImportFailureOptions { Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            DestinationPath = path, TemporaryDirectory = directory });

    /// <summary>
    /// 根据失败工作簿选项创建导入请求。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <returns>配置完成的工作簿导入请求。</returns>
    private static ExcelWorkbookImportRequest<Root> Request(ExcelImportFailureOptions options) =>
        ExcelImport.Workbook<Root>(builder => builder.FailureWorkbook(options)
            .Sheet<Row>("Data", root => root.Rows, sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 按指定模式执行工作簿导入。
    /// </summary>
    /// <param name="importer">执行导入的导入器。</param>
    /// <param name="source">来源工作簿流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    /// <param name="token">导入取消令牌。</param>
    /// <returns>工作簿导入结果。</returns>
    private static async Task<ExcelWorkbookImportResult<Root>> Import(IExcelImporter importer, Stream source,
        ExcelWorkbookImportRequest<Root> request, bool asynchronous, CancellationToken token = default) =>
        asynchronous ? await importer.ImportAsync(source, request, token) : importer.Import(source, request, token);

    /// <summary>
    /// 从文件读取并断言失败工作簿内容。
    /// </summary>
    /// <param name="path">失败工作簿路径。</param>
    /// <param name="result">导入结果。</param>
    private static void AssertWorkbook(string path, ExcelWorkbookImportResult<Root> result)
    {
        using var stream = File.OpenRead(path);
        using var workbook = WorkbookFactory.Create(stream);
        AssertWorkbook(workbook, result);
    }

    /// <summary>
    /// 断言失败工作簿结构、数据和错误摘要。
    /// </summary>
    /// <param name="workbook">待断言的工作簿。</param>
    /// <param name="result">导入结果。</param>
    private static void AssertWorkbook(IWorkbook workbook, ExcelWorkbookImportResult<Root> result)
    {
        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(new[] { "Data", "_ImportErrors" }, Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName));
        var sheet = workbook.GetSheetAt(0);
        Assert.Equal(2, sheet.PhysicalNumberOfRows);
        Assert.Equal(new[] { "Code", "Quantity" }, sheet.GetRow(0).Cells.Select(cell => cell.StringCellValue));
        Assert.Equal("", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(1, sheet.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Equal(error.Message, sheet.GetRow(1).GetCell(0).CellComment.String.String);
        var summary = workbook.GetSheetAt(1);
        Assert.Equal(2, summary.PhysicalNumberOfRows);
        Assert.Equal(new[] { "Code", "Message", "Sheet", "Row", "Column", "Property", "Header", "RawValue" },
            summary.GetRow(0).Cells.Select(cell => cell.StringCellValue));
        Assert.Equal(new[] { error.Code.ToString(), error.Message, "Data", "2", "1", "Code", "Code", "" },
            summary.GetRow(1).Cells.Select(cell => cell.ToString()));
    }

    private sealed class Row
    {
        /// <summary>
        /// 获取或设置必填编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 失败工作簿导入的根模型。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取或设置数据行集合。
        /// </summary>
        public List<Row> Rows { get; set; } = new();
    }

    /// <summary>
    /// 记录导入异常观察调用。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 获取已观察的异常集合。
        /// </summary>
        public List<BingOfficesException> Exceptions { get; } = new();

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }

    /// <summary>
    /// 记录提交调用并委托真实默认提交器。
    /// </summary>
    /// <remarks>
    /// 支持在写入结束、原子替换前注入取消或异常。
    /// </remarks>
    private sealed class RecordingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 默认文件提交实现。
        /// </summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>
        /// 获取同步提交调用次数。
        /// </summary>
        public int SyncCalls { get; private set; }

        /// <summary>
        /// 获取异步提交调用次数。
        /// </summary>
        public int AsyncCalls { get; private set; }

        /// <summary>
        /// 获取最近一次提交路径。
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// 获取最近一次提交格式。
        /// </summary>
        public string Format { get; private set; }

        /// <summary>
        /// 获取最近一次提交使用的取消令牌。
        /// </summary>
        public CancellationToken Token { get; private set; }

        /// <summary>
        /// 获取或设置提交前抛出的异常。
        /// </summary>
        public Exception Failure { get; set; }

        /// <summary>
        /// 获取或设置写入完成后的回调。
        /// </summary>
        public Action AfterWrite { get; set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        {
            SyncCalls++;
            Record(path, cancellationToken, format);
            _inner.Commit(path, stream => { write(stream); AfterWrite?.Invoke(); }, cancellationToken, format);
        }

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            AsyncCalls++;
            Record(path, cancellationToken, format);
            return _inner.CommitAsync(path, async (stream, token) =>
            {
                await writeAsync(stream, token);
                AfterWrite?.Invoke();
            }, cancellationToken, format);
        }

        /// <summary>
        /// 记录提交参数并注入预设异常。
        /// </summary>
        /// <param name="path">提交目标路径。</param>
        /// <param name="token">提交取消令牌。</param>
        /// <param name="format">提交内容格式。</param>
        private void Record(string path, CancellationToken token, string format)
        {
            Path = path;
            Token = token;
            Format = format;
            if (Failure != null)
                throw Failure;
        }
    }

    /// <summary>
    /// 为单个测试创建并清理临时目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取测试临时目录路径。
        /// </summary>
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "npoi-committer-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, true);
    }
}
