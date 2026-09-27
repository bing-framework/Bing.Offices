using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Microsoft.Extensions.DependencyInjection;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML 失败工作簿使用公开文件提交扩展点。
/// </summary>
public sealed class ClosedXmlFailureWorkbookCommitterTest
{
    /// <summary>
    /// 验证构造注入的提交器在同步和异步路径各提交一次完整工作簿。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InjectedCommitter_ShouldCommitCompleteWorkbookExactlyOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source();
        using var cancellation = new CancellationTokenSource();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlExcelImporter(null, null, null, null, null, committer);
        var path = Path.Combine(directory.Path, "failure.xlsx");
        File.WriteAllBytes(path, new byte[] { 7, 8, 9 });

        var result = await Import(importer, source, Request(directory.Path, path), asynchronous,
            cancellation.Token);

        AssertCalls(committer, asynchronous);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.Equal(cancellation.Token, committer.Token);
        Assert.NotNull(committer.Bytes);
        AssertWorkbook(committer.Bytes, result);
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证空提交器和旧构造函数均回退到默认原子提交实现。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NullCommitter_ShouldUseDefaultWithoutChangingOldConstructor(bool asynchronous)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        foreach (var importer in new[]
        {
            new ClosedXmlExcelImporter(),
            new ClosedXmlExcelImporter(null, null, null, null, null, (IFileExportCommitter)null)
        })
        {
            using var source = Source();
            var result = await Import(importer, source, Request(directory.Path, path), asynchronous);

            AssertWorkbook(path, result);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
            File.Delete(path);
        }
    }

    /// <summary>
    /// 验证宿主注册的提交器由 ClosedXML 导入器使用，且不会被默认注册覆盖。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Registration_ShouldPreserveHostCommitter(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source();
        var committer = new RecordingCommitter();
        var services = new ServiceCollection();
        services.AddSingleton<IFileExportCommitter>(committer);
        services.AddBingOfficesClosedXml().AddBingOfficesClosedXml();
        using var provider = services.BuildServiceProvider();
        var path = Path.Combine(directory.Path, "failure.xlsx");

        var importer = provider.GetRequiredService<IExcelImporter>();
        var result = await Import(importer, source, Request(directory.Path, path), asynchronous);

        Assert.Same(committer, provider.GetRequiredService<IFileExportCommitter>());
        Assert.Single(services.Where(item => item.ServiceType == typeof(IFileExportCommitter)));
        AssertCalls(committer, asynchronous);
        Assert.Equal(path, committer.Path);
        AssertWorkbook(committer.Bytes, result);
    }

    /// <summary>
    /// 验证没有失败工作簿文件路径时不会调用提交器，调用方流仍由调用方拥有。
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
        using var source = Source(valid: scenario == "valid");
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlExcelImporter(null, null, null, null, null, committer);
        var options = new ExcelImportFailureOptions
        {
            Mode = scenario == "disabled"
                ? ExcelImportFailureWorkbookMode.None
                : ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = scenario == "stream" ? output : null,
            DestinationPath = scenario == "stream"
                ? null
                : Path.Combine(directory.Path, "unused.xlsx"),
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
            using var workbook = new XLWorkbook(output);
            AssertWorkbook(workbook, result);
        }
        else
            Assert.Equal(0, output.Length);
    }

    /// <summary>
    /// 验证提交失败保留原异常、仅观察一次且不破坏原目标文件。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CommitFailure_ShouldPreserveExceptionAndObserveOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new ClosedXmlExcelImporter(null, null, null, null, new[] { observer }, committer);

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
    /// 验证失败工作簿写入回调异常由默认提交器清理临时文件、保留原目标并只观察一次。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WriteFailure_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 4, 6, 8 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesImportException("write failed", provider: "ClosedXML",
            stage: BingOfficesStage.Write);
        var committer = new RecordingCommitter { AfterWrite = () => throw expected };
        var observer = new RecordingObserver();
        var importer = new ClosedXmlExcelImporter(null, null, null, null, new[] { observer }, committer);

        var actual = await Assert.ThrowsAsync<BingOfficesImportException>(() =>
            Import(importer, source, Request(directory.Path, path), asynchronous));

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(BingOfficesOperation.Import, actual.Operation);
        Assert.Equal(BingOfficesStage.Write, actual.Stage);
        AssertCalls(committer, asynchronous);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证写入完成后取消不会替换目标文件并且清理临时文件。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDuringWrite_ShouldPreserveTargetAndCleanup(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source();
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 1, 3, 5 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var observer = new RecordingObserver();
        var importer = new ClosedXmlExcelImporter(null, null, null, null, new[] { observer }, committer);

        var exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Import(importer, source, Request(directory.Path, path), asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        AssertCalls(committer, asynchronous);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证预取消在进入文件提交扩展点前即被拒绝。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreCanceled_ShouldNotCallCommitter(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlExcelImporter(null, null, null, null, null, committer);

        await Assert.ThrowsAsync<OperationCanceledException>(() => Import(importer, source,
            Request(directory.Path, Path.Combine(directory.Path, "unused.xlsx")), asynchronous,
            cancellation.Token));

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
    /// 创建用于导入测试的 XLSX 工作簿流。
    /// </summary>
    /// <param name="valid">是否写入有效编码。</param>
    /// <returns>定位到起始位置的工作簿流。</returns>
    private static MemoryStream Source(bool valid = false)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = "Code";
            sheet.Cell("A2").Value = valid ? "valid" : string.Empty;
            sheet.Cell("B1").Value = "Quantity";
            sheet.Cell("B2").Value = 1;
            workbook.SaveAs(stream);
        }
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
        new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            DestinationPath = path,
            TemporaryDirectory = directory
        });

    /// <summary>
    /// 根据失败工作簿选项创建导入请求。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <returns>配置完成的工作簿导入请求。</returns>
    private static ExcelWorkbookImportRequest<Root> Request(ExcelImportFailureOptions options) =>
        ExcelImport.Workbook<Root>(builder => builder.FailureWorkbook(options)
            .Sheet<Row>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 按指定模式执行工作簿导入。
    /// </summary>
    /// <param name="importer">执行导入的导入器。</param>
    /// <param name="source">来源工作簿流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    /// <param name="token">导入取消令牌。</param>
    /// <returns>工作簿导入结果。</returns>
    private static async Task<ExcelWorkbookImportResult<Root>> Import(IExcelImporter importer,
        Stream source, ExcelWorkbookImportRequest<Root> request, bool asynchronous,
        CancellationToken token = default) => asynchronous
        ? await importer.ImportAsync(source, request, token)
        : importer.Import(source, request, token);

    /// <summary>
    /// 从内存字节读取并断言失败工作簿内容。
    /// </summary>
    /// <param name="bytes">失败工作簿字节。</param>
    /// <param name="result">导入结果。</param>
    private static void AssertWorkbook(byte[] bytes, ExcelWorkbookImportResult<Root> result)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        AssertWorkbook(workbook, result);
    }

    /// <summary>
    /// 从文件读取并断言失败工作簿内容。
    /// </summary>
    /// <param name="path">失败工作簿路径。</param>
    /// <param name="result">导入结果。</param>
    private static void AssertWorkbook(string path, ExcelWorkbookImportResult<Root> result)
    {
        using var workbook = new XLWorkbook(path);
        AssertWorkbook(workbook, result);
    }

    /// <summary>
    /// 断言失败工作簿结构、数据和错误摘要。
    /// </summary>
    /// <param name="workbook">待断言的工作簿。</param>
    /// <param name="result">导入结果。</param>
    private static void AssertWorkbook(XLWorkbook workbook, ExcelWorkbookImportResult<Root> result)
    {
        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(new[] { "Data", "_ImportErrors" },
            workbook.Worksheets.Select(item => item.Name));
        Assert.True(workbook.TryGetWorksheet("Data", out var sheet));
        Assert.True(workbook.TryGetWorksheet("_ImportErrors", out var summary));
        Assert.Equal(2, sheet.LastRowUsed().RowNumber());
        Assert.Equal(2, sheet.LastColumnUsed().ColumnNumber());
        Assert.Equal("Code", sheet.Cell("A1").GetString());
        Assert.Equal("Quantity", sheet.Cell("B1").GetString());
        Assert.Equal(string.Empty, sheet.Cell("A2").GetString());
        Assert.Equal(1, sheet.Cell("B2").GetValue<int>());
        Assert.True(sheet.Cell("A2").HasComment);
        Assert.Equal("Bing.Offices", sheet.Cell("A2").GetComment().Author);
        Assert.Equal(error.Message, sheet.Cell("A2").GetComment().ToString());
        Assert.Equal(2, summary.LastRowUsed().RowNumber());
        Assert.Equal(8, summary.LastColumnUsed().ColumnNumber());
        Assert.Equal(new[] { "Code", "Message", "Sheet", "Row", "Column", "Property", "Header", "RawValue" },
            Enumerable.Range(1, 8).Select(column => summary.Cell(1, column).GetString()));
        Assert.Equal(new[]
        {
            error.Code.ToString(), error.Message, "Data", "2", "1", nameof(Row.Code), string.Empty, string.Empty
        }, Enumerable.Range(1, 8).Select(column => summary.Cell(2, column).GetString()));
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

    private sealed class Root
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

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
    /// 委托真实默认提交器，并可在写入结束、原子替换前注入取消。
    /// </summary>
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
        /// 获取最近一次提交生成的文件内容。
        /// </summary>
        public byte[] Bytes { get; private set; }

        /// <summary>
        /// 获取或设置提交前抛出的异常。
        /// </summary>
        public Exception Failure { get; set; }

        /// <summary>
        /// 获取或设置写入完成后的回调。
        /// </summary>
        public Action AfterWrite { get; set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format)
        {
            SyncCalls++;
            Record(path, cancellationToken, format);
            _inner.Commit(path, stream =>
            {
                write(stream);
                AfterWrite?.Invoke();
            }, cancellationToken, format);
            Bytes = File.ReadAllBytes(path);
        }

        /// <inheritdoc />
        public async Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            AsyncCalls++;
            Record(path, cancellationToken, format);
            await _inner.CommitAsync(path, async (stream, token) =>
            {
                await writeAsync(stream, token);
                AfterWrite?.Invoke();
            }, cancellationToken, format);
            Bytes = File.ReadAllBytes(path);
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
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "closedxml-committer-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, true);
    }
}
