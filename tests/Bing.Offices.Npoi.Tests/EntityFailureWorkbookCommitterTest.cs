using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Entities;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 实体失败工作簿的文件提交扩展点。
/// </summary>
public sealed class EntityFailureWorkbookCommitterTest
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
    /// 枚举 XLS、XLSX、同步、异步与失败工作簿模式组合。
    /// </summary>
    /// <returns>测试参数集合。</returns>
    public static IEnumerable<object[]> FormatsCallsAndModes()
    {
        foreach (var xls in new[] { false, true })
        foreach (var asynchronous in new[] { false, true })
        foreach (var mode in new[]
                 {
                     ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                     ExcelImportFailureWorkbookMode.AnnotatedOriginal
                 })
            yield return new object[] { xls, asynchronous, mode };
    }

    /// <summary>
    /// 验证实体失败工作簿通过注入提交器只提交一次。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task InjectedCommitter_ShouldCommitEntityFailureWorkbookExactlyOnce(bool xls,
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);
        var path = Path.Combine(directory.Path, xls ? "entity-failure.xls" : "entity-failure.xlsx");
        File.WriteAllBytes(path, new byte[] { 7, 8, 9 });

        var result = await Import(importer, source, Options(directory.Path, path), asynchronous,
            cancellation.Token);

        Assert.False(result.IsSuccess);
        AssertCalls(committer, asynchronous);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.Equal(cancellation.Token, committer.Token);
        AssertEntityFailureWorkbook(path, xls, result);
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体失败工作簿提交失败时保留目标文件并只观察一次异常。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task CommitFailure_ShouldPreserveEntityTargetAndObserveOnce(bool xls,
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        var path = Path.Combine(directory.Path, xls ? "entity-failure.xls" : "entity-failure.xlsx");
        var sentinel = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new NpoiExcelImporter(null, null, null, null, new[] { observer }, committer);

        var actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            Import(importer, source, Options(directory.Path, path), asynchronous));

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(BingOfficesErrorCode.FileCommitFailed, actual.Code);
        Assert.Equal(BingOfficesStage.Commit, actual.Stage);
        AssertCalls(committer, asynchronous);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证写入完成后取消不会替换实体失败目标并且会清理暂存文件。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task CancellationDuringWrite_ShouldPreserveEntityTargetAndCleanup(bool xls,
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, xls ? "entity-failure.xls" : "entity-failure.xlsx");
        var sentinel = new byte[] { 1, 3, 5 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);

        var exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Import(importer, source, Options(directory.Path, path), asynchronous,
                cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        AssertCalls(committer, asynchronous);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证预取消在实体失败工作簿进入提交扩展点前即被拒绝。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task PreCanceled_ShouldNotCallEntityCommitter(bool xls, bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);
        var path = Path.Combine(directory.Path, xls ? "unused.xls" : "unused.xlsx");

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Import(importer, source, Options(directory.Path, path), asynchronous,
                cancellation.Token));

        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体失败工作簿直接输出到调用方流时不调用文件提交器。
    /// </summary>
    [Fact]
    public async Task CallerStream_ShouldBypassEntityCommitter()
    {
        using var source = Source(false);
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = output
            });

        var result = importer.ImportEntity(source, CreateLayout(), options);

        Assert.False(result.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.True(output.Length > 0);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证没有实体错误时不会创建失败工作簿或调用提交器。
    /// </summary>
    [Fact]
    public void SuccessfulEntityImport_ShouldNotCallCommitter()
    {
        using var directory = new TestDirectory();
        using var source = Source(false, valid: true);
        var committer = new RecordingCommitter();
        var importer = new NpoiExcelImporter(null, null, null, null, null, committer);
        var path = Path.Combine(directory.Path, "unused.xlsx");

        var result = importer.ImportEntity(source, CreateLayout(), Options(directory.Path, path));

        Assert.True(result.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证 XLSX Forward 实体失败工作簿通过提交器使用对应入口。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [InlineData(false, ExcelImportFailureWorkbookMode.ErrorRowsOnly)]
    [InlineData(true, ExcelImportFailureWorkbookMode.ErrorRowsOnly)]
    [InlineData(false, ExcelImportFailureWorkbookMode.AnnotatedOriginal)]
    [InlineData(true, ExcelImportFailureWorkbookMode.AnnotatedOriginal)]
    public async Task ForwardEntityFailure_ShouldCommitThroughMatchingEntry(bool asynchronous,
        ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = Source(false);
        var committer = new RecordingCommitter();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            null, committer));
        var path = Path.Combine(directory.Path, "forward-entity-failure.xlsx");
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = mode,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.True(File.Exists(path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证组合页眉和每页尾部的 Forward 失败工作簿使用对应提交入口。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [MemberData(nameof(FormatsCallsAndModes))]
    public async Task ForwardHeaderFooterFailure_ShouldCommitThroughMatchingEntry(bool xls,
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = CombinedSource(xls);
        var committer = new RecordingCommitter();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null,
            null, null, committer));
        var path = Path.Combine(directory.Path,
            xls ? "forward-header-footer.xls" : "forward-header-footer.xlsx");
        var options = CombinedOptions(directory.Path, path, mode);
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateCombinedLayout(), batchOptions,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateCombinedLayout(), batchOptions, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.True(File.Exists(path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证组合页眉和每页尾部提交失败时保留原目标并只观察一次异常。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [MemberData(nameof(FormatsCallsAndModes))]
    public async Task ForwardHeaderFooterCommitFailure_ShouldObserveOnce(bool xls,
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = CombinedSource(xls);
        var path = Path.Combine(directory.Path,
            xls ? "forward-header-footer.xls" : "forward-header-footer.xlsx");
        var sentinel = new byte[] { 8, 6, 4 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null,
            null, new[] { observer }, committer));
        var batchOptions = new ExcelEntityBatchImportOptions(1,
            CombinedOptions(directory.Path, path, mode), ExcelEntityBatchParsingMode.Forward);

        BingOfficesFileCommitException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateCombinedLayout(), batchOptions,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, CreateCombinedLayout(), batchOptions,
                    _ => { }));
        }

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证组合页眉和每页尾部写入完成后取消时保留原目标并清理暂存文件。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [MemberData(nameof(FormatsCallsAndModes))]
    public async Task ForwardHeaderFooterCancellationDuringWrite_ShouldPreserveTarget(bool xls,
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = CombinedSource(xls);
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path,
            xls ? "forward-header-footer.xls" : "forward-header-footer.xlsx");
        var sentinel = new byte[] { 3, 1, 4 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null,
            null, null, committer));
        var batchOptions = new ExcelEntityBatchImportOptions(1,
            CombinedOptions(directory.Path, path, mode), ExcelEntityBatchParsingMode.Forward);

        OperationCanceledException exception;
        if (asynchronous)
        {
            exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateCombinedLayout(), batchOptions,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            exception = Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateCombinedLayout(), batchOptions,
                    _ => { }, cancellation.Token));
        }

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX Forward 提交失败时只观察一次异常并保留原目标。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [InlineData(false, ExcelImportFailureWorkbookMode.ErrorRowsOnly)]
    [InlineData(true, ExcelImportFailureWorkbookMode.ErrorRowsOnly)]
    [InlineData(false, ExcelImportFailureWorkbookMode.AnnotatedOriginal)]
    [InlineData(true, ExcelImportFailureWorkbookMode.AnnotatedOriginal)]
    public async Task ForwardCommitFailure_ShouldObserveOnce(bool asynchronous,
        ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = Source(false);
        var path = Path.Combine(directory.Path, "forward-entity-failure.xlsx");
        var sentinel = new byte[] { 6, 4, 2 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            new[] { observer }, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = mode,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Forward);

        BingOfficesFileCommitException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { }));
        }

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要实体批次提交失败时只观察一次异常并保留原目标。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task BufferedSummaryCommitFailure_ShouldObserveOnce(bool xls,
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        var path = Path.Combine(directory.Path, xls ? "buffered-failure.xls" : "buffered-failure.xlsx");
        var sentinel = new byte[] { 4, 5, 6 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            new[] { observer }, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        BingOfficesFileCommitException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { }));
        }

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要实体批次提交后取消时保留原目标并清理暂存文件。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task BufferedSummaryCancellationDuringCommit_ShouldPreserveTarget(bool xls,
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = Source(xls);
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, xls ? "buffered-failure.xls" : "buffered-failure.xlsx");
        var sentinel = new byte[] { 5, 6, 7 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            null, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        OperationCanceledException exception;
        if (asynchronous)
        {
            exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            exception = Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { },
                    cancellation.Token));
        }

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要失败工作簿输出到调用方流时绕过文件提交器。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task BufferedSummaryCallerStream_ShouldBypassCommitter(bool xls,
        bool asynchronous)
    {
        using var source = Source(xls);
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            null, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = output
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        ExcelEntityBatchImportSummary<EntityDocument> summary;
        if (asynchronous)
        {
            summary = await importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                (_, _) => Task.CompletedTask);
        }
        else
        {
            summary = importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { });
        }

        Assert.False(summary.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.True(output.Length > 0);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要模式从不可定位输入写入调用方流时不调用文件提交器。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task BufferedSummaryNonSeekableCallerStream_ShouldBypassCommitter(bool xls,
        bool asynchronous)
    {
        using var sourceBytes = Source(xls);
        using var source = new NonSeekableReadStream(sourceBytes.ToArray());
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            null, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = output
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.True(output.Length > 0);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX Forward 无错误时不调用失败工作簿提交器。
    /// </summary>
    [Fact]
    public void ForwardSuccessful_ShouldNotCallCommitter()
    {
        using var directory = new TestDirectory();
        using var source = Source(false, valid: true);
        var committer = new RecordingCommitter();
        var importer = new NpoiEntityBatchImporter(new NpoiExcelImporter(null, null, null, null,
            null, committer));
        var path = Path.Combine(directory.Path, "unused-forward-entity-failure.xlsx");
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            _ => { });

        Assert.True(summary.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.False(File.Exists(path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证同步和异步入口使用对应的提交方法。
    /// </summary>
    /// <param name="committer">记录提交调用的提交器。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    private static void AssertCalls(RecordingCommitter committer, bool asynchronous)
    {
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
    }

    /// <summary>
    /// 按指定模式执行实体导入。
    /// </summary>
    /// <param name="importer">执行导入的导入器。</param>
    /// <param name="source">来源工作簿流。</param>
    /// <param name="options">实体导入选项。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    /// <param name="token">导入取消令牌。</param>
    /// <returns>实体导入结果。</returns>
    private static async Task<ExcelEntityImportResult<EntityDocument>> Import(
        NpoiExcelImporter importer, Stream source, ExcelEntityImportOptions options,
        bool asynchronous, CancellationToken token = default) => asynchronous
        ? await importer.ImportEntityAsync(source, CreateLayout(), options, token)
        : importer.ImportEntity(source, CreateLayout(), options, token);

    /// <summary>
    /// 创建失败工作簿导出选项。
    /// </summary>
    /// <param name="directory">暂存文件目录。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <returns>实体导入选项。</returns>
    private static ExcelEntityImportOptions Options(string directory, string path) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                DestinationPath = path,
                TemporaryDirectory = directory
            });

    /// <summary>
    /// 创建实体布局。
    /// </summary>
    /// <returns>包含明细列表区域的实体布局。</returns>
    private static ExcelEntityLayout<EntityDocument> CreateLayout() =>
        ExcelEntity.Layout<EntityDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含实际重复页眉和每页尾部的实体布局。
    /// </summary>
    /// <returns>组合页眉和尾部布局。</returns>
    private static ExcelEntityLayout<EntityDocument> CreateCombinedLayout() =>
        ExcelEntity.Layout<EntityDocument>(builder => builder
            .ListRegion("明细", "A2", document => document.Lines, region => region
                .End("B8")
                .PageBreak(1)
                .PageHeader("A1:B2", "A1", "HDR")
                .PageFooter("SIGN")));

    /// <summary>
    /// 创建组合页眉和每页尾部使用的失败工作簿选项。
    /// </summary>
    /// <param name="directory">暂存文件目录。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <param name="mode">失败工作簿模式。</param>
    /// <returns>实体导入选项。</returns>
    private static ExcelEntityImportOptions CombinedOptions(string directory, string path,
        ExcelImportFailureWorkbookMode mode) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = mode,
                DestinationPath = path,
                TemporaryDirectory = directory
            });

    /// <summary>
    /// 创建用于实体导入的工作簿流。
    /// </summary>
    /// <param name="xls">是否生成 XLS 工作簿。</param>
    /// <param name="valid">是否生成通过校验的数据。</param>
    /// <returns>定位到起始位置的工作簿流。</returns>
    private static MemoryStream Source(bool xls, bool valid = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var data = sheet.CreateRow(1);
        data.CreateCell(0).SetCellValue(valid ? "A" : string.Empty);
        data.CreateCell(1).SetCellValue(1);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return new MemoryStream(output.ToArray(), writable: true);
    }

    /// <summary>
    /// 创建首个明细编码为空的组合页眉失败工作簿。
    /// </summary>
    /// <param name="xls">是否生成 XLS 工作簿。</param>
    /// <returns>位于起始位置的工作簿流。</returns>
    private static MemoryStream CombinedSource(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        WriteCombinedRow(sheet, 0, "HDR", "采购单");
        WriteCombinedRow(sheet, 1, "Code", "Quantity");
        WriteCombinedRow(sheet, 2, string.Empty, 1d);
        WriteCombinedRow(sheet, 3, "SIGN", null);
        WriteCombinedRow(sheet, 4, "HDR", "采购单");
        WriteCombinedRow(sheet, 5, "Code", "Quantity");
        WriteCombinedRow(sheet, 6, "B-1", 2d);
        WriteCombinedRow(sheet, 7, "SIGN", null);

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return new MemoryStream(output.ToArray(), writable: true);
    }

    /// <summary>
    /// 写入组合页眉失败工作簿中的一行文本。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="rowIndex">零基行号。</param>
    /// <param name="code">编码文本。</param>
    /// <param name="quantity">数量文本。</param>
    private static void WriteCombinedRow(ISheet sheet, int rowIndex, string code,
        object quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code ?? string.Empty);
        if (quantity is double number)
            row.CreateCell(1).SetCellValue(number);
        else if (quantity is string text)
            row.CreateCell(1).SetCellValue(text);
    }

    /// <summary>
    /// 断言实体失败工作簿的结构和错误标注。
    /// </summary>
    /// <param name="path">失败工作簿路径。</param>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="result">实体导入结果。</param>
    private static void AssertEntityFailureWorkbook(string path, bool xls,
        ExcelEntityImportResult<EntityDocument> result)
    {
        using var input = File.OpenRead(path);
        using var workbook = WorkbookFactory.Create(input);
        Assert.False(result.IsSuccess);
        Assert.NotEmpty(result.Errors);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(1d, sheet.GetRow(1).GetCell(1).NumericCellValue);
        Assert.Equal(string.Empty, sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.NotNull(sheet.GetRow(1).GetCell(0).CellComment);
        Assert.Equal(xls ? ".xls" : ".xlsx", Path.GetExtension(path));
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
    }

    private sealed class EntityDocument
    {
        /// <summary>
        /// 获取明细集合。
        /// </summary>
        public List<EntityLine> Lines { get; } = new();
    }

    /// <summary>
    /// 提供不可定位但可读取的失败工作簿输入流。
    /// </summary>
    private sealed class NonSeekableReadStream : Stream
    {
        /// <summary>
        /// 保存实际输入内容的只读流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个 <see cref="NonSeekableReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="buffer">输入工作簿内容。</param>
        internal NonSeekableReadStream(byte[] buffer)
        {
            _inner = new MemoryStream(buffer, writable: false);
        }

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
        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) =>
            _inner.ReadAsync(buffer, offset, count, cancellationToken);

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default) =>
            _inner.ReadAsync(buffer, cancellationToken);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class EntityLine
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
    /// 记录提交调用并委托默认提交器。
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
            "npoi-entity-committer-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, true);
    }
}
