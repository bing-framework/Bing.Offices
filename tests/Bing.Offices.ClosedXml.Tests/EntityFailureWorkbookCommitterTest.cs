using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Entities;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML Entity 失败工作簿的文件提交和取消边界。
/// </summary>
public sealed class EntityFailureWorkbookCommitterTest
{
    /// <summary>
    /// 枚举同步、异步与失败工作簿模式组合。
    /// </summary>
    /// <returns>测试参数集合。</returns>
    public static IEnumerable<object[]> CallsAndModes()
    {
        foreach (var asynchronous in new[] { false, true })
        foreach (var mode in new[]
                 {
                     ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                     ExcelImportFailureWorkbookMode.AnnotatedOriginal
                 })
            yield return new object[] { asynchronous, mode };
    }

    /// <summary>
    /// 验证同步和异步实体导入使用对应的失败工作簿提交入口。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EntityFailure_ShouldCommitThroughMatchingEntry(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "entity-failure.xlsx");
        var committer = new RecordingCommitter();
        var importer = CreateImporter(committer);

        var result = await Import(importer, source, path, directory.Path, asynchronous,
            cancellation.Token);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.Equal(cancellation.Token, committer.Token);
        Assert.NotNull(committer.Bytes);
        AssertFailureWorkbook(committer.Bytes, result.Errors.Single().Message);
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体导入预取消时不会进入文件提交器。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreCanceled_ShouldNotCommitEntityFailureWorkbook(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var path = Path.Combine(directory.Path, "entity-failure.xlsx");
        var committer = new RecordingCommitter();
        var importer = CreateImporter(committer);

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Import(importer, source, path, directory.Path, asynchronous, cancellation.Token));

        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体失败工作簿写入期间取消会保留原目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDuringCommit_ShouldPreserveTargetAndCleanup(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "entity-failure.xlsx");
        var sentinel = new byte[] { 2, 4, 6, 8 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var importer = CreateImporter(committer);

        var exception = await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Import(importer, source, path, directory.Path, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(path, committer.Path);
        Assert.Equal("FailureWorkbook", committer.Format);
        Assert.Equal(cancellation.Token, committer.Token);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体失败工作簿提交失败时保留原目标文件。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CommitFailure_ShouldPreserveEntityTarget(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        var path = Path.Combine(directory.Path, "entity-failure.xlsx");
        var sentinel = new byte[] { 9, 7, 5 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var importer = CreateImporter(committer);

        var actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            Import(importer, source, path, directory.Path, asynchronous, CancellationToken.None));

        Assert.Same(expected, actual);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
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
        using var source = CreateInvalidWorkbook();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
        var path = Path.Combine(directory.Path, "forward-entity-failure.xlsx");
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = mode,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, batchOptions,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, batchOptions, _ => { });

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
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [MemberData(nameof(CallsAndModes))]
    public async Task ForwardHeaderFooterFailure_ShouldCommitThroughMatchingEntry(bool asynchronous,
        ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = CreateCombinedWorkbook();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var path = Path.Combine(directory.Path, "forward-header-footer.xlsx");
        var batchOptions = new ExcelEntityBatchImportOptions(1,
            CombinedOptions(directory.Path, path, mode), ExcelEntityBatchParsingMode.Forward);

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
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [MemberData(nameof(CallsAndModes))]
    public async Task ForwardHeaderFooterCommitFailure_ShouldObserveOnce(bool asynchronous,
        ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = CreateCombinedWorkbook();
        var path = Path.Combine(directory.Path, "forward-header-footer.xlsx");
        var sentinel = new byte[] { 8, 6, 4 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, new[] { observer }, committer));
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="mode">失败工作簿模式。</param>
    [Theory]
    [MemberData(nameof(CallsAndModes))]
    public async Task ForwardHeaderFooterCancellationDuringWrite_ShouldPreserveTarget(
        bool asynchronous, ExcelImportFailureWorkbookMode mode)
    {
        using var directory = new TestDirectory();
        using var source = CreateCombinedWorkbook();
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "forward-header-footer.xlsx");
        var sentinel = new byte[] { 3, 1, 4 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
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
        using var source = CreateInvalidWorkbook();
        var path = Path.Combine(directory.Path, "forward-entity-failure.xlsx");
        var sentinel = new byte[] { 6, 4, 2 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, new[] { observer }, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
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
                importer.ImportEntityBatchesAsync(source, layout, batchOptions,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, layout, batchOptions, _ => { }));
        }

        Assert.Same(expected, actual);
        Assert.Same(expected, Assert.Single(observer.Exceptions));
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要实体批次提交失败时只观察一次异常并保留原目标。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummaryCommitFailure_ShouldObserveOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        var path = Path.Combine(directory.Path, "buffered-failure.xlsx");
        var sentinel = new byte[] { 4, 5, 6 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new RecordingCommitter { Failure = expected };
        var observer = new RecordingObserver();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, new[] { observer }, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
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
                importer.ImportEntityBatchesAsync(source, layout, batchOptions,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, layout, batchOptions, _ => { }));
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummaryCancellationDuringCommit_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "buffered-failure.xlsx");
        var sentinel = new byte[] { 5, 6, 7 };
        File.WriteAllBytes(path, sentinel);
        var committer = new RecordingCommitter { AfterWrite = cancellation.Cancel };
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
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
                importer.ImportEntityBatchesAsync(source, layout, batchOptions,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            exception = Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, layout, batchOptions, _ => { },
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummaryCallerStream_ShouldBypassCommitter(bool asynchronous)
    {
        using var source = CreateInvalidWorkbook();
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = output
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        ExcelEntityBatchImportSummary<Root> summary;
        if (asynchronous)
        {
            summary = await importer.ImportEntityBatchesAsync(source, layout, batchOptions,
                (_, _) => Task.CompletedTask);
        }
        else
        {
            summary = importer.ImportEntityBatches(source, layout, batchOptions, _ => { });
        }

        Assert.False(summary.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.True(output.Length > 0);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要模式从不可定位输入写入调用方流时不调用文件提交器。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BufferedSummaryNonSeekableCallerStream_ShouldBypassCommitter(bool asynchronous)
    {
        using var sourceBytes = CreateInvalidWorkbook();
        using var source = new NonSeekableReadStream(sourceBytes.ToArray());
        using var output = new MemoryStream();
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = output
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, batchOptions,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, batchOptions, _ => { });

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
        using var source = CreateInvalidWorkbook(valid: true);
        var committer = new RecordingCommitter();
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
        var path = Path.Combine(directory.Path, "unused-forward-entity-failure.xlsx");
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            _ => { });

        Assert.True(summary.IsSuccess);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        Assert.False(File.Exists(path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 根据提交器创建 ClosedXML 实体导入器。
    /// </summary>
    /// <param name="committer">记录文件提交参数的提交器。</param>
    /// <returns>使用指定提交器的导入器。</returns>
    private static ClosedXmlExcelImporter CreateImporter(IFileExportCommitter committer) =>
        new(null, null, null, null, null, committer);

    /// <summary>
    /// 创建含一个必填字段错误的真实 XLSX 工作簿。
    /// </summary>
    /// <returns>位于起始位置的输入流。</returns>
    private static MemoryStream CreateInvalidWorkbook(bool valid = false)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = "Code";
            sheet.Cell("B1").Value = "Quantity";
            sheet.Cell("A2").Value = valid ? "A-1" : string.Empty;
            sheet.Cell("B2").Value = 7;
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建包含实际重复页眉和每页尾部的实体布局。
    /// </summary>
    /// <returns>组合页眉和尾部布局。</returns>
    private static ExcelEntityLayout<Root> CreateCombinedLayout() =>
        ExcelEntity.Layout<Root>(builder => builder
            .ListRegion("明细", "A2", root => root.Rows, region => region
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
    /// 创建首个明细编码为空的组合页眉失败工作簿。
    /// </summary>
    /// <returns>位于起始位置的工作簿流。</returns>
    private static MemoryStream CreateCombinedWorkbook()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("明细");
            sheet.Cell("A1").Value = "HDR";
            sheet.Cell("B1").Value = "采购单";
            sheet.Cell("A2").Value = "Code";
            sheet.Cell("B2").Value = "Quantity";
            sheet.Cell("A3").Value = string.Empty;
            sheet.Cell("B3").Value = 1;
            sheet.Cell("A4").Value = "SIGN";
            sheet.Cell("A5").Value = "HDR";
            sheet.Cell("B5").Value = "采购单";
            sheet.Cell("A6").Value = "Code";
            sheet.Cell("B6").Value = "Quantity";
            sheet.Cell("A7").Value = "B-1";
            sheet.Cell("B7").Value = 2;
            sheet.Cell("A8").Value = "SIGN";
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 按指定同步或异步入口执行实体导入。
    /// </summary>
    /// <param name="importer">执行导入的 Provider。</param>
    /// <param name="source">真实 XLSX 输入流。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <param name="temporaryDirectory">失败工作簿临时目录。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="token">传给导入器的取消令牌。</param>
    /// <returns>实体导入结果。</returns>
    private static async Task<ExcelEntityImportResult<Root>> Import(
        ClosedXmlExcelImporter importer, Stream source, string path, string temporaryDirectory,
        bool asynchronous, CancellationToken token)
    {
        var layout = ExcelEntity.Layout<Root>(builder =>
            builder.ListRegion("Data", "A1", root => root.Rows));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = temporaryDirectory
            });

        return asynchronous
            ? await importer.ImportEntityAsync(source, layout, options, token).ConfigureAwait(false)
            : importer.ImportEntity(source, layout, options, token);
    }

    /// <summary>
    /// 读取提交后的失败工作簿并验证原始行及错误摘要。
    /// </summary>
    /// <param name="bytes">已提交的 XLSX 内容。</param>
    /// <param name="message">导入结果中的错误消息。</param>
    private static void AssertFailureWorkbook(byte[] bytes, string message)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        Assert.Equal(new[] { "Data", "_ImportErrors" },
            workbook.Worksheets.Select(sheet => sheet.Name));
        var data = workbook.Worksheet("Data");
        Assert.Equal("Code", data.Cell("A1").GetString());
        Assert.Equal("Quantity", data.Cell("B1").GetString());
        Assert.Equal(string.Empty, data.Cell("A2").GetString());
        Assert.Equal(7, data.Cell("B2").GetValue<int>());
        Assert.False(data.Row(2).IsEmpty());
        var summary = workbook.Worksheet("_ImportErrors");
        Assert.True(summary.LastRowUsed().RowNumber() >= 2);
        Assert.Contains(summary.CellsUsed().Select(cell => cell.GetString()),
            value => value == message);
    }

    /// <summary>
    /// 为单个测试创建并删除临时目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取本测试专用的临时目录路径。
        /// </summary>
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "closedxml-entity-failure-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, recursive: true);
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

    /// <summary>
    /// 接收实体列表的根模型。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取实体行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 表示实体失败工作簿中的一行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置必填业务编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量值。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 记录 Provider 观察到的公共异常。
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
    /// 记录提交入口，并委托给真实原子文件提交器。
    /// </summary>
    private sealed class RecordingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 执行真实文件提交的默认实现。
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
        /// 获取最近一次提交目标路径。
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
        /// 获取默认提交器完成后写入目标的工作簿字节。
        /// </summary>
        public byte[] Bytes { get; private set; }

        /// <summary>
        /// 获取或设置写入完成后的测试回调。
        /// </summary>
        public Action AfterWrite { get; set; }

        /// <summary>
        /// 获取或设置提交入口抛出的异常。
        /// </summary>
        public Exception Failure { get; set; }

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
                await writeAsync(stream, token).ConfigureAwait(false);
                AfterWrite?.Invoke();
            }, cancellationToken, format).ConfigureAwait(false);
            Bytes = File.ReadAllBytes(path);
        }

        /// <summary>
        /// 保存提交参数，供职责断言读取。
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
}
