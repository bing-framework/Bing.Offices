using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Tests.Exports.PivotTables.Resources;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 透视导出文件提交、取消与调用方流边界。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotResourceTest
{
    /// <summary>
    /// 验证同步和异步文件提交保留完整透视缓存并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_Success_ShouldCommitCompletePivot(bool asynchronous)
    {
        using var directory = new NpoiPivotResourceTestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var observer = new NpoiPivotRecordingObserver();
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);

        await ExportToFile(exporter, NpoiPivotPublicExportTest.Request(true), target, asynchronous);

        Assert.Empty(observer.Exceptions);
        AssertCommitCall(committer, target, CancellationToken.None, asynchronous);
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        NpoiPivotPublicExportTest.AssertWorkbook(File.ReadAllBytes(target), true);
    }

    /// <summary>
    /// 验证真实暂存写入失败保留旧目标并只观察导出异常一次。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_TemporaryWriteFailure_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new NpoiPivotResourceTestDirectory();
        var target = SeedTarget(directory);
        var observer = new NpoiPivotRecordingObserver();
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.FailDuringWrite);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
            ExportToFile(exporter, NpoiPivotPublicExportTest.Request(true), target, asynchronous));

        Assert.Equal(BingOfficesStage.Write, exception.Stage);
        Assert.IsType<IOException>(exception.InnerException);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.True(committer.BytesWritten > 0);
        AssertCommitCall(committer, target, CancellationToken.None, asynchronous);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证真实暂存写入期间取消保留旧目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CanceledDuringTemporaryWrite_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new NpoiPivotResourceTestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = SeedTarget(directory);
        var observer = new NpoiPivotRecordingObserver();
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.CancelDuringWrite, cancellation);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, NpoiPivotPublicExportTest.Request(true), target, asynchronous,
                cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Empty(observer.Exceptions);
        Assert.True(committer.BytesWritten > 0);
        AssertCommitCall(committer, target, cancellation.Token, asynchronous);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证完整序列化后提交回调失败保留旧目标并只观察一次。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_FailureAfterSerialization_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new NpoiPivotResourceTestDirectory();
        var target = SeedTarget(directory);
        var observer = new NpoiPivotRecordingObserver();
        var failure = new BingOfficesFileCommitException("注入透视文件提交失败。", provider: "NPOI");
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.FailAfterSerialization,
            failure: failure);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            ExportToFile(exporter, NpoiPivotPublicExportTest.Request(true), target, asynchronous));

        Assert.Same(failure, exception);
        Assert.Equal(BingOfficesStage.Commit, exception.Stage);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.True(committer.SerializedBytes > 0);
        AssertCommitCall(committer, target, CancellationToken.None, asynchronous);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证完整序列化后取消保留旧目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CanceledAfterSerialization_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new NpoiPivotResourceTestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = SeedTarget(directory);
        var observer = new NpoiPivotRecordingObserver();
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.CancelAfterSerialization,
            cancellation);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, NpoiPivotPublicExportTest.Request(true), target, asynchronous,
                cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Empty(observer.Exceptions);
        Assert.True(committer.SerializedBytes > 0);
        AssertCommitCall(committer, target, cancellation.Token, asynchronous);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证同步和异步流导出完整写入透视且不关闭调用方流。
    /// </summary>
    /// <param name="asynchronous">是否使用异步流导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StreamExport_ShouldLeaveCallerStreamOpen(bool asynchronous)
    {
        using var output = new MemoryStream();
        var observer = new NpoiPivotRecordingObserver();
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);

        await ExportToStream(exporter, NpoiPivotPublicExportTest.Request(true), output, asynchronous);

        Assert.True(output.CanRead && output.CanWrite);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
        NpoiPivotPublicExportTest.AssertWorkbook(output.ToArray(), true);
    }

    /// <summary>
    /// 验证同步序列化在首次底层写入取消后立即停止。
    /// </summary>
    /// <param name="format">工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void StreamExport_CanceledDuringSynchronousSerialization_ShouldObserveAtWriteBoundary(
        ExcelFormat format)
    {
        using var cancellation = new CancellationTokenSource();
        using var underlying = new MemoryStream();
        using var output = new NpoiPivotInjectedWriteStream(underlying,
            _ => cancellation.Cancel());
        var observer = new NpoiPivotRecordingObserver();
        var committer = new NpoiPivotRecordingCommitter(NpoiPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);
        var request = ExcelExport.Workbook(builder => builder
            .Format(format)
            .AddSheet("Data", new[]
            {
                new NpoiPivotExportRow
                {
                    Region = "North", Amount = 10, State = "Open", Channel = "Retail", Code = "A"
                }
            }));

        var exception = Assert.Throws<OperationCanceledException>(() =>
            exporter.Export(request, output, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(1, output.WriteOperations);
        Assert.True(output.BytesWritten > 0);
        Assert.True(output.CanWrite && underlying.CanWrite);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(0, committer.SyncCalls + committer.AsyncCalls);
    }

    /// <summary>
    /// 创建配置异常观察器和原子文件提交器的 NPOI Provider。
    /// </summary>
    /// <param name="observer">记录导出异常的观察器。</param>
    /// <param name="committer">包装真实默认文件提交器的测试提交器。</param>
    /// <returns>配置完成的 NPOI Provider。</returns>
    private static NpoiExcelExporter CreateExporter(NpoiPivotRecordingObserver observer,
        NpoiPivotRecordingCommitter committer) => new(exceptionObservers: new[] { observer },
        fileExportCommitter: committer);

    /// <summary>
    /// 根据参数调用同步或异步文件导出入口。
    /// </summary>
    /// <param name="exporter">执行导出的 NPOI Provider。</param>
    /// <param name="request">基础或高级透视请求。</param>
    /// <param name="target">最终目标文件路径。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    /// <param name="cancellationToken">文件提交使用的取消令牌。</param>
    private static async Task ExportToFile(NpoiExcelExporter exporter, ExcelWorkbookExportRequest request,
        string target, bool asynchronous, CancellationToken cancellationToken = default)
    {
        if (asynchronous)
            await exporter.ExportToFileAsync(request, target, cancellationToken).ConfigureAwait(false);
        else
            exporter.ExportToFile(request, target, cancellationToken);
    }

    /// <summary>
    /// 根据参数调用同步或异步流导出入口。
    /// </summary>
    /// <param name="exporter">执行导出的 NPOI Provider。</param>
    /// <param name="request">基础或高级透视请求。</param>
    /// <param name="destination">由调用方拥有的目标流。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    private static async Task ExportToStream(NpoiExcelExporter exporter, ExcelWorkbookExportRequest request,
        Stream destination, bool asynchronous)
    {
        if (asynchronous)
            await exporter.ExportAsync(request, destination).ConfigureAwait(false);
        else
            exporter.Export(request, destination);
    }

    /// <summary>
    /// 验证提交器收到匹配入口、路径、格式和取消令牌的单次调用。
    /// </summary>
    /// <param name="committer">记录调用信息的提交器。</param>
    /// <param name="target">请求的目标路径。</param>
    /// <param name="token">请求使用的取消令牌。</param>
    /// <param name="asynchronous">是否调用异步提交入口。</param>
    private static void AssertCommitCall(NpoiPivotRecordingCommitter committer, string target,
        CancellationToken token, bool asynchronous)
    {
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(target, committer.LastPath);
        Assert.Equal("Excel", committer.LastFormat);
        Assert.Equal(token, committer.LastToken);
    }

    /// <summary>
    /// 创建包含固定旧字节的目标文件。
    /// </summary>
    /// <param name="directory">隔离的测试目录。</param>
    /// <returns>旧目标的绝对路径。</returns>
    private static string SeedTarget(NpoiPivotResourceTestDirectory directory)
    {
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        File.WriteAllBytes(target, new byte[] { 2, 3, 5, 7, 11 });
        return target;
    }

    /// <summary>
    /// 验证旧目标字节未变且真实默认提交器没有留下暂存文件。
    /// </summary>
    /// <param name="directory">隔离的测试目录。</param>
    /// <param name="target">保留原内容的目标路径。</param>
    private static void AssertOldTargetAndNoTemporaryFile(NpoiPivotResourceTestDirectory directory,
        string target)
    {
        Assert.Equal(new byte[] { 2, 3, 5, 7, 11 }, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }
}
