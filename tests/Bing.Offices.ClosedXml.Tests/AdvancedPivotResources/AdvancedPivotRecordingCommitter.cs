using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.IO;

namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;

/// <summary>
/// 记录提交调用并在真实默认暂存提交中注入取消或失败。
/// </summary>
internal sealed class AdvancedPivotRecordingCommitter : IFileExportCommitter
{
    /// <summary>
    /// 委托真实同目录暂存、刷新、替换和清理的默认提交器。
    /// </summary>
    private readonly DefaultFileExportCommitter _inner = new();

    /// <summary>
    /// 指定提交调用要执行的测试行为。
    /// </summary>
    private readonly AdvancedPivotCommitMode _mode;

    /// <summary>
    /// 需要取消或提交后取消时使用的令牌源。
    /// </summary>
    private readonly CancellationTokenSource _cancellation;

    /// <summary>
    /// 完整序列化后抛出的提交异常。
    /// </summary>
    private readonly BingOfficesFileCommitException _failure;

    /// <summary>
    /// 实际暂存写入期间拦截写操作的流。
    /// </summary>
    private AdvancedPivotCancelingWriteStream _cancelingStream;

    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotRecordingCommitter" /> 类型的实例。
    /// </summary>
    /// <param name="mode">提交测试行为。</param>
    /// <param name="cancellation">可选的测试取消令牌源。</param>
    /// <param name="failure">可选的提交失败异常。</param>
    public AdvancedPivotRecordingCommitter(AdvancedPivotCommitMode mode,
        CancellationTokenSource cancellation = null, BingOfficesFileCommitException failure = null)
    {
        _mode = mode;
        _cancellation = cancellation;
        _failure = failure;
    }

    /// <summary>
    /// 获取同步提交调用次数。
    /// </summary>
    public int SyncCalls { get; private set; }

    /// <summary>
    /// 获取异步提交调用次数。
    /// </summary>
    public int AsyncCalls { get; private set; }

    /// <summary>
    /// 获取最后一次提交接收的取消令牌。
    /// </summary>
    public CancellationToken LastToken { get; private set; }

    /// <summary>
    /// 获取取消前已写入暂存文件的字节数。
    /// </summary>
    public long BytesWritten => _cancelingStream?.BytesWritten ?? 0;

    /// <summary>
    /// 获取提交动作前完整序列化的字节数。
    /// </summary>
    public long SerializedBytes { get; private set; }

    /// <inheritdoc />
    public void Commit(string path, Action<Stream> write,
        CancellationToken cancellationToken, string format)
    {
        SyncCalls++;
        LastToken = cancellationToken;
        _inner.Commit(path, destination => InvokeWrite(destination, write), cancellationToken, format);
    }

    /// <inheritdoc />
    public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
        CancellationToken cancellationToken, string format)
    {
        AsyncCalls++;
        LastToken = cancellationToken;
        return _inner.CommitAsync(path,
            (destination, token) => InvokeWriteAsync(destination, writeAsync, token),
            cancellationToken, format);
    }

    /// <summary>
    /// 向真实暂存文件同步写入并执行所选测试行为。
    /// </summary>
    /// <param name="destination">默认提交器创建的暂存流。</param>
    /// <param name="write">Provider 的同步写入操作。</param>
    private void InvokeWrite(Stream destination, Action<Stream> write)
    {
        if (_mode == AdvancedPivotCommitMode.CancelDuringWrite)
        {
            _cancelingStream = new AdvancedPivotCancelingWriteStream(destination, _cancellation);
            write(_cancelingStream);
            return;
        }

        write(destination);
        AfterSerialization(destination.Length);
    }

    /// <summary>
    /// 向真实暂存文件异步写入并执行所选测试行为。
    /// </summary>
    /// <param name="destination">默认提交器创建的暂存流。</param>
    /// <param name="writeAsync">Provider 的异步写入操作。</param>
    /// <param name="cancellationToken">当前提交取消令牌。</param>
    private async Task InvokeWriteAsync(Stream destination,
        Func<Stream, CancellationToken, Task> writeAsync, CancellationToken cancellationToken)
    {
        if (_mode == AdvancedPivotCommitMode.CancelDuringWrite)
        {
            _cancelingStream = new AdvancedPivotCancelingWriteStream(destination, _cancellation);
            await writeAsync(_cancelingStream, cancellationToken).ConfigureAwait(false);
            return;
        }

        await writeAsync(destination, cancellationToken).ConfigureAwait(false);
        AfterSerialization(destination.Length);
    }

    /// <summary>
    /// 在完整序列化后执行取消或提交失败行为。
    /// </summary>
    /// <param name="serializedBytes">暂存流的序列化字节长度。</param>
    private void AfterSerialization(long serializedBytes)
    {
        SerializedBytes = serializedBytes;
        if (_mode == AdvancedPivotCommitMode.CancelAfterSerialization)
            _cancellation.Cancel();
        else if (_mode == AdvancedPivotCommitMode.FailAfterSerialization)
            throw _failure;
    }
}
