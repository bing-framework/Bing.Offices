using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.IO;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Resources;

/// <summary>
/// 记录提交参数并通过真实默认提交器注入透视文件故障。
/// </summary>
internal sealed class NpoiPivotRecordingCommitter : IFileExportCommitter
{
    /// <summary>
    /// 委托真实同目录暂存、刷新、替换和失败清理的默认提交器。
    /// </summary>
    private readonly DefaultFileExportCommitter _inner = new();

    /// <summary>
    /// 当前测试选择的提交故障阶段。
    /// </summary>
    private readonly NpoiPivotCommitMode _mode;

    /// <summary>
    /// 写入后取消使用的令牌源。
    /// </summary>
    private readonly CancellationTokenSource _cancellation;

    /// <summary>
    /// 序列化完成后注入的提交异常。
    /// </summary>
    private readonly BingOfficesFileCommitException _failure;

    /// <summary>
    /// 实际暂存写入期间拦截写操作的流。
    /// </summary>
    private NpoiPivotInjectedWriteStream _injectedStream;

    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotRecordingCommitter" /> 类型的实例。
    /// </summary>
    /// <param name="mode">提交测试行为。</param>
    /// <param name="cancellation">可选的测试取消令牌源。</param>
    /// <param name="failure">可选的序列化后提交异常。</param>
    public NpoiPivotRecordingCommitter(NpoiPivotCommitMode mode,
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
    /// 获取最后一次提交收到的目标路径。
    /// </summary>
    public string LastPath { get; private set; }

    /// <summary>
    /// 获取最后一次提交收到的格式名称。
    /// </summary>
    public string LastFormat { get; private set; }

    /// <summary>
    /// 获取最后一次提交收到的取消令牌。
    /// </summary>
    public CancellationToken LastToken { get; private set; }

    /// <summary>
    /// 获取故障前已经写入真实暂存文件的字节数。
    /// </summary>
    public long BytesWritten => _injectedStream?.BytesWritten ?? 0;

    /// <summary>
    /// 获取故障前已经完整序列化到暂存文件的字节数。
    /// </summary>
    public long SerializedBytes { get; private set; }

    /// <inheritdoc />
    public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
    {
        RecordCall(path, cancellationToken, format, asynchronous: false);
        _inner.Commit(path, destination => InvokeWrite(destination, write, cancellationToken),
            cancellationToken, format);
    }

    /// <inheritdoc />
    public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
        CancellationToken cancellationToken, string format)
    {
        RecordCall(path, cancellationToken, format, asynchronous: true);
        return _inner.CommitAsync(path,
            (destination, token) => InvokeWriteAsync(destination, writeAsync, token),
            cancellationToken, format);
    }

    /// <summary>
    /// 记录 Provider 交给提交 SPI 的实际参数。
    /// </summary>
    /// <param name="path">导出目标路径。</param>
    /// <param name="cancellationToken">提交取消令牌。</param>
    /// <param name="format">Provider 指定的输出格式名称。</param>
    /// <param name="asynchronous">是否调用异步提交入口。</param>
    private void RecordCall(string path, CancellationToken cancellationToken, string format,
        bool asynchronous)
    {
        if (asynchronous)
            AsyncCalls++;
        else
            SyncCalls++;
        LastPath = path;
        LastToken = cancellationToken;
        LastFormat = format;
    }

    /// <summary>
    /// 在真实同步暂存流上执行 Provider 写入及所选故障。
    /// </summary>
    /// <param name="destination">默认提交器创建的真实暂存流。</param>
    /// <param name="write">Provider 的工作簿写入回调。</param>
    /// <param name="cancellationToken">当前提交的取消令牌。</param>
    private void InvokeWrite(Stream destination, Action<Stream> write,
        CancellationToken cancellationToken)
    {
        if (_mode == NpoiPivotCommitMode.FailDuringWrite
            || _mode == NpoiPivotCommitMode.CancelDuringWrite)
        {
            _injectedStream = CreateInjectedStream(destination);
            write(_injectedStream);
            return;
        }

        write(destination);
        AfterSerialization(destination.Length, cancellationToken);
    }

    /// <summary>
    /// 在真实异步暂存流上执行 Provider 写入及所选故障。
    /// </summary>
    /// <param name="destination">默认提交器创建的真实暂存流。</param>
    /// <param name="writeAsync">Provider 的异步工作簿写入回调。</param>
    /// <param name="cancellationToken">当前提交的取消令牌。</param>
    private async Task InvokeWriteAsync(Stream destination,
        Func<Stream, CancellationToken, Task> writeAsync, CancellationToken cancellationToken)
    {
        if (_mode == NpoiPivotCommitMode.FailDuringWrite
            || _mode == NpoiPivotCommitMode.CancelDuringWrite)
        {
            _injectedStream = CreateInjectedStream(destination);
            await writeAsync(_injectedStream, cancellationToken).ConfigureAwait(false);
            return;
        }

        await writeAsync(destination, cancellationToken).ConfigureAwait(false);
        AfterSerialization(destination.Length, cancellationToken);
    }

    /// <summary>
    /// 创建首次真实写入后触发当前故障模式的流包装器。
    /// </summary>
    /// <param name="destination">默认提交器拥有的暂存流。</param>
    /// <returns>写入后执行故障回调的流。</returns>
    private NpoiPivotInjectedWriteStream CreateInjectedStream(Stream destination) =>
        new(destination, _ =>
        {
            if (_mode == NpoiPivotCommitMode.FailDuringWrite)
                throw new IOException("注入真实暂存写入失败。");
            _cancellation.Cancel();
            throw new OperationCanceledException(_cancellation.Token);
        });

    /// <summary>
    /// 在完整序列化后执行提交失败或取消。
    /// </summary>
    /// <param name="serializedBytes">已经写入真实暂存文件的完整序列化字节数。</param>
    /// <param name="cancellationToken">当前提交的取消令牌。</param>
    private void AfterSerialization(long serializedBytes, CancellationToken cancellationToken)
    {
        SerializedBytes = serializedBytes;
        if (_mode == NpoiPivotCommitMode.CancelAfterSerialization)
        {
            _cancellation.Cancel();
            cancellationToken.ThrowIfCancellationRequested();
        }
        else if (_mode == NpoiPivotCommitMode.FailAfterSerialization)
            throw _failure ?? throw new InvalidOperationException("提交失败测试未提供异常。");
    }
}
