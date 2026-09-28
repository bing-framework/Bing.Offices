using System.Diagnostics;
using System.Globalization;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Imports;

/// <summary>
/// 表示已完成序列化、等待复制到调用方流的失败工作簿临时产物。
/// </summary>
internal sealed class ClosedXmlFailureWorkbookArtifact : IDisposable
{
    /// <summary>
    /// 保存完整失败工作簿的临时流，由当前产物负责释放。
    /// </summary>
    private readonly Stream _stream;
    /// <summary>
    /// 失败工作簿临时文件路径，用于释放产物时清理。
    /// </summary>
    private readonly string _path;
    /// <summary>
    /// 失败工作簿的临时文件清理与诊断选项。
    /// </summary>
    private readonly ExcelImportFailureOptions _options;
    /// <summary>
    /// 标记临时产物是否已释放，防止重复清理。
    /// </summary>
    private bool _disposed;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlFailureWorkbookArtifact"/> 类型的实例。
    /// </summary>
    /// <param name="stream">保存完整失败工作簿的临时流。</param>
    /// <param name="path">待清理的临时文件路径。</param>
    /// <param name="options">失败工作簿输出及清理选项。</param>
    internal ClosedXmlFailureWorkbookArtifact(Stream stream, string path, ExcelImportFailureOptions options)
    {
        _stream = stream;
        _path = path;
        _options = options;
    }

    /// <summary>
    /// 将失败工作簿复制到目标流。
    /// </summary>
    /// <param name="destination">接收失败工作簿内容的可写流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal void CopyTo(Stream destination, CancellationToken cancellationToken)
    {
        ValidateDestination(destination);
        try
        {
            var buffer = new byte[81920];
            int count;
            while ((count = _stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                destination.Write(buffer, 0, count);
                cancellationToken.ThrowIfCancellationRequested();
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("失败工作簿复制到目标流失败。", exception,
                "ClosedXML", BingOfficesStage.Write);
        }
    }

    /// <summary>
    /// 异步将失败工作簿复制到目标流。
    /// </summary>
    /// <param name="destination">接收失败工作簿内容的可写流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal async Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
    {
        ValidateDestination(destination);
        try
        {
            await _stream.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("失败工作簿异步复制到目标流失败。", exception,
                "ClosedXML", BingOfficesStage.Write);
        }
    }

    /// <summary>
    /// 校验失败工作簿目标流可写。
    /// </summary>
    /// <param name="destination">接收失败工作簿内容的可写流。</param>
    private static void ValidateDestination(Stream destination)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _stream.Dispose();
        ClosedXmlFailureWorkbookWriter.DeleteTemporaryFile(_path, _options, null);
    }
}
