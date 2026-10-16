using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Entities;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 探测前向导入格式并在必要时创建磁盘暂存输入。
/// </summary>
internal static class NpoiForwardInputDetector
{
    private static readonly byte[] Ole2Signature =
        { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    /// <summary>
    /// 同步打开前向导入输入。
    /// </summary>
    /// <param name="source">调用方拥有的源流。</param>
    /// <param name="maxBytes">允许读取的最大字节数。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="temporaryStorage">与后续解析共享的临时存储预算。</param>
    /// <returns>带格式标识和所有权信息的输入。</returns>
    internal static NpoiForwardInput Open(Stream source, long? maxBytes,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        ValidateSource(source);
        if (CanReadInPlace(source, maxBytes))
            return new NpoiForwardInput(source, IsHssf(source), ownsStream: false);

        var buffer = CreateTempBuffer(temporaryStorage);
        try
        {
            NpoiStreamCopier.Copy(source, buffer, cancellationToken, maxBytes);
            buffer.Flush();
            buffer.Position = 0;
            return new NpoiForwardInput(buffer, IsHssf(buffer), ownsStream: true);
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 异步打开前向导入输入。
    /// </summary>
    /// <param name="source">调用方拥有的源流。</param>
    /// <param name="maxBytes">允许读取的最大字节数。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="temporaryStorage">与后续解析共享的临时存储预算。</param>
    /// <returns>带格式标识和所有权信息的输入任务。</returns>
    internal static async Task<NpoiForwardInput> OpenAsync(Stream source, long? maxBytes,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        ValidateSource(source);
        if (CanReadInPlace(source, maxBytes))
            return new NpoiForwardInput(source, IsHssf(source), ownsStream: false);

        var buffer = CreateTempBuffer(temporaryStorage);
        try
        {
            await NpoiStreamCopier.CopyAsync(source, buffer, cancellationToken, maxBytes)
                .ConfigureAwait(false);
            await buffer.FlushAsync(cancellationToken).ConfigureAwait(false);
            buffer.Position = 0;
            return new NpoiForwardInput(buffer, IsHssf(buffer), ownsStream: true);
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 校验输入流可读。
    /// </summary>
    private static void ValidateSource(Stream source)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
    }

    /// <summary>
    /// 判断是否可以从调用方流的当前位置直接读取。
    /// </summary>
    private static bool CanReadInPlace(Stream source, long? maxBytes)
    {
        if (!source.CanSeek || source.Position != 0)
            return false;
        if (maxBytes.HasValue && source.Length > maxBytes.Value)
            throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                provider: "NPOI", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Open);
        return true;
    }

    /// <summary>
    /// 读取输入签名并恢复流位置。
    /// </summary>
    private static bool IsHssf(Stream source)
    {
        var position = source.Position;
        try
        {
            var signature = new byte[Ole2Signature.Length];
            var count = 0;
            while (count < signature.Length)
            {
                var read = source.Read(signature, count, signature.Length - count);
                if (read == 0)
                    break;
                count += read;
            }
            if (count != signature.Length)
                return false;
            for (var index = 0; index < signature.Length; index++)
            {
                if (signature[index] != Ole2Signature[index])
                    return false;
            }
            return true;
        }
        finally
        {
            source.Position = position;
        }
    }

    /// <summary>
    /// 创建由所有权对象自动删除的输入暂存文件。
    /// </summary>
    private static Stream CreateTempBuffer(ExcelTemporaryStorageBudget temporaryStorage)
    {
        var path = Path.Combine(Path.GetTempPath(),
            $"bing-offices-forward-input-{Guid.NewGuid():N}.bin");
        return (temporaryStorage ?? new ExcelTemporaryStorageBudget(null, "NPOI")).CreateFile(path,
            FileOptions.Asynchronous | FileOptions.SequentialScan | FileOptions.DeleteOnClose);
    }
}
