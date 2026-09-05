using System;
using System.IO;
using System.Threading;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Npoi;

/// <summary>
/// NPOI 管线内部流复制协作者，统一取消和输入大小限制处理。
/// </summary>
internal static class NpoiStreamCopier
{
    /// <summary>
    /// 将源流复制到目标流，保持调用方流所有权，并在块边界检查取消和大小限制。
    /// </summary>
    /// <param name="source">调用方拥有的可读源流。</param>
    /// <param name="destination">实现拥有的目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="maxBytes">允许复制的最大字节数；为空表示不额外限制。</param>
    public static void Copy(Stream source, Stream destination, CancellationToken cancellationToken,
        long? maxBytes = null)
    {
        var buffer = new byte[81920];
        long total = 0;
        int count;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            count = source.Read(buffer, 0, buffer.Length);
            if (count == 0)
                break;
            total += count;
            if (maxBytes.HasValue && total > maxBytes.Value)
                throw new BingOfficesResourceLimitException($"输入工作簿超过最大字节数: {maxBytes.Value}",
                    provider: "NPOI", operation: BingOfficesOperation.Import, stage: BingOfficesStage.Open);
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
        }
    }
}
