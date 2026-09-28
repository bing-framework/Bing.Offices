using System.IO;

namespace Bing.Offices.IO;

/// <summary>
/// 异步 staging 的最小读写合同。
/// </summary>
internal interface INpoiAsyncStaging : IDisposable
{
    /// <summary>
    /// 获取用于写入 staging 内容的流。
    /// </summary>
    Stream WriteStream { get; }

    /// <summary>
    /// 异步刷新 staging 内容并准备读取。
    /// </summary>
    /// <param name="cancellationToken">用于取消刷新的令牌。</param>
    Task FlushAsync(CancellationToken cancellationToken);

    /// <summary>
    /// 将 staging 内容异步复制到目标流。
    /// </summary>
    /// <param name="destination">接收 staging 内容的目标流。</param>
    /// <param name="cancellationToken">用于取消复制的令牌。</param>
    Task CopyToAsync(Stream destination, CancellationToken cancellationToken);
}
