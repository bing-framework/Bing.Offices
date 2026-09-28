using System.IO;

namespace Bing.Offices.IO;

/// <summary>
/// 提供 Excel 异步 staging 工厂。
/// </summary>
/// <remarks>
/// 支持内存、临时文件和按阈值迁移的混合 staging 策略。
/// </remarks>
internal sealed class NpoiAsyncStagingFactory : INpoiAsyncStagingFactory
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiAsyncStagingFactory" /> 类型的实例。
    /// </summary>
    /// <param name="strategy">内存、临时文件或混合 staging 策略。</param>
    /// <param name="hybridThresholdBytes">混合策略迁移到临时文件的字节阈值。</param>
    /// <param name="directory">临时文件目录；为空时使用系统临时目录。</param>
    internal NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy strategy,
        long hybridThresholdBytes = 8L * 1024 * 1024, string directory = null)
    {
        if (hybridThresholdBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(hybridThresholdBytes));
        Strategy = strategy;
        HybridThresholdBytes = hybridThresholdBytes;
        Directory = directory;
    }

    /// <summary>
    /// 获取异步暂存策略。
    /// </summary>
    internal NpoiAsyncStagingStrategy Strategy { get; }

    /// <summary>
    /// 获取内存与文件混合暂存阈值（字节）。
    /// </summary>
    internal long HybridThresholdBytes { get; }

    /// <summary>
    /// 获取文件暂存目录。
    /// </summary>
    internal string Directory { get; }

    /// <inheritdoc />
    public INpoiAsyncStaging Create(string prefix) => new NpoiAsyncStaging(
        Strategy, HybridThresholdBytes, prefix, Directory);
}
