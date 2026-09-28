using System.IO;

namespace Bing.Offices.IO;

/// <summary>
/// Excel 异步外围 IO 的 staging 策略。
/// </summary>
internal enum NpoiAsyncStagingStrategy
{
    /// <summary>
    /// 将 staging 内容保存在内存流中。
    /// </summary>
    Memory,
    /// <summary>
    /// 将 staging 内容保存在临时文件中。
    /// </summary>
    TempFile,
    /// <summary>
    /// 以内存为起点，超过阈值后迁移到临时文件。
    /// </summary>
    Hybrid
}
