using System.IO;

namespace Bing.Offices.IO;

/// <summary>
/// 可替换的 Excel 异步 staging 工厂，供职责测试和资源探针使用。
/// </summary>
internal interface INpoiAsyncStagingFactory
{
    /// <summary>
    /// 创建一个新的异步 staging 会话。
    /// </summary>
    /// <param name="prefix">临时文件名使用的前缀。</param>
    /// <returns>新建的 staging 会话。</returns>
    INpoiAsyncStaging Create(string prefix);
}
