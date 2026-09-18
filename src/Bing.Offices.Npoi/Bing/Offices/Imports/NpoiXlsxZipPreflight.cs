using Bing.Offices.IO;

namespace Bing.Offices.Imports;

/// <summary>
/// NPOI 的 XLSX ZIP 预检适配器。
/// </summary>
/// <remarks>
/// 预检逻辑由项目编译链接的共享 XLSX ZIP/XML 实现提供。
/// </remarks>
internal static class NpoiXlsxZipPreflight
{
    /// <summary>
    /// 按 NPOI 导入流程的资源限制预检工作簿流。
    /// </summary>
    /// <param name="source">待检查的可定位工作簿流。</param>
    /// <param name="limits">ZIP/XML 资源限制；为 null 时跳过限制检查。</param>
    /// <param name="cancellationToken">用于取消预检的令牌。</param>
    /// <remarks>
    /// 非 ZIP 流由 NPOI 继续处理；启用资源限制且输入为 ZIP 时执行共享的条目、大小和 XML 安全检查。
    /// </remarks>
    internal static void Validate(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken = default) =>
        ExcelXlsxZipPreflight.Validate(source, limits, "NPOI", requireZip: false, cancellationToken);
}
