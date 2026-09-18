using Bing.Offices.IO;
using Bing.Offices.Imports;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// MiniExcel 的 XLSX ZIP 预检适配器。
/// </summary>
/// <remarks>
/// 预检逻辑由项目编译链接的共享 XLSX ZIP/XML 实现提供。
/// </remarks>
internal static class MiniExcelXlsxPreflight
{
    /// <summary>
    /// 按 MiniExcel 的输入要求预检工作簿流。
    /// </summary>
    /// <param name="source">待检查的可定位工作簿流。</param>
    /// <param name="limits">ZIP/XML 资源限制；为 null 时跳过限制检查。</param>
    /// <param name="cancellationToken">用于取消预检的令牌。</param>
    /// <remarks>
    /// MiniExcel 适配器要求输入为 XLSX ZIP；进入预检后，可定位流会在结束时置于文件头。
    /// </remarks>
    internal static void Validate(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken = default) =>
        ExcelXlsxZipPreflight.Validate(source, limits, "MiniExcel", requireZip: true, cancellationToken);

    /// <summary>
    /// 读取当前工作簿是否采用 Excel 1904 日期系统。
    /// </summary>
    /// <param name="source">待读取的可定位 XLSX 流；不可定位或长度不足时返回 false。</param>
    /// <param name="cancellationToken">读取过程使用的取消令牌。</param>
    /// <returns>date1904 值为 1、true 或 on 时返回 true；缺失或为其他值时返回 false。</returns>
    /// <remarks>
    /// 读取完成后恢复调用前的流位置。
    /// </remarks>
    internal static bool GetDate1904(Stream source, CancellationToken cancellationToken = default) =>
        ExcelXlsxZipPreflight.GetDate1904(source, cancellationToken);
}
