using Bing.Offices.IO;
using Bing.Offices.Imports;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// MiniExcel 的 XLSX ZIP 预检适配器，实际限制逻辑由 Core 共享实现提供。
/// </summary>
internal static class MiniExcelXlsxPreflight
{
    internal static void Validate(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken = default) =>
        ExcelXlsxZipPreflight.Validate(source, limits, "MiniExcel", requireZip: true, cancellationToken);

    /// <summary>读取当前工作簿是否采用 Excel 1904 日期系统。</summary>
    /// <param name="source">已缓冲且可定位的 XLSX 流。</param>
    /// <param name="cancellationToken">读取过程使用的取消令牌。</param>
    /// <returns>工作簿声明 1904 日期系统时返回 true。</returns>
    internal static bool GetDate1904(Stream source, CancellationToken cancellationToken = default) =>
        ExcelXlsxZipPreflight.GetDate1904(source, cancellationToken);
}
