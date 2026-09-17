using Bing.Offices.IO;

namespace Bing.Offices.Imports;

/// <summary>
/// NPOI 的 XLSX ZIP 预检适配器，实际限制逻辑由 Core 共享实现提供。
/// </summary>
internal static class NpoiXlsxZipPreflight
{
    internal static void Validate(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken = default) =>
        ExcelXlsxZipPreflight.Validate(source, limits, "NPOI", requireZip: false, cancellationToken);
}
