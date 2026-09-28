namespace Bing.Offices.Rendering;

/// <summary>
/// PDF 输出兼容级别。
/// </summary>
public enum ExcelPdfCompliance
{
    /// <summary>
    /// 默认 PDF 兼容级别。
    /// </summary>
    None,
    /// <summary>
    /// PDF/A-1b。
    /// </summary>
    PdfA1b,
    /// <summary>
    /// PDF/A-2b。
    /// </summary>
    PdfA2b,
    /// <summary>
    /// PDF/A-3b。
    /// </summary>
    PdfA3b
}
