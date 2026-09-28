namespace Bing.Offices.Rendering;

/// <summary>
/// 工作簿渲染请求。
/// </summary>
public sealed class ExcelRenderRequest
{
    /// <summary>
    /// 获取或初始化输入工作簿格式。
    /// </summary>
    public ExcelFormat InputFormat { get; init; } = ExcelFormat.Xlsx;
    /// <summary>
    /// 获取或初始化目标工作表名称。
    /// </summary>
    /// <remarks>为空时渲染整本工作簿。</remarks>
    public string SheetName { get; init; }
    /// <summary>
    /// 获取或初始化起始页（从 1 开始）。
    /// </summary>
    public int? StartPage { get; init; }
    /// <summary>
    /// 获取或初始化结束页（从 1 开始）。
    /// </summary>
    public int? EndPage { get; init; }
    /// <summary>
    /// 获取或初始化输出图片格式。
    /// </summary>
    public ExcelRenderFormat ImageFormat { get; init; } = ExcelRenderFormat.Png;
    /// <summary>
    /// 获取或初始化 PDF 兼容级别。
    /// </summary>
    /// <remarks>仅对 PDF 输出生效。</remarks>
    public ExcelPdfCompliance PdfCompliance { get; init; } = ExcelPdfCompliance.None;
    /// <summary>
    /// 获取或初始化图片 DPI。
    /// </summary>
    public int Dpi { get; init; } = 150;
    /// <summary>
    /// 获取或初始化可选的 A1 区域。
    /// </summary>
    public string Area { get; init; }
    /// <summary>
    /// 获取或初始化是否以严格字体模式运行。
    /// </summary>
    public bool StrictFonts { get; init; }
}
