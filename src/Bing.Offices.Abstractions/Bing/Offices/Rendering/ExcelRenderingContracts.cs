namespace Bing.Offices.Rendering;

/// <summary>
/// 页面渲染输出格式。
/// </summary>
public enum ExcelRenderFormat
{
    /// <summary>
    /// PDF 文档。
    /// </summary>
    Pdf,
    /// <summary>
    /// PNG 页面图片。
    /// </summary>
    Png,
    /// <summary>
    /// JPEG 页面图片。
    /// </summary>
    Jpeg,
    /// <summary>
    /// TIFF 页面图片。
    /// </summary>
    Tiff
}

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

/// <summary>
/// 渲染警告。
/// </summary>
public sealed class ExcelRenderWarning
{
    /// <summary>
    /// 获取或初始化警告代码。
    /// </summary>
    public string Code { get; init; }
    /// <summary>
    /// 获取或初始化工作表名称。
    /// </summary>
    public string SheetName { get; init; }
    /// <summary>
    /// 获取或初始化警告消息。
    /// </summary>
    public string Message { get; init; }
}

/// <summary>
/// 文档渲染结果。
/// </summary>
public sealed class ExcelRenderResult
{
    /// <summary>
    /// 获取或初始化输出页数。
    /// </summary>
    public int PageCount { get; init; }
    /// <summary>
    /// 获取或初始化渲染警告。
    /// </summary>
    public IReadOnlyList<ExcelRenderWarning> Warnings { get; init; } = Array.Empty<ExcelRenderWarning>();
}

/// <summary>
/// Excel 工作簿到 PDF 的渲染能力。
/// </summary>
public interface IExcelDocumentRenderer
{
    /// <summary>
    /// 将工作簿渲染到调用方拥有的目标流。
    /// </summary>
    /// <param name="source">调用方拥有的输入工作簿流。</param>
    /// <param name="destination">调用方拥有的 PDF 输出流。</param>
    /// <param name="request">渲染范围、格式和字体选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含输出页数与渲染警告的结果。</returns>
    ExcelRenderResult Render(Stream source, Stream destination, ExcelRenderRequest request,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步渲染到调用方拥有的目标流。
    /// </summary>
    /// <param name="source">调用方拥有的输入工作簿流。</param>
    /// <param name="destination">调用方拥有的 PDF 输出流。</param>
    /// <param name="request">渲染范围、格式和字体选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果包含输出页数与渲染警告。</returns>
    Task<ExcelRenderResult> RenderAsync(Stream source, Stream destination, ExcelRenderRequest request,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// 按页输出图片的 Excel 渲染能力。
/// </summary>
public interface IExcelPageRenderer
{
    /// <summary>
    /// 将工作簿逐页渲染为图片。
    /// </summary>
    /// <remarks>页面回调串行执行，图片流仅在回调期间有效。</remarks>
    /// <param name="source">调用方拥有的输入工作簿流。</param>
    /// <param name="request">渲染范围、图片格式和字体选项。</param>
    /// <param name="pageSink">接收从 1 开始的输出页号、图片流及文件扩展名的回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含输出页数与渲染警告的结果。</returns>
    ExcelRenderResult RenderPages(Stream source, ExcelRenderRequest request,
        Action<int, Stream, string> pageSink, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步将工作簿逐页渲染为图片。
    /// </summary>
    /// <remarks>逐页等待回调完成，图片流仅在对应回调完成前有效。</remarks>
    /// <param name="source">调用方拥有的输入工作簿流。</param>
    /// <param name="request">渲染范围、图片格式和字体选项。</param>
    /// <param name="pageSink">接收从 1 开始的输出页号、图片流及文件扩展名的异步回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果包含输出页数与渲染警告。</returns>
    Task<ExcelRenderResult> RenderPagesAsync(Stream source, ExcelRenderRequest request,
        Func<int, Stream, string, Task> pageSink, CancellationToken cancellationToken = default);
}
