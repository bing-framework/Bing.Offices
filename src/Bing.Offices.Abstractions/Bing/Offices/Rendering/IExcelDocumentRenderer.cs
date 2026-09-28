namespace Bing.Offices.Rendering;

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
