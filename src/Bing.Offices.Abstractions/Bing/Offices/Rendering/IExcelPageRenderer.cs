namespace Bing.Offices.Rendering;

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
