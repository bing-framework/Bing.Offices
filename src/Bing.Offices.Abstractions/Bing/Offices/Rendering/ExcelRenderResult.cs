namespace Bing.Offices.Rendering;

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
