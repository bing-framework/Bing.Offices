namespace Bing.Offices.Rendering;

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
