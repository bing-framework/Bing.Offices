namespace Bing.Offices.Entities;

/// <summary>
/// 定义模板中本次未使用的逐页命名锚点处理方式。
/// </summary>
public enum ExcelEntityPageAnchorStaleNamePolicy
{
    /// <summary>
    /// 遇到本次未使用的旧名称时拒绝构建。
    /// </summary>
    Reject = 0,

    /// <summary>
    /// 删除当前系列中合法且已过期的旧名称。
    /// </summary>
    RemoveObsolete = 1
}
