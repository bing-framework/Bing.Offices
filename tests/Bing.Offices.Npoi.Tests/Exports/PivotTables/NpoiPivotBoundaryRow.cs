namespace Bing.Offices.Tests;

/// <summary>
/// 提供流式选择边界的两列透视来源。
/// </summary>
public sealed class NpoiPivotBoundaryRow
{
    /// <summary>
    /// 获取或设置仓库分组。
    /// </summary>
    public string Region { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public double Amount { get; set; }
}
