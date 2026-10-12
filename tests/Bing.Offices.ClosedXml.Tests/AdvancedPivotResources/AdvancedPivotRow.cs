namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;

/// <summary>
/// 提供高级透视资源测试使用的来源行。
/// </summary>
internal sealed class AdvancedPivotRow
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotRow" /> 类型的实例。
    /// </summary>
    public AdvancedPivotRow() { }

    /// <summary>
    /// 获取或设置区域名称。
    /// </summary>
    public string Region { get; set; }

    /// <summary>
    /// 获取或设置状态名称。
    /// </summary>
    public string Status { get; set; }

    /// <summary>
    /// 获取或设置数值金额。
    /// </summary>
    public decimal Amount { get; set; }
}
