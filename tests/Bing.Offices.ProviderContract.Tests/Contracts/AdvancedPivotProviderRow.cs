namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 提供高级透视 Provider 拒绝测试的来源行。
/// </summary>
internal sealed class AdvancedPivotProviderRow
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotProviderRow" /> 类型的实例。
    /// </summary>
    public AdvancedPivotProviderRow() { }

    /// <summary>
    /// 获取或设置区域名称。
    /// </summary>
    public string Region { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public decimal Amount { get; set; }

    /// <summary>
    /// 获取或设置状态。
    /// </summary>
    public string Status { get; set; }

    /// <summary>
    /// 获取或设置销售渠道。
    /// </summary>
    public string Channel { get; set; }
}
