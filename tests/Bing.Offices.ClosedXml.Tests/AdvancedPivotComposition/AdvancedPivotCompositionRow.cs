namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotComposition;

/// <summary>
/// 提供高级透视组合测试使用的来源记录。
/// </summary>
internal sealed class AdvancedPivotCompositionRow
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotCompositionRow" /> 类型的实例。
    /// </summary>
    public AdvancedPivotCompositionRow() { }

    /// <summary>
    /// 获取或设置订单编号。
    /// </summary>
    public string OrderId { get; set; }

    /// <summary>
    /// 获取或设置仓库区域。
    /// </summary>
    public string Region { get; set; }

    /// <summary>
    /// 获取或设置订单状态。
    /// </summary>
    public string Status { get; set; }

    /// <summary>
    /// 获取或设置销售渠道。
    /// </summary>
    public string Channel { get; set; }

    /// <summary>
    /// 获取或设置订单金额。
    /// </summary>
    public decimal Amount { get; set; }
}
