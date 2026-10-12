namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotComposition;

/// <summary>
/// 提供组合模板中的订单录入行。
/// </summary>
internal sealed class AdvancedPivotCompositionEntry
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotCompositionEntry" /> 类型的实例。
    /// </summary>
    public AdvancedPivotCompositionEntry() { }

    /// <summary>
    /// 获取或设置订单编号。
    /// </summary>
    public string OrderId { get; set; }

    /// <summary>
    /// 获取或设置订单状态。
    /// </summary>
    public string Status { get; set; }
}
