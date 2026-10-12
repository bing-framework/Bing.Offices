namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 高级透视数字格式测试的来源行。
/// </summary>
internal sealed class AdvancedPivotNumberFormatRow
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotNumberFormatRow"/> 类型的实例。
    /// </summary>
    public AdvancedPivotNumberFormatRow() { }

    /// <summary>
    /// 获取或设置类别名称。
    /// </summary>
    public string Category { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public double Amount { get; set; }

    /// <summary>
    /// 获取或设置订单编号。
    /// </summary>
    public string OrderId { get; set; }
}
