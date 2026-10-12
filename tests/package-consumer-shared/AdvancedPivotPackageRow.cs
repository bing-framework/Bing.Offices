namespace Bing.Offices.PackageConsumer;

/// <summary>
/// 高级透视包消费者合同使用的业务来源行。
/// </summary>
internal sealed class AdvancedPivotPackageRow
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotPackageRow"/> 类型的实例。
    /// </summary>
    public AdvancedPivotPackageRow() { }

    /// <summary>
    /// 获取或设置仓库名称。
    /// </summary>
    public string Warehouse { get; set; }

    /// <summary>
    /// 获取或设置商品类别。
    /// </summary>
    public string Category { get; set; }

    /// <summary>
    /// 获取或设置月份。
    /// </summary>
    public string Month { get; set; }

    /// <summary>
    /// 获取或设置采购状态。
    /// </summary>
    public string Status { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public decimal Amount { get; set; }

    /// <summary>
    /// 获取或设置文本商品编号。
    /// </summary>
    public string Item { get; set; }
}
