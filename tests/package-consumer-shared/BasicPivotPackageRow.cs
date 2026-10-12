namespace Bing.Offices.PackageConsumer;

/// <summary>
/// 基础透视包消费者合同使用的最小来源行。
/// </summary>
internal sealed class BasicPivotPackageRow
{
    /// <summary>
    /// 初始化一个 <see cref="BasicPivotPackageRow" /> 类型的实例。
    /// </summary>
    public BasicPivotPackageRow()
    {
    }

    /// <summary>
    /// 获取或设置仓库名称。
    /// </summary>
    public string Warehouse { get; set; }

    /// <summary>
    /// 获取或设置采购金额。
    /// </summary>
    public decimal Amount { get; set; }

    /// <summary>
    /// 获取或设置商品编号。
    /// </summary>
    public string Item { get; set; }
}
