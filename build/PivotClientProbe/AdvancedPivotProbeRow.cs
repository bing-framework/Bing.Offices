namespace Bing.Offices.Tools.PivotClientProbe;

/// <summary>
/// 表示高级透视客户端验收的采购记录。
/// </summary>
internal sealed class AdvancedPivotProbeRow
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotProbeRow"/> 类型的实例。
    /// </summary>
    public AdvancedPivotProbeRow()
    {
    }

    /// <summary>
    /// 获取或设置仓库。
    /// </summary>
    public string Warehouse { get; set; }

    /// <summary>
    /// 获取或设置商品分类。
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
    /// 获取或设置报表纳入标记。
    /// </summary>
    public string Inclusion { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public double Amount { get; set; }

    /// <summary>
    /// 获取或设置商品编号。
    /// </summary>
    public string Item { get; set; }
}
