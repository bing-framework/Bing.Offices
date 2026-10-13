namespace Bing.Offices.Tools.PivotClientProbe;

/// <summary>
/// 表示透视客户端验收的仓库金额记录。
/// </summary>
internal sealed class PivotProbeRow
{
    /// <summary>
    /// 初始化一个 <see cref="PivotProbeRow"/> 类型的实例。
    /// </summary>
    public PivotProbeRow()
    {
    }

    /// <summary>
    /// 获取或设置仓库名称。
    /// </summary>
    public string Warehouse { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public double Amount { get; set; }
}
