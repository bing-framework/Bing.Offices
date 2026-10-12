namespace Bing.Offices.Tests;

/// <summary>
/// 提供原生透视公开导出的采购来源。
/// </summary>
public sealed class NpoiPivotExportRow
{
    /// <summary>
    /// 获取或设置区域。
    /// </summary>
    public string Region { get; set; }

    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public double Amount { get; set; }

    /// <summary>
    /// 获取或设置状态。
    /// </summary>
    public string State { get; set; }

    /// <summary>
    /// 获取或设置渠道。
    /// </summary>
    public string Channel { get; set; }

    /// <summary>
    /// 获取或设置商品编码。
    /// </summary>
    public string Code { get; set; }
}
