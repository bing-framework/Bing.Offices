using System.Collections.Generic;

using Bing.Offices.Attributes;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Business;

/// <summary>
/// 提供透视业务组合测试使用的来源记录。
/// </summary>
internal sealed class NpoiPivotBusinessRow
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotBusinessRow" /> 类型的实例。
    /// </summary>
    public NpoiPivotBusinessRow() { }

    /// <summary>
    /// 获取或设置业务订单编号。
    /// </summary>
    public string OrderNo { get; set; }

    /// <summary>
    /// 获取或设置订单金额。
    /// </summary>
    public double Amount { get; set; }

    /// <summary>
    /// 获取或设置订单状态。
    /// </summary>
    public string Status { get; set; }

    /// <summary>
    /// 获取或设置动态来源字段。
    /// </summary>
    [DynamicColumn]
    public IDictionary<string, object> DynamicValues { get; set; } = new Dictionary<string, object>();

}
