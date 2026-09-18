using System.ComponentModel.DataAnnotations.Schema;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 导出订单（合并单元格）
/// </summary>
public class ExportOrderWithMerged
{
    /// <summary>
    /// 获取或设置索引。
    /// </summary>
    [ColumnName("索引")]
    public string Index { get; set; }

    /// <summary>
    /// 获取或设置订单编码。
    /// </summary>
    [ColumnName("订单编码")]
    [MergeColumns]
    public string OrderNumber { get; set; }

    /// <summary>
    /// 获取或设置下单人。
    /// </summary>
    [ColumnName("下单人")]
    [MergeColumns]
    public string Buyer { get; set; }

    /// <summary>
    /// 获取或设置商品名称。
    /// </summary>
    [ColumnName("商品")]
    public string ProductName { get; set; }

    /// <summary>
    /// 获取或设置价格。
    /// </summary>
    [ColumnName("价格")]
    public double Price { get; set; }

    /// <summary>
    /// 获取或设置购买数量。
    /// </summary>
    [ColumnName("购买数量")]
    public int BuyQty { get; set; }

    /// <summary>
    /// 获取或设置应忽略的属性。
    /// </summary>
    [ExcelIgnore]
    [ColumnName("忽略属性")]
    public string IgnoreProperty { get; set; }

    /// <summary>
    /// 获取或设置不参与映射的属性。
    /// </summary>
    [NotMapped]
    [ColumnName("忽略映射属性")]
    public string NotMappedProperty { get; set; }
}
