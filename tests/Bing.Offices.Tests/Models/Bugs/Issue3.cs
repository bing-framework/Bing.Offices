using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Bugs;

/// <summary>
/// 表示 Issue3 场景使用的 Excel 测试模型。
/// </summary>
public class Issue3
{
    /// <summary>
    /// 获取或设置商品条码。
    /// </summary>
    [ColumnName("商品条码")]
    [ExcelRequired(ErrorMsg = "商品条码为必填项")]
    public string Barcode { get; set; }

    /// <summary>
    /// 获取或设置调拨数量。
    /// </summary>
    [ColumnName("调拨数量")]
    [ExcelRequired(ErrorMsg = "调拨数量为必填项")]
    [ExcelRange(1, 9999999, ErrorMsg = "调拨数量值不在允许[1~9999999]范围")]
    public int? Qty { get; set; }
}
