using Bing.Offices.Attributes;
using System;

namespace Bing.Offices.Tests.Models.Purchase;

/// <summary>
/// 导入采购单_不含税
/// </summary>
public class ImportPurchaseOrderIncludeTax
{
    /// <summary>
    /// 获取或设置仓库名称。
    /// </summary>
    [ColumnName("仓库名称")]
    public string WarehouseName { get; set; }

    /// <summary>
    /// 获取或设置供应商名称。
    /// </summary>
    [ColumnName("供应商名称")]
    public string SupplierName { get; set; }

    /// <summary>
    /// 获取或设置税率，单位为百分比。
    /// </summary>
    [ColumnName("税率(%)")]
    public decimal? TaxRate { get; set; }

    /// <summary>
    /// 获取或设置预计到货日期。
    /// </summary>
    [ColumnName("预计到货日期")]
    public DateTime? EstimatedArrivalDate { get; set; }

    /// <summary>
    /// 获取或设置商品条码。
    /// </summary>
    [ColumnName("商品条码")]
    public string Barcode { get; set; }

    /// <summary>
    /// 获取或设置采购数量。
    /// </summary>
    [ColumnName("采购数量")]
    public int? Qty { get; set; }

    /// <summary>
    /// 获取或设置含税单价。
    /// </summary>
    [ColumnName("含税单价")]
    public decimal? Price4IncludeTax { get; set; }
}
