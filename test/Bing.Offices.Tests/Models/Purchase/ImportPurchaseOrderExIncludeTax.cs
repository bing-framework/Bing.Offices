using System;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Purchase
{

    /// <summary>
    /// 导入不含税价
    /// </summary>
    public class ImportPurchaseOrderExIncludeTax
    {
        /// <summary>
        /// 仓库名称
        /// </summary>
        [ColumnName("仓库名称")]
        public string WarehouseName { get; set; }

        /// <summary>
        /// 供应商名称
        /// </summary>
        [ColumnName("供应商名称")]
        public string SupplierName { get; set; }

        /// <summary>
        /// 税率(%)
        /// </summary>
        [ColumnName("税率(%)")]
        public decimal? TaxRate { get; set; }

        /// <summary>
        /// 预计到货日期
        /// </summary>
        [ColumnName("预计到货日期")]
        public DateTime? EstimatedArrivalDate { get; set; }

        /// <summary>
        /// 商品条码
        /// </summary>
        [ColumnName("商品条码")]
        public string Barcode { get; set; }

        /// <summary>
        /// 采购数量
        /// </summary>
        [ColumnName("采购数量")]
        public int? Qty { get; set; }

        /// <summary>
        /// 单价
        /// </summary>
        [ColumnName("不含税单价")]
        public decimal? Price4ExIncludeTax { get; set; }
    }
}