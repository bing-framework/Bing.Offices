using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Bugs
{
    public class Issue2
    {
        /// <summary>
        /// 不确定字段
        /// </summary>
        [DynamicColumn]
        public System.Collections.Generic.IDictionary<string, object> Extend { set; get; }

        /// <summary>
        /// 商品条码
        /// </summary>
        [ColumnName("商品条码")]
        [Required(ErrorMsg = "商品条码为必填项")]
        public string Barcode { get; set; }

        /// <summary>
        /// 调拨数量
        /// </summary>
        [ColumnName("调拨数量")]
        [Required(ErrorMsg = "调拨数量为必填项")]
        [Range(1, 9999999, ErrorMsg = "调拨数量值不在允许[1~9999999]范围")]
        public int? Qty { get; set; }

        /// <summary>
        /// 商品编号
        /// </summary>
        [ColumnName("商品编号")]
        public string Code { get; set; }

        /// <summary>
        /// 商品名称
        /// </summary>
        [ColumnName("商品名称")]
        public string Name { get; set; }
    }
}