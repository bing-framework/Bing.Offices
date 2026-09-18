using System.Collections.Generic;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 导入商品
/// </summary>
public class ImportGoods
{
    /// <summary>
    /// 获取或设置未预定义的动态字段集合。
    /// </summary>
    [DynamicColumn]
    public IDictionary<string, object> Extend { set; get; }

    #region 商品

    /// <summary>
    /// 获取或设置商品编码。
    /// </summary>
    [ColumnName("商品编码")]
    public string GoodsCode { get; set; }

    /// <summary>
    /// 获取或设置商品名称。
    /// </summary>
    [ColumnName("商品名称")]
    public string GoodsName { get; set; }

    /// <summary>
    /// 获取或设置末级分类编码。
    /// </summary>
    [ColumnName("末级分类编码")]
    public string CategoryCode { get; set; }

    /// <summary>
    /// 获取或设置基本单位。
    /// </summary>
    [ColumnName("基本单位")]
    public string Unit { get; set; }

    /// <summary>
    /// 获取或设置副单位 1。
    /// </summary>
    [ColumnName("副单位1")]
    public string Unit1 { get; set; }

    /// <summary>
    /// 获取或设置副单位 1 的箱规。
    /// </summary>
    [ColumnName("箱规1")]
    public int? Multiple1 { get; set; }

    /// <summary>
    /// 获取或设置副单位 2。
    /// </summary>
    [ColumnName("副单位2")]
    public string Unit2 { get; set; }

    /// <summary>
    /// 获取或设置副单位 2 的箱规。
    /// </summary>
    [ColumnName("箱规2")]
    public int? Multiple2 { get; set; }

    /// <summary>
    /// 获取或设置品牌名称。
    /// </summary>
    [ColumnName("品牌")]
    public string BrandName { set; get; }

    /// <summary>
    /// 获取或设置保质期天数。
    /// </summary>
    [ColumnName("保质期天数")]
    public int? ValidDays { get; set; }

    /// <summary>
    /// 获取或设置预警天数。
    /// </summary>
    [ColumnName("预警天数")]
    public int? WarningDays { get; set; }


    #endregion

    #region 产品

    /// <summary>
    /// 获取或设置 SKU 编码。
    /// </summary>
    [ColumnName("SKU编码")]
    public string ProductCode { get; set; }

    /// <summary>
    /// 获取或设置商品条码。
    /// </summary>
    [ColumnName("商品条码")]
    public string Barcode { get; set; }

    /// <summary>
    /// 获取或设置国家标准码。
    /// </summary>
    [ColumnName("国际码")]
    public string NationalCode { get; set; }

    /// <summary>
    /// 获取或设置 SKU 单位。
    /// </summary>
    [ColumnName("SKU单位")]
    public string ProductUnit { get; set; }

    /// <summary>
    /// 获取或设置建议零售价。
    /// </summary>
    [ColumnName("建议零售价")]
    public decimal? SuggestedRetailPrice { get; set; }

    /// <summary>
    /// 获取或设置参考采购价。
    /// </summary>
    [ColumnName("参考采购价")]
    public decimal? ReferPurchasePrice { get; set; }

    /// <summary>
    /// 获取或设置重量，单位为千克。
    /// </summary>
    [ColumnName("重量(kg)")]
    public decimal? Weight { get; set; }


    #region 产品属性分类

    /// <summary>
    /// 获取或设置属性分类 1。
    /// </summary>
    [ColumnName("属性分类1")]
    public string ProductAttribute1 { get; set; }

    /// <summary>
    /// 获取或设置属性值 1。
    /// </summary>
    [ColumnName("属性值1")]
    public string ProductAttributeValue1 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 2。
    /// </summary>
    [ColumnName("属性分类2")]
    public string ProductAttribute2 { get; set; }

    /// <summary>
    /// 获取或设置属性值 2。
    /// </summary>
    [ColumnName("属性值2")]
    public string ProductAttributeValue2 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 3。
    /// </summary>
    [ColumnName("属性分类3")]
    public string ProductAttribute3 { get; set; }

    /// <summary>
    /// 获取或设置属性值 3。
    /// </summary>
    [ColumnName("属性值3")]
    public string ProductAttributeValue3 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 4。
    /// </summary>
    [ColumnName("属性分类4")]
    public string ProductAttribute4 { get; set; }

    /// <summary>
    /// 获取或设置属性值 4。
    /// </summary>
    [ColumnName("属性值4")]
    public string ProductAttributeValue4 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 5。
    /// </summary>
    [ColumnName("属性分类5")]
    public string ProductAttribute5 { get; set; }

    /// <summary>
    /// 获取或设置属性值 5。
    /// </summary>
    [ColumnName("属性值5")]
    public string ProductAttributeValue5 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 6。
    /// </summary>
    [ColumnName("属性分类6")]
    public string ProductAttribute6 { get; set; }

    /// <summary>
    /// 获取或设置属性值 6。
    /// </summary>
    [ColumnName("属性值6")]
    public string ProductAttributeValue6 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 7。
    /// </summary>
    [ColumnName("属性分类7")]
    public string ProductAttribute7 { get; set; }

    /// <summary>
    /// 获取或设置属性值 7。
    /// </summary>
    [ColumnName("属性值7")]
    public string ProductAttributeValue7 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 8。
    /// </summary>
    [ColumnName("属性分类8")]
    public string ProductAttribute8 { get; set; }

    /// <summary>
    /// 获取或设置属性值 8。
    /// </summary>
    [ColumnName("属性值8")]
    public string ProductAttributeValue8 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 9。
    /// </summary>
    [ColumnName("属性分类9")]
    public string ProductAttribute9 { get; set; }

    /// <summary>
    /// 获取或设置属性值 9。
    /// </summary>
    [ColumnName("属性值9")]
    public string ProductAttributeValue9 { get; set; }

    /// <summary>
    /// 获取或设置属性分类 10。
    /// </summary>
    [ColumnName("属性分类10")]
    public string ProductAttribute10 { get; set; }

    /// <summary>
    /// 获取或设置属性值 10。
    /// </summary>
    [ColumnName("属性值10")]
    public string ProductAttributeValue10 { get; set; }

    #endregion
    #endregion
}
