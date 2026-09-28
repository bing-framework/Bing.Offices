using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Bing.Offices.Exports;
using Bing.Offices.Styles;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体列表计算列的配置构建器。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityCalculatedColumnBuilder
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCalculatedColumnBuilder" /> 类型的实例。
    /// </summary>
    public ExcelEntityCalculatedColumnBuilder()
    {
    }

    /// <summary>
    /// 保存计算列的排序值。
    /// </summary>
    private int _order;

    /// <summary>
    /// 保存计算列的位置配置。
    /// </summary>
    private ExcelColumnPlacement _placement;

    /// <summary>
    /// 保存计算列表头样式。
    /// </summary>
    private ExcelCellStyle _headerStyle;

    /// <summary>
    /// 保存计算列正文样式。
    /// </summary>
    private ExcelCellStyle _bodyStyle;

    /// <summary>
    /// 保存计算列数字格式。
    /// </summary>
    private string _numberFormat;

    /// <summary>
    /// 设置计算列的默认排序值。
    /// </summary>
    /// <param name="order">排序值，值越小越靠前。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder Order(int order)
    {
        _order = order;
        return this;
    }

    /// <summary>
    /// 设置计算列的相对位置或物理列位置。
    /// </summary>
    /// <param name="placement">列位置。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder Placement(ExcelColumnPlacement placement)
    {
        _placement = placement ?? throw new ArgumentNullException(nameof(placement));
        return this;
    }

    /// <summary>
    /// 设置计算列表头样式。
    /// </summary>
    /// <param name="style">表头样式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder HeaderStyle(ExcelCellStyle style)
    {
        _headerStyle = style;
        return this;
    }

    /// <summary>
    /// 设置计算列正文样式。
    /// </summary>
    /// <param name="style">正文样式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder BodyStyle(ExcelCellStyle style)
    {
        _bodyStyle = style;
        return this;
    }

    /// <summary>
    /// 设置计算列数字格式。
    /// </summary>
    /// <param name="numberFormat">数字格式字符串。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder NumberFormat(string numberFormat)
    {
        _numberFormat = numberFormat;
        return this;
    }

    /// <summary>
    /// 获取计算列排序值。
    /// </summary>
    internal int OrderValue => _order;

    /// <summary>
    /// 获取计算列相对位置。
    /// </summary>
    internal ExcelColumnPlacement PlacementValue => _placement;

    /// <summary>
    /// 获取计算列表头样式。
    /// </summary>
    internal ExcelCellStyle HeaderStyleValue => _headerStyle;

    /// <summary>
    /// 获取计算列正文样式。
    /// </summary>
    internal ExcelCellStyle BodyStyleValue => _bodyStyle;

    /// <summary>
    /// 获取计算列数字格式。
    /// </summary>
    internal string NumberFormatValue => _numberFormat;
}
