using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Exports.Attributes;
using Bing.Offices.Internals;
using System.Reflection;

namespace Bing.Offices.Settings;

/// <summary>
/// 导出列属性设置
/// </summary>
public sealed class ExportColumnPropertySetting
{
    /// <summary>
    /// 初始化一个<see cref="ExportColumnPropertySetting"/>类型的实例
    /// </summary>
    /// <param name="property">属性信息</param>
    public ExportColumnPropertySetting(PropertyInfo property)
    {
        PropertyInfo = property;
    }

    /// <summary>
    /// 列宽
    /// </summary>
    public int ColumnWidth { get; set; }

    /// <summary>
    /// 列索引
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 表头样式
    /// </summary>
    public IBaseStyle HeaderStyle { get; set; }

    /// <summary>
    /// 内容样式
    /// </summary>
    public IBaseStyle ColumnStyle { get; set; }

    /// <summary>
    /// 字符串格式化
    /// </summary>
    public string StringFormat { get; set; }

    /// <summary>
    /// 单元行合并
    /// </summary>
    public bool RowMerged { get; set; }

    /// <summary>
    /// 是否忽略属性
    /// </summary>
    public bool Ignored { get; set; }

    /// <summary>
    /// 默认值
    /// </summary>
    public object DefaultValue { get; set; }

    /// <summary>
    /// 属性信息
    /// </summary>
    public PropertyInfo PropertyInfo { get; set; }

    /// <summary>
    /// 通过特性初始化信息
    /// </summary>
    public void InitByAttribute()
    {
        Name = PropertyInfo.GetCustomAttribute(typeof(ColumnNameAttribute)) is ColumnNameAttribute columnNameAttr
            ? columnNameAttr.Name
            : PropertyInfo.Name;
        if (PropertyInfo.GetCustomAttribute(typeof(HeaderStyleAttribute)) is HeaderStyleAttribute headerStyleAttr)
            HeaderStyle = headerStyleAttr.Style;
        if (PropertyInfo.GetCustomAttribute(typeof(ColumnStyleAttribute)) is ColumnStyleAttribute columnStyleAttr)
            ColumnStyle = columnStyleAttr.Style;
        if (PropertyInfo.GetCustomAttribute(typeof(StringFormatterAttribute)) is StringFormatterAttribute formatterAttr)
            StringFormat = formatterAttr.Format;
        if (PropertyInfo.GetCustomAttribute(typeof(RowMergedAttribute)) is RowMergedAttribute)
            RowMerged = true;
        if (InternalHelper.HasIgnore(PropertyInfo))
            Ignored = true;
    }
}
