using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Exports;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 描述一个 ClosedXML 导出物理列及其映射来源。
/// </summary>
internal sealed class ClosedXmlExportColumn
{
    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExportColumn" /> 类型的实例。
    /// </summary>
    /// <param name="key">列键。</param>
    /// <param name="title">列标题。</param>
    /// <param name="fixedColumn">固定映射列。</param>
    /// <param name="dynamicColumn">动态映射列。</param>
    /// <param name="property">固定列对应的实体属性。</param>
    /// <param name="headerStyle">列级表头样式。</param>
    /// <param name="bodyStyle">列级正文样式。</param>
    /// <param name="mappingNumberFormat">映射数字格式。</param>
    /// <param name="numberFormat">列级数字格式。</param>
    /// <param name="physicalColumnIndex">初始物理列索引。</param>
    /// <param name="calculated">实体布局计算列；普通列时为空。</param>
    internal ClosedXmlExportColumn(string key, string title, IExcelMappingColumn fixedColumn,
        IExcelDynamicMappingColumn dynamicColumn, PropertyInfo property, ExcelCellStyle headerStyle,
        ExcelCellStyle bodyStyle, string mappingNumberFormat, string numberFormat,
        int? physicalColumnIndex, IExcelEntityCalculatedColumn calculated = null)
    {
        Key = key;
        Title = title;
        Fixed = fixedColumn;
        Dynamic = dynamicColumn;
        Property = property;
        HeaderStyle = headerStyle;
        BodyStyle = bodyStyle;
        MappingNumberFormat = mappingNumberFormat;
        NumberFormat = numberFormat;
        PhysicalColumnIndex = physicalColumnIndex ?? 0;
        Calculated = calculated;
    }

    /// <summary>
    /// 获取列键。
    /// </summary>
    internal string Key { get; }

    /// <summary>
    /// 获取列标题。
    /// </summary>
    internal string Title { get; }

    /// <summary>
    /// 获取固定映射列；动态列时返回 <see langword="null" />。
    /// </summary>
    internal IExcelMappingColumn Fixed { get; }

    /// <summary>
    /// 获取动态映射列；固定列时返回 <see langword="null" />。
    /// </summary>
    internal IExcelDynamicMappingColumn Dynamic { get; }

    /// <summary>
    /// 获取实体布局计算列；普通固定列和动态列时返回 <see langword="null" />。
    /// </summary>
    internal IExcelEntityCalculatedColumn Calculated { get; }

    /// <summary>
    /// 获取固定列对应的实体属性。
    /// </summary>
    internal PropertyInfo Property { get; }

    /// <summary>
    /// 获取列级表头样式。
    /// </summary>
    internal ExcelCellStyle HeaderStyle { get; }

    /// <summary>
    /// 获取列级正文样式。
    /// </summary>
    internal ExcelCellStyle BodyStyle { get; }

    /// <summary>
    /// 获取映射数字格式。
    /// </summary>
    internal string MappingNumberFormat { get; }

    /// <summary>
    /// 获取列级数字格式。
    /// </summary>
    internal string NumberFormat { get; }

    /// <summary>
    /// 获取或设置零基物理列索引。
    /// </summary>
    internal int PhysicalColumnIndex { get; set; }
}
