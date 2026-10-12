using Bing.Offices.Exports;

namespace Bing.Offices.Internals;

/// <summary>
/// 已定位的透视模板更新。
/// </summary>
internal sealed class ExcelPivotTemplateUpdate
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateUpdate" /> 类型的实例。
    /// </summary>
    /// <param name="entry">已有透视描述。</param>
    /// <param name="definition">请求中的不可变更新定义。</param>
    internal ExcelPivotTemplateUpdate(ExcelPivotTemplateEntry entry, ExcelPivotTableDefinition definition)
    {
        Entry = entry;
        Definition = definition;
    }

    /// <summary>
    /// 获取已有透视描述。
    /// </summary>
    internal ExcelPivotTemplateEntry Entry { get; }

    /// <summary>
    /// 获取更新定义。
    /// </summary>
    internal ExcelPivotTableDefinition Definition { get; }
}
