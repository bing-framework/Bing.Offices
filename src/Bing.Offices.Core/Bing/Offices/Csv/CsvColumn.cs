using Bing.Offices.Providers;

namespace Bing.Offices.Csv;

/// <summary>
/// 表示一个 CSV 源列及其固定或动态实体属性绑定。
/// </summary>
internal sealed class CsvColumn
{
    /// <summary>
    /// 初始化一个 <see cref="CsvColumn" /> 类型的实例。
    /// </summary>
    /// <param name="index">源 CSV 中的零基列索引。</param>
    /// <param name="property">固定或动态目标属性绑定。</param>
    /// <param name="headerName">源 CSV 表头名称。</param>
    /// <param name="isDynamic">是否将字段写入动态值字典。</param>
    /// <param name="dynamicColumn">与表头匹配的动态列计划。</param>
    public CsvColumn(int index, CsvPropertyBinding property, string headerName, bool isDynamic,
        IExcelDynamicMappingColumn dynamicColumn = null)
    {
        Index = index;
        Property = property;
        HeaderName = headerName;
        IsDynamic = isDynamic;
        DynamicColumn = dynamicColumn;
    }

    /// <summary>
    /// 获取源 CSV 中的零基列索引。
    /// </summary>
    public int Index { get; }
    /// <summary>
    /// 获取目标实体属性绑定。
    /// </summary>
    public CsvPropertyBinding Property { get; }
    /// <summary>
    /// 获取源 CSV 表头名称。
    /// </summary>
    public string HeaderName { get; }
    /// <summary>
    /// 获取是否将字段写入动态值字典。
    /// </summary>
    public bool IsDynamic { get; }
    /// <summary>
    /// 获取与表头匹配的动态列计划；固定列时为 null。
    /// </summary>
    public IExcelDynamicMappingColumn DynamicColumn { get; }
}
