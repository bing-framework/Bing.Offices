namespace Bing.Offices.Mappings;

/// <summary>
/// Excel 类型的不可变静态映射。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
internal sealed class ExcelTypeMap<T>
{
    /// <summary>
    /// 初始化一个<see cref="ExcelTypeMap{T}"/>类型的实例。
    /// </summary>
    /// <param name="properties">属性映射集合。</param>
    internal ExcelTypeMap(IReadOnlyList<ExcelPropertyMap> properties) => Properties = properties;

    /// <summary>
    /// 获取按声明顺序排列的属性映射。
    /// </summary>
    public IReadOnlyList<ExcelPropertyMap> Properties { get; }
}
