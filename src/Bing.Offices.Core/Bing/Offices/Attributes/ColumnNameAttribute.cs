namespace Bing.Offices.Attributes;

/// <summary>
/// 指定实体属性对应的 Excel 列标题。
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ColumnNameAttribute : Attribute
{
    /// <summary>
    /// 获取或设置实体属性对应的 Excel 列标题。
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 初始化一个 <see cref="ColumnNameAttribute" /> 类型的实例。
    /// </summary>
    /// <param name="name">实体属性对应的 Excel 列标题。</param>
    public ColumnNameAttribute(string name) => Name = name;
}
