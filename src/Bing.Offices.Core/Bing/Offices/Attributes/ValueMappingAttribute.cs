namespace Bing.Offices.Attributes;

/// <summary>
/// 声明 Excel 显示文本与实体属性值之间的映射关系。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class ValueMappingAttribute : Attribute
{
    /// <summary>
    /// 初始化一个 <see cref="ValueMappingAttribute" /> 类型的实例。
    /// </summary>
    /// <param name="text">Excel 单元格中使用的显示文本。</param>
    /// <param name="value">与显示文本对应的实体属性值。</param>
    public ValueMappingAttribute(string text, object value)
    {
        Text = text;
        Value = value;
    }

    /// <summary>
    /// 获取 Excel 单元格中使用的显示文本。
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// 获取与显示文本对应的实体属性值。
    /// </summary>
    public object Value { get; }
}
