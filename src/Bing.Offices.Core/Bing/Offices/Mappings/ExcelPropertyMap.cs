using System.Reflection;
using Bing.Offices.Imports;

namespace Bing.Offices.Mappings;

/// <summary>
/// Excel 属性的不可变静态映射。
/// </summary>
public sealed class ExcelPropertyMap
{
    /// <summary>
    /// 初始化一个<see cref="ExcelPropertyMap"/>类型的实例。
    /// </summary>
    /// <param name="property">属性元数据。</param>
    /// <param name="title">默认列标题。</param>
    /// <param name="formatter">默认格式化字符串。</param>
    /// <param name="ignored">是否忽略。</param>
    /// <param name="isDynamicColumn">是否为动态列。</param>
    /// <param name="decimalScale">小数精度。</param>
    /// <param name="converterName">值转换器名称。</param>
    /// <param name="validationRuleNames">校验规则名称集合。</param>
    /// <param name="valueMap">显示文本到原始值的映射。</param>
    /// <param name="getter">已编译的属性读取器。</param>
    /// <param name="setter">已编译的属性写入器。</param>
    internal ExcelPropertyMap(PropertyInfo property, string title, string formatter, bool ignored, bool isDynamicColumn,
        byte? decimalScale, string converterName, IReadOnlyList<string> validationRuleNames,
        IReadOnlyDictionary<string, object> valueMap, Func<object, object> getter, Action<object, object> setter,
        ExcelImageMultiplicityPolicy imageMultiplicity = ExcelImageMultiplicityPolicy.First)
    {
        Property = property;
        Title = title;
        Formatter = formatter;
        Ignored = ignored;
        IsDynamicColumn = isDynamicColumn;
        DecimalScale = decimalScale;
        ConverterName = converterName;
        ValidationRuleNames = validationRuleNames;
        ValueMap = valueMap;
        Getter = getter;
        Setter = setter;
        ImageMultiplicity = imageMultiplicity;
    }

    /// <summary>
    /// 获取属性元数据。
    /// </summary>
    public PropertyInfo Property { get; }

    /// <summary>
    /// 获取属性名称。
    /// </summary>
    public string Name => Property.Name;

    /// <summary>
    /// 获取默认列标题。
    /// </summary>
    public string Title { get; }

    /// <summary>
    /// 获取默认格式化字符串。
    /// </summary>
    public string Formatter { get; }

    /// <summary>
    /// 获取是否忽略该属性。
    /// </summary>
    public bool Ignored { get; }

    /// <summary>
    /// 获取是否为动态列。
    /// </summary>
    public bool IsDynamicColumn { get; }

    /// <summary>
    /// 获取小数精度。
    /// </summary>
    public byte? DecimalScale { get; }

    /// <summary>
    /// 获取已注册值转换器的配置名称。
    /// </summary>
    public string ConverterName { get; }

    /// <summary>
    /// 获取已注册校验规则的配置名称集合。
    /// </summary>
    public IReadOnlyList<string> ValidationRuleNames { get; }

    /// <summary>
    /// 获取显示文本到原始值的映射。
    /// </summary>
    public IReadOnlyDictionary<string, object> ValueMap { get; }

    /// <summary>
    /// 获取已编译的属性读取器。
    /// </summary>
    public Func<object, object> Getter { get; }

    /// <summary>
    /// 获取已编译的属性写入器。
    /// </summary>
    public Action<object, object> Setter { get; }

    /// <summary>
    /// 获取图片列多重性策略。
    /// </summary>
    public ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
}
