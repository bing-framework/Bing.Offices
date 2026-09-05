using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.Csv;

/// <summary>保存 CSV 列与实体属性之间的反射、转换器和校验绑定。</summary>
internal sealed class CsvPropertyBinding
{
    /// <summary>创建 CSV 属性绑定。</summary>
    /// <param name="mapping">跨提供程序列映射。</param>
    /// <param name="property">目标实体属性。</param>
    /// <param name="getter">读取实体属性值的委托。</param>
    /// <param name="setter">写入实体属性值的委托。</param>
    /// <param name="attributes">目标属性上的特性快照。</param>
    private CsvPropertyBinding(IExcelMappingColumn mapping, PropertyInfo property,
        Func<object, object> getter, Action<object, object> setter, IReadOnlyList<Attribute> attributes)
    {
        Mapping = mapping;
        Property = property;
        Getter = getter;
        Setter = setter;
        Attributes = attributes;
        ValueConverters = mapping.ValueConverters;
        ValidationBindings = mapping.ValidationBindings;
        IsUnique = mapping.IsUnique;
        UniqueIgnoreEmpty = mapping.UniqueIgnoreEmpty;
    }

    /// <summary>获取跨提供程序列映射。</summary>
    internal IExcelMappingColumn Mapping { get; }
    /// <summary>获取目标实体属性元数据。</summary>
    internal PropertyInfo Property { get; }
    /// <summary>获取实体属性名称。</summary>
    internal string Name => Mapping.Name;
    /// <summary>获取 CSV 列标题。</summary>
    internal string Title => Mapping.Title;
    /// <summary>获取 CSV 标题别名。</summary>
    internal IReadOnlyList<string> Aliases => Mapping.Aliases;
    /// <summary>获取值格式化字符串。</summary>
    internal string Formatter => Mapping.Formatter;
    /// <summary>获取是否忽略该列。</summary>
    internal bool Ignored => Mapping.Ignored;
    /// <summary>获取是否为动态列容器。</summary>
    internal bool IsDynamicColumn => Mapping.IsDynamicColumn;
    /// <summary>获取列级导入空白策略。</summary>
    internal ExcelWhitespacePolicy? ImportWhitespace => Mapping.ImportWhitespace;
    /// <summary>获取小数精度。</summary>
    internal byte? DecimalScale => Mapping.DecimalScale;
    /// <summary>获取值转换器名称。</summary>
    internal string ConverterName => Mapping.ConverterName;
    /// <summary>获取命名校验规则名称。</summary>
    internal IReadOnlyList<string> ValidationRuleNames => Mapping.ValidationRuleNames;
    /// <summary>获取显示文本到属性值文本的映射。</summary>
    internal IReadOnlyDictionary<string, string> ValueMap => Mapping.ValueMap;
    /// <summary>获取读取实体属性的委托。</summary>
    internal Func<object, object> Getter { get; }
    /// <summary>获取写入实体属性的委托。</summary>
    internal Action<object, object> Setter { get; }
    /// <summary>获取属性特性快照。</summary>
    internal IReadOnlyList<Attribute> Attributes { get; }
    /// <summary>获取已绑定的值转换器。</summary>
    internal IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <summary>获取已绑定的校验规则。</summary>
    internal IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    /// <summary>获取是否执行唯一性校验。</summary>
    internal bool IsUnique { get; }
    /// <summary>获取唯一性校验是否忽略空值。</summary>
    internal bool UniqueIgnoreEmpty { get; }

    /// <summary>创建指定实体类型的 CSV 属性绑定。</summary>
    /// <typeparam name="T">目标实体类型。</typeparam>
    /// <param name="mapping">跨提供程序列映射。</param>
    /// <returns>绑定了属性访问器、转换器和校验规则的 CSV 列定义。</returns>
    internal static CsvPropertyBinding Create<T>(IExcelMappingColumn mapping) where T : class, new()
    {
        if (mapping == null)
            throw new ArgumentNullException(nameof(mapping));
        if (mapping is IExcelCompiledMappingColumn compiled)
            return new CsvPropertyBinding(mapping, compiled.Property, compiled.Getter, compiled.Setter,
                compiled.Attributes);
        var property = typeof(T).GetProperty(mapping.Name, BindingFlags.Instance | BindingFlags.Public);
        if (property == null)
            throw new InvalidOperationException($"无法解析映射属性: {mapping.Name}");
        return new CsvPropertyBinding(mapping, property, instance => property.GetValue(instance),
            (instance, value) => SetPropertyValue(property, instance, value),
            property.GetCustomAttributes<Attribute>().ToArray());
    }

    /// <summary>写入 CSV 实体属性并解包反射调用产生的目标异常。</summary>
    /// <param name="property">目标属性。</param>
    /// <param name="instance">目标实体。</param>
    /// <param name="value">待写入的值。</param>
    private static void SetPropertyValue(PropertyInfo property, object instance, object value)
    {
        try
        {
            property.SetValue(instance, value);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }
}
