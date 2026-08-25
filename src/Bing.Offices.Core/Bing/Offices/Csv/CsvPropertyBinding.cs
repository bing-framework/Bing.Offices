using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.Csv;

internal sealed class CsvPropertyBinding
{
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

    internal IExcelMappingColumn Mapping { get; }
    internal PropertyInfo Property { get; }
    internal string Name => Mapping.Name;
    internal string Title => Mapping.Title;
    internal IReadOnlyList<string> Aliases => Mapping.Aliases;
    internal string Formatter => Mapping.Formatter;
    internal bool Ignored => Mapping.Ignored;
    internal bool IsDynamicColumn => Mapping.IsDynamicColumn;
    internal ExcelWhitespacePolicy? ImportWhitespace => Mapping.ImportWhitespace;
    internal byte? DecimalScale => Mapping.DecimalScale;
    internal string ConverterName => Mapping.ConverterName;
    internal IReadOnlyList<string> ValidationRuleNames => Mapping.ValidationRuleNames;
    internal IReadOnlyDictionary<string, string> ValueMap => Mapping.ValueMap;
    internal Func<object, object> Getter { get; }
    internal Action<object, object> Setter { get; }
    internal IReadOnlyList<Attribute> Attributes { get; }
    internal IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    internal IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    internal bool IsUnique { get; }
    internal bool UniqueIgnoreEmpty { get; }

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
            (instance, value) => property.SetValue(instance, value),
            property.GetCustomAttributes<Attribute>().ToArray());
    }
}
