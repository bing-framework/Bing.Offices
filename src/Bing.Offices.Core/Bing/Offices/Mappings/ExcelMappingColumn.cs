using System.Globalization;
using System.Reflection;
using System.Collections.ObjectModel;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Validations;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

internal sealed class ExcelMappingColumn : IExcelMappingColumn, IExcelCompiledMappingColumn
{
    /// <summary>从属性映射及已绑定规则创建不可变列计划。</summary>
    /// <param name="property">已解析的实体属性映射。</param>
    /// <param name="valueConverters">已绑定的值转换器。</param>
    /// <param name="validationBindings">已绑定的校验规则。</param>
    internal ExcelMappingColumn(ExcelPropertyMap property,
        IReadOnlyList<IExcelValueConverter> valueConverters,
        IReadOnlyList<IExcelValidationBinding> validationBindings)
    {
        if (property == null)
            throw new ArgumentNullException(nameof(property));
        Name = property.Name;
        Title = property.Title;
        Aliases = new ReadOnlyCollection<string>(property.Aliases.ToArray());
        Formatter = property.Formatter;
        Ignored = property.Ignored;
        IsDynamicColumn = property.IsDynamicColumn;
        ImportWhitespace = property.ImportWhitespace;
        DecimalScale = property.DecimalScale;
        ConverterName = property.ConverterName;
        ValidationRuleNames = new ReadOnlyCollection<string>(property.ValidationRuleNames.ToArray());
        ValueMap = new ReadOnlyDictionary<string, string>(property.ValueMap.ToDictionary(pair => pair.Key,
            pair => Convert.ToString(pair.Value, CultureInfo.InvariantCulture), StringComparer.Ordinal));
        ImageMultiplicity = property.ImageMultiplicity;
        ValueConverters = new ReadOnlyCollection<IExcelValueConverter>(valueConverters.ToArray());
        ValidationBindings = new ReadOnlyCollection<IExcelValidationBinding>(validationBindings.ToArray());
        IsUnique = ValidationBindings.Any(binding => binding.Kind == ExcelValidationBindingKind.Unique);
        UniqueIgnoreEmpty = property.Property.GetCustomAttributes<ExcelUniqueAttribute>().FirstOrDefault()?.IgnoreEmpty
            ?? true;
        Property = property.Property;
        Getter = property.Getter;
        Setter = property.Setter;
        Attributes = new ReadOnlyCollection<Attribute>(property.Property.GetCustomAttributes<Attribute>().ToArray());
    }

    /// <inheritdoc />
    public string Name { get; }
    /// <inheritdoc />
    public string Title { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> Aliases { get; }
    /// <inheritdoc />
    public string Formatter { get; }
    /// <inheritdoc />
    public bool Ignored { get; }
    /// <inheritdoc />
    public bool IsDynamicColumn { get; }
    /// <inheritdoc />
    public Bing.Offices.Imports.ExcelWhitespacePolicy? ImportWhitespace { get; }
    /// <inheritdoc />
    public byte? DecimalScale { get; }
    /// <inheritdoc />
    public string ConverterName { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> ValidationRuleNames { get; }
    /// <inheritdoc />
    public IReadOnlyDictionary<string, string> ValueMap { get; }
    /// <inheritdoc />
    public Bing.Offices.Imports.ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <inheritdoc />
    public bool IsUnique { get; }
    /// <inheritdoc />
    public bool UniqueIgnoreEmpty { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    /// <summary>获取映射到的实体属性元数据。</summary>
    public PropertyInfo Property { get; }
    /// <summary>获取从实体读取属性值的委托。</summary>
    public Func<object, object> Getter { get; }
    /// <summary>获取将转换后值写入实体属性的委托。</summary>
    public Action<object, object> Setter { get; }
    /// <summary>获取实体属性上的特性快照。</summary>
    public IReadOnlyList<Attribute> Attributes { get; }
}
