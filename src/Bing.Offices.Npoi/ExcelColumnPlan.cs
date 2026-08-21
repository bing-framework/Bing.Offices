using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi;

/// <summary>
/// 导入和导出共享的不可变列计划基础元数据。
/// </summary>
internal sealed class ExcelColumnPlan
{
    internal ExcelColumnPlan(string headerName, ExcelPropertyMap property, bool isDynamic, int columnIndex,
        ExcelDynamicColumnDefinition dynamicDefinition, string key = null,
        IReadOnlyList<IExcelValueConverter> valueConverters = null,
        IReadOnlyList<ExcelAttributeValidationBinding> attributeValidations = null,
        IReadOnlyList<INamedExcelValidationRule> namedValidationRules = null)
    {
        HeaderName = headerName;
        Title = headerName;
        Property = property;
        IsDynamic = isDynamic;
        ColumnIndex = columnIndex;
        DynamicDefinition = dynamicDefinition;
        Key = key ?? dynamicDefinition?.Key ?? (isDynamic ? headerName : property.Name);
        Getter = property.Getter;
        Setter = property.Setter;
        ConverterName = dynamicDefinition?.ConverterName ?? property.ConverterName;
        ValidationRuleNames = property.ValidationRuleNames;
        ValidatorName = dynamicDefinition?.ValidatorName;
        ValueType = dynamicDefinition?.DataType ?? property.Property.PropertyType;
        Formatter = dynamicDefinition?.NumberFormat ?? property.Formatter;
        DecimalScale = property.DecimalScale;
        ValueMap = property.ValueMap;
        Ignored = property.Ignored;
        IsUnique = property.Property.IsDefined(typeof(DuplicationAttribute), false);
        IsMerged = property.Property.IsDefined(typeof(MergeColumnsAttribute), false);
        ImageMultiplicity = dynamicDefinition?.ImageMultiplicity ?? property.ImageMultiplicity;
        HeaderStyle = dynamicDefinition?.HeaderStyle;
        BodyStyle = dynamicDefinition?.BodyStyle;
        ValueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
        AttributeValidations = attributeValidations ?? Array.Empty<ExcelAttributeValidationBinding>();
        NamedValidationRules = namedValidationRules ?? Array.Empty<INamedExcelValidationRule>();
    }

    internal string HeaderName { get; }
    internal string Title { get; }
    internal ExcelPropertyMap Property { get; }
    internal bool IsDynamic { get; }
    internal int ColumnIndex { get; }
    internal ExcelDynamicColumnDefinition DynamicDefinition { get; }
    internal string Key { get; }
    internal Func<object, object> Getter { get; }
    internal Action<object, object> Setter { get; }
    internal string ConverterName { get; }
    internal IReadOnlyList<string> ValidationRuleNames { get; }
    internal string ValidatorName { get; }
    internal Type ValueType { get; }
    internal string Formatter { get; }
    internal byte? DecimalScale { get; }
    internal IReadOnlyDictionary<string, object> ValueMap { get; }
    internal bool Ignored { get; }
    internal bool IsUnique { get; }
    internal bool IsMerged { get; }
    internal ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    internal ExcelCellStyle HeaderStyle { get; }
    internal ExcelCellStyle BodyStyle { get; }
    internal IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    internal IReadOnlyList<ExcelAttributeValidationBinding> AttributeValidations { get; }
    internal IReadOnlyList<INamedExcelValidationRule> NamedValidationRules { get; }

    internal object ConvertFrom(string value, ExcelCellValue cellValue, string sheetName, int rowIndex,
        int columnIndex, CultureInfo culture)
    {
        if (cellValue.Kind == ExcelCellKind.Error || cellValue.CachedKind == ExcelCellKind.Error)
            throw new InvalidOperationException($"单元格包含公式错误码: {cellValue.ErrorCode}");
        var context = new ExcelConversionContext(value, Key, ValueType, sheetName, rowIndex, columnIndex,
            culture, cellValue);
        foreach (var converter in ValueConverters)
        {
            if (converter.TryConvertFrom(context, out var convertedValue))
                return convertedValue;
        }
        if (string.IsNullOrWhiteSpace(value))
        {
            if (!ValueType.IsValueType || Nullable.GetUnderlyingType(ValueType) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {ValueType.FullName}");
        }
        if (ValueMap.TryGetValue(value, out var mappedValue))
            return mappedValue;
        var targetType = Nullable.GetUnderlyingType(ValueType) ?? ValueType;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, value, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(value);
        if (targetType == typeof(Version))
            return new Version(value);
        if (targetType == typeof(DateTime))
            return DateTime.Parse(value, culture);
        return Convert.ChangeType(value, targetType, culture);
    }

    internal object ConvertTo(object value, string sheetName, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var context = new ExcelConversionContext(value, Key, ValueType, sheetName, rowIndex, columnIndex, culture);
        foreach (var converter in ValueConverters)
        {
            if (converter.TryConvertTo(context, out var convertedValue))
                return convertedValue;
        }
        if (string.IsNullOrWhiteSpace(Formatter) && ValueMap.Count > 0)
        {
            var mapping = ValueMap.FirstOrDefault(pair => IsMappedValue(pair.Value, value));
            if (mapping.Key != null)
                return mapping.Key;
        }
        if (value == null || ValueType == null || ValueType == typeof(object))
            return value;
        var targetType = Nullable.GetUnderlyingType(ValueType) ?? ValueType;
        if (targetType.IsInstanceOfType(value))
            return value;
        try
        {
            if (targetType.IsEnum)
                return value is string text ? Enum.Parse(targetType, text, true) : Enum.ToObject(targetType, value);
            if (targetType == typeof(Guid))
                return Guid.Parse(Convert.ToString(value, culture));
            if (targetType == typeof(DateTime))
                return value is DateTime date ? date : DateTime.Parse(Convert.ToString(value, culture), culture);
            return Convert.ChangeType(value, targetType, culture);
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException($"列 {Key} 的值无法转换为 {targetType.FullName}: {exception.Message}",
                exception);
        }
    }

    internal void WriteValue(ICell cell, object value)
    {
        if (!string.IsNullOrWhiteSpace(Formatter))
            cell.SetCellValue(value, Formatter);
        else
            cell.SetValue(value, DecimalScale);
    }

    private static bool IsMappedValue(object mappingValue, object value)
    {
        if (mappingValue == null || value == null)
            return mappingValue == null && value == null;
        if (Equals(mappingValue, value))
            return true;
        if (value is not Enum enumValue)
            return false;
        var underlyingType = Enum.GetUnderlyingType(enumValue.GetType());
        return Equals(mappingValue, Convert.ChangeType(enumValue, underlyingType, CultureInfo.InvariantCulture));
    }
}

internal sealed class ExcelAttributeValidationBinding
{
    internal ExcelAttributeValidationBinding(FilterAttributeBase attribute, IExcelValidationRule rule,
        bool isRaw)
    {
        Attribute = attribute;
        Rule = rule;
        IsRaw = isRaw;
    }

    internal FilterAttributeBase Attribute { get; }
    internal IExcelValidationRule Rule { get; }
    internal bool IsRaw { get; }
}
