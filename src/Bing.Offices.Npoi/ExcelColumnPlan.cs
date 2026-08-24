using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
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
    internal ExcelColumnPlan(string headerName, IExcelMappingColumn property, bool isDynamic, int columnIndex,
        ExcelDynamicColumnDefinition dynamicDefinition, string key = null,
        IReadOnlyList<IExcelValueConverter> valueConverters = null,
        IReadOnlyList<IExcelValidationBinding> validationBindings = null,
        PropertyInfo reflectionProperty = null,
        bool? isUnique = null,
        bool uniqueIgnoreEmpty = true)
    {
        HeaderName = headerName;
        Title = headerName;
        Property = property;
        ReflectionProperty = reflectionProperty;
        IsDynamic = isDynamic;
        ColumnIndex = columnIndex;
        DynamicDefinition = dynamicDefinition;
        Key = key ?? dynamicDefinition?.Key ?? (isDynamic ? headerName : property.Name);
        if (reflectionProperty == null)
            throw new InvalidOperationException($"无法解析映射属性: {property.Name}");
        Getter = instance => reflectionProperty.GetValue(instance);
        Setter = (instance, value) => reflectionProperty.SetValue(instance, value);
        ConverterName = dynamicDefinition?.ConverterName ?? property.ConverterName;
        ValidationRuleNames = property.ValidationRuleNames;
        ValidatorName = dynamicDefinition?.ValidatorName;
        ValueType = dynamicDefinition?.DataType ?? reflectionProperty.PropertyType;
        Formatter = dynamicDefinition?.NumberFormat ?? property.Formatter;
        DecimalScale = property.DecimalScale;
        ValueMap = property.ValueMap;
        Ignored = property.Ignored;
        var attributes = reflectionProperty.GetCustomAttributes<Attribute>().ToArray();
        IsUnique = isUnique ?? attributes.Any(attribute => attribute is DuplicationAttribute
            || attribute is ExcelUniqueAttribute);
        UniqueIgnoreEmpty = isUnique.HasValue ? uniqueIgnoreEmpty
            : attributes.OfType<ExcelUniqueAttribute>().FirstOrDefault()?.IgnoreEmpty ?? true;
        IsMerged = attributes.Any(attribute => attribute is MergeColumnsAttribute);
        ImageMultiplicity = dynamicDefinition?.ImageMultiplicity ?? property.ImageMultiplicity;
        HeaderStyle = dynamicDefinition?.HeaderStyle;
        BodyStyle = dynamicDefinition?.BodyStyle;
        ValueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
        ValidationBindings = validationBindings ?? Array.Empty<IExcelValidationBinding>();
    }

    internal string HeaderName { get; }
    internal string Title { get; }
    internal IExcelMappingColumn Property { get; }
    internal PropertyInfo ReflectionProperty { get; }
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
    internal IReadOnlyDictionary<string, string> ValueMap { get; }
    internal bool Ignored { get; }
    internal bool IsUnique { get; }
    internal bool UniqueIgnoreEmpty { get; }
    internal bool IsMerged { get; }
    internal ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    internal ExcelCellStyle HeaderStyle { get; }
    internal ExcelCellStyle BodyStyle { get; }
    internal IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    internal IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }

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
            return ConvertMappedValue(mappedValue, culture);
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
            var mapping = ValueMap.FirstOrDefault(pair => IsMappedValue(pair.Value, value, culture));
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

    private object ConvertMappedValue(string value, CultureInfo culture)
    {
        if (value == null)
            return null;
        var targetType = Nullable.GetUnderlyingType(ValueType) ?? ValueType;
        if (targetType == typeof(string))
            return value;
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

    private static bool IsMappedValue(string mappingValue, object value, CultureInfo culture)
    {
        if (mappingValue == null || value == null)
            return mappingValue == null && value == null;
        return string.Equals(mappingValue, Convert.ToString(value, culture), StringComparison.Ordinal);
    }
}
