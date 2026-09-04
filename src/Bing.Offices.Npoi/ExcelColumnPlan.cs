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
    /// <summary>根据映射列及反射属性创建导入导出共享的列执行计划。</summary>
    /// <param name="headerName">当前工作表中使用的表头名称。</param>
    /// <param name="property">提供程序无关的列映射。</param>
    /// <param name="isDynamic">是否为运行时定义的动态列。</param>
    /// <param name="columnIndex">当前列在工作表中的零基索引。</param>
    /// <param name="dynamicDefinition">动态列的请求级定义。</param>
    /// <param name="key">用于错误定位和动态字典访问的列键。</param>
    /// <param name="valueConverters">按优先级绑定的值转换器。</param>
    /// <param name="validationBindings">按配置绑定的校验规则。</param>
    /// <param name="reflectionProperty">实体上对应的可读写属性。</param>
    /// <param name="isUnique">是否启用唯一性校验；为 null 时从特性推断。</param>
    /// <param name="uniqueIgnoreEmpty">唯一性校验是否忽略空值。</param>
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
        IsUnique = isUnique ?? attributes.Any(attribute => attribute is ExcelUniqueAttribute);
        UniqueIgnoreEmpty = isUnique.HasValue ? uniqueIgnoreEmpty
            : attributes.OfType<ExcelUniqueAttribute>().FirstOrDefault()?.IgnoreEmpty ?? true;
        IsMerged = attributes.Any(attribute => attribute is MergeColumnsAttribute);
        ImageMultiplicity = dynamicDefinition?.ImageMultiplicity ?? property.ImageMultiplicity;
        HeaderStyle = dynamicDefinition?.HeaderStyle;
        BodyStyle = dynamicDefinition?.BodyStyle;
        ValueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
        ValidationBindings = validationBindings ?? Array.Empty<IExcelValidationBinding>();
    }

    /// <summary>获取实际匹配到的表头名称。</summary>
    internal string HeaderName { get; }
    /// <summary>获取导出时写入的列标题。</summary>
    internal string Title { get; }
    /// <summary>获取提供程序无关的列映射。</summary>
    internal IExcelMappingColumn Property { get; }
    /// <summary>获取实体上对应的反射属性。</summary>
    internal PropertyInfo ReflectionProperty { get; }
    /// <summary>获取是否为运行时定义的动态列。</summary>
    internal bool IsDynamic { get; }
    /// <summary>获取工作表中的零基列索引。</summary>
    internal int ColumnIndex { get; }
    /// <summary>获取动态列的请求级定义；固定列时为 null。</summary>
    internal ExcelDynamicColumnDefinition DynamicDefinition { get; }
    /// <summary>获取用于错误定位和动态字典访问的稳定列键。</summary>
    internal string Key { get; }
    /// <summary>获取从实体读取当前列值的委托。</summary>
    internal Func<object, object> Getter { get; }
    /// <summary>获取将转换后值写回实体的委托。</summary>
    internal Action<object, object> Setter { get; }
    /// <summary>获取显式配置的转换器名称。</summary>
    internal string ConverterName { get; }
    /// <summary>获取配置的命名校验规则名称。</summary>
    internal IReadOnlyList<string> ValidationRuleNames { get; }
    /// <summary>获取兼容的单个命名校验规则名称。</summary>
    internal string ValidatorName { get; }
    /// <summary>获取当前列最终使用的值类型。</summary>
    internal Type ValueType { get; }
    /// <summary>获取导出值使用的格式化字符串。</summary>
    internal string Formatter { get; }
    /// <summary>获取未指定格式化字符串时的十进制小数位数。</summary>
    internal byte? DecimalScale { get; }
    /// <summary>获取显示文本到配置值文本的映射。</summary>
    internal IReadOnlyDictionary<string, string> ValueMap { get; }
    /// <summary>获取是否忽略该映射列。</summary>
    internal bool Ignored { get; }
    /// <summary>获取是否对列值执行唯一性校验。</summary>
    internal bool IsUnique { get; }
    /// <summary>获取唯一性校验是否忽略空值。</summary>
    internal bool UniqueIgnoreEmpty { get; }
    /// <summary>获取列属性是否声明合并单元格行为。</summary>
    internal bool IsMerged { get; }
    /// <summary>获取同一单元格含多个图片时使用的处理策略。</summary>
    internal ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <summary>获取动态列表头使用的样式。</summary>
    internal ExcelCellStyle HeaderStyle { get; }
    /// <summary>获取动态列数据单元格使用的样式。</summary>
    internal ExcelCellStyle BodyStyle { get; }
    /// <summary>获取按优先级绑定的值转换器。</summary>
    internal IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <summary>获取按配置绑定的校验规则。</summary>
    internal IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }

    /// <summary>将导入文本转换为当前列的目标值。</summary>
    /// <param name="value">规范化后的文本值。</param>
    /// <param name="cellValue">保留原始类型和公式信息的单元格值。</param>
    /// <param name="sheetName">单元格所在工作表名称。</param>
    /// <param name="rowIndex">单元格所在的一基行号。</param>
    /// <param name="columnIndex">单元格所在的一基列号。</param>
    /// <param name="culture">文本转换使用的区域性。</param>
    /// <returns>可写入实体属性的转换结果。</returns>
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

    /// <summary>将实体属性值转换为可写入工作表的值。</summary>
    /// <param name="value">实体属性的当前值。</param>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="rowIndex">目标单元格的一基行号。</param>
    /// <param name="columnIndex">目标单元格的一基列号。</param>
    /// <param name="culture">值转换使用的区域性。</param>
    /// <returns>适合写入 NPOI 单元格的值。</returns>
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

    /// <summary>按照列格式和精度规则将值写入单元格。</summary>
    /// <param name="cell">目标 NPOI 单元格。</param>
    /// <param name="value">已转换的列值。</param>
    internal void WriteValue(ICell cell, object value)
    {
        if (!string.IsNullOrWhiteSpace(Formatter))
            cell.SetCellValue(value, Formatter);
        else
            cell.SetValue(value, DecimalScale);
    }

    /// <summary>将配置值映射文本转换为目标列类型。</summary>
    /// <param name="value">映射配置中保存的文本值。</param>
    /// <param name="culture">值转换使用的区域性。</param>
    /// <returns>目标类型的映射值；文本为空引用时返回 null。</returns>
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

    /// <summary>比较配置映射值与当前实体值的文本表示。</summary>
    /// <param name="mappingValue">映射配置中的值。</param>
    /// <param name="value">待比较的实体值。</param>
    /// <param name="culture">格式化实体值使用的区域性。</param>
    /// <returns>两个值语义相等时为 true。</returns>
    private static bool IsMappedValue(string mappingValue, object value, CultureInfo culture)
    {
        if (mappingValue == null || value == null)
            return mappingValue == null && value == null;
        return string.Equals(mappingValue, Convert.ToString(value, culture), StringComparison.Ordinal);
    }
}
