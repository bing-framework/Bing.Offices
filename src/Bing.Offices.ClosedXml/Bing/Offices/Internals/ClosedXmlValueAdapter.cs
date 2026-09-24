using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 将 ClosedXML 的对象值转换为公共 Bing.Offices 单元格语义。
/// </summary>
internal static class ClosedXmlValueAdapter
{
    /// <summary>
    /// 将固定映射列的单元格值转换为目标属性值。
    /// </summary>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="column">固定映射列。</param>
    /// <param name="property">目标属性。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">工作表行号。</param>
    /// <param name="columnIndex">工作表列号。</param>
    /// <param name="culture">转换使用的区域性。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="cell">已有的公共单元格值；为空时按原始值创建。</param>
    /// <returns>转换后的属性值。</returns>
    public static object ConvertFrom(object raw, IExcelMappingColumn column, PropertyInfo property,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture, bool isDate1904 = false,
        ExcelCellValue cell = null)
    {
        var text = ToText(raw, culture);
        var context = new ExcelConversionContext(raw, column.Name, property.PropertyType, sheetName,
            rowIndex, columnIndex, culture, cell ?? CreateCell(raw, text, isDate1904));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
            if (converter.TryConvertFrom(context, out var converted))
                return NormalizeDateValue(converted, property.PropertyType);
        if (column.ValueMap != null && column.ValueMap.TryGetValue(text, out var mapped))
            return NormalizeDateValue(ConvertRaw(mapped, mapped, property.PropertyType, culture, isDate1904,
                property.GetCustomAttribute<ExcelDateAttribute>(), cell), property.PropertyType);
        return NormalizeDateValue(ConvertRaw(raw, text, property.PropertyType, culture, isDate1904,
            property.GetCustomAttribute<ExcelDateAttribute>(), cell), property.PropertyType);
    }

    /// <summary>
    /// 将动态映射列的单元格值转换为动态列目标值。
    /// </summary>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="column">动态映射列。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">工作表行号。</param>
    /// <param name="columnIndex">工作表列号。</param>
    /// <param name="culture">转换使用的区域性。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="cell">已有的公共单元格值；为空时按原始值创建。</param>
    /// <returns>转换后的动态列值。</returns>
    public static object ConvertDynamicFrom(object raw, IExcelDynamicMappingColumn column,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture, bool isDate1904 = false,
        ExcelCellValue cell = null)
    {
        var type = ResolveDynamicType(column.DataTypeName);
        var text = ToText(raw, culture);
        var context = new ExcelConversionContext(raw, column.Key, type, sheetName,
            rowIndex, columnIndex, culture, cell ?? CreateCell(raw, text, isDate1904));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
            if (converter.TryConvertFrom(context, out var converted))
                return NormalizeDateValue(converted, type);
        return NormalizeDateValue(ConvertRaw(raw, text, type, culture, isDate1904,
            type == typeof(DateTime) || type == typeof(DateTimeOffset) ? new ExcelDateAttribute() : null, cell), type);
    }

    /// <summary>
    /// 将固定映射列的属性值转换为可写入单元格的值。
    /// </summary>
    /// <param name="raw">属性原始值。</param>
    /// <param name="column">固定映射列。</param>
    /// <param name="property">源属性。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">工作表行号。</param>
    /// <param name="columnIndex">工作表列号。</param>
    /// <param name="culture">转换使用的区域性。</param>
    /// <returns>可写入 ClosedXML 单元格的值。</returns>
    public static object ConvertTo(object raw, IExcelMappingColumn column, PropertyInfo property,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var context = new ExcelConversionContext(raw, column.Name, property.PropertyType, sheetName,
            rowIndex, columnIndex, culture, CreateCell(raw, ToText(raw, culture)));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
            if (converter.TryConvertTo(context, out var converted))
                return converted;
        if (raw == null)
            return null;
        if (column.ValueMap != null)
        {
            var rawText = Convert.ToString(raw, culture);
            var mapped = column.ValueMap.FirstOrDefault(pair =>
                string.Equals(pair.Value, rawText, StringComparison.Ordinal));
            if (!string.IsNullOrEmpty(mapped.Key))
                return mapped.Key;
        }
        if (!string.IsNullOrWhiteSpace(column.Formatter) && raw is IFormattable formattable)
            return formattable.ToString(column.Formatter, culture);
        return raw is DateTimeOffset offset ? offset.DateTime : raw;
    }

    /// <summary>
    /// 将动态映射列的值转换为可写入单元格的值。
    /// </summary>
    /// <param name="raw">动态列原始值。</param>
    /// <param name="column">动态映射列。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">工作表行号。</param>
    /// <param name="columnIndex">工作表列号。</param>
    /// <param name="culture">转换使用的区域性。</param>
    /// <returns>可写入 ClosedXML 单元格的值。</returns>
    public static object ConvertDynamicTo(object raw, IExcelDynamicMappingColumn column,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var type = ResolveDynamicType(column.DataTypeName);
        var context = new ExcelConversionContext(raw, column.Key, type, sheetName,
            rowIndex, columnIndex, culture, CreateCell(raw, ToText(raw, culture)));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
            if (converter.TryConvertTo(context, out var converted))
                return converted;
        return raw is DateTimeOffset offset ? offset.DateTime : raw;
    }

    /// <summary>
    /// 按指定区域性将值转换为文本。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <param name="culture">转换使用的区域性。</param>
    /// <returns>转换后的文本；值为空时返回空字符串。</returns>
    public static string ToText(object value, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;
        return value is IFormattable formattable
            ? formattable.ToString(null, culture)
            : Convert.ToString(value, culture) ?? string.Empty;
    }

    /// <summary>
    /// 按目标类型将原始单元格值转换为公共属性值。
    /// </summary>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">按区域性格式化的文本值。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <param name="culture">转换使用的区域性。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="dateAttribute">目标属性的日期配置。</param>
    /// <param name="cell">公共单元格值。</param>
    /// <returns>转换后的值；可空目标遇到空输入时返回 <see langword="null" />。</returns>
    public static object ConvertRaw(object raw, string text, Type propertyType, CultureInfo culture,
        bool isDate1904 = false, ExcelDateAttribute dateAttribute = null, ExcelCellValue cell = null)
    {
        var nullableType = Nullable.GetUnderlyingType(propertyType);
        var targetType = nullableType ?? propertyType;
        if (raw == null || string.IsNullOrWhiteSpace(text))
        {
            if (!propertyType.IsValueType || nullableType != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {propertyType.FullName}");
        }
        if (targetType == typeof(string))
            return text;
        if (targetType == typeof(DateTime) || targetType == typeof(DateTimeOffset))
        {
            DateTime date;
            if (raw is DateTime dateValue)
                date = dateValue;
            else if (raw is DateTimeOffset offset)
                date = offset.DateTime;
            else if (double.TryParse(text, NumberStyles.Float, culture, out var serial))
                date = DateTime.FromOADate(isDate1904 ? serial + 1462 : serial);
            else if (!DateTime.TryParse(text, culture, DateTimeStyles.AllowWhiteSpaces, out date))
                throw new InvalidCastException($"日期值转换失败。输入值: {text}，目标类型: {propertyType.FullName}");
            date = DateTime.SpecifyKind(date, DateTimeKind.Unspecified);
            return targetType == typeof(DateTimeOffset) ? new DateTimeOffset(date, TimeSpan.Zero) : date;
        }
        if (targetType.IsInstanceOfType(raw))
            return raw;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, text, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(text);
        if (targetType == typeof(Version))
            return new Version(text);
        if (targetType == typeof(byte[]))
            return raw is byte[] bytes ? bytes : Convert.FromBase64String(text);
        if (targetType == typeof(bool) && bool.TryParse(text, out var boolean))
            return boolean;
        return Convert.ChangeType(raw, targetType, culture);
    }

    /// <summary>
    /// 将动态列配置名称解析为运行时类型。
    /// </summary>
    /// <param name="name">动态列数据类型名称。</param>
    /// <returns>对应的运行时类型。</returns>
    public static Type ResolveDynamicType(string name)
    {
        switch ((name ?? "string").Trim().ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int":
            case "int32": return typeof(int);
            case "long":
            case "int64": return typeof(long);
            case "decimal": return typeof(decimal);
            case "double": return typeof(double);
            case "float":
            case "single": return typeof(float);
            case "bool":
            case "boolean": return typeof(bool);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            case "string": return typeof(string);
            default: throw new NotSupportedException($"ClosedXML 动态列类型不受支持: {name}");
        }
    }

    /// <summary>
    /// 从原始值创建公共单元格值描述。
    /// </summary>
    /// <param name="value">单元格原始值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <returns>公共单元格值描述。</returns>
    public static ExcelCellValue CreateCell(object value, string text, bool isDate1904 = false)
    {
        var kind = value switch
        {
            null => ExcelCellKind.Empty,
            bool => ExcelCellKind.Boolean,
            DateTime or DateTimeOffset => ExcelCellKind.DateTime,
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
                => ExcelCellKind.Number,
            _ => ExcelCellKind.Text
        };
        return new ExcelCellValue(value, text, kind, isDate1904: isDate1904);
    }

    /// <summary>
    /// 将公共校验绑定类型转换为导入错误代码。
    /// </summary>
    /// <param name="binding">校验绑定。</param>
    /// <returns>对应的导入错误代码。</returns>
    public static ExcelImportErrorCode GetValidationCode(IExcelValidationBinding binding) => binding.Kind switch
    {
        ExcelValidationBindingKind.MaxLength => ExcelImportErrorCode.MaxLength,
        ExcelValidationBindingKind.MaxValue => ExcelImportErrorCode.MaxValue,
        _ => ExcelImportErrorCode.Validation
    };

    /// <summary>
    /// 在日期类型之间执行公共转换结果的归一化。
    /// </summary>
    /// <param name="value">待归一化的值。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <returns>归一化后的日期值或原值；输入为空时返回 <see langword="null" />。</returns>
    private static object NormalizeDateValue(object value, Type propertyType)
    {
        if (value == null)
            return null;
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (targetType == typeof(DateTime) && value is DateTimeOffset offset)
            return offset.DateTime;
        if (targetType == typeof(DateTimeOffset) && value is DateTime date)
            return new DateTimeOffset(date, TimeSpan.Zero);
        return value;
    }
}
