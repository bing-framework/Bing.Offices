using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// MiniExcel 原始值与 Bing.Offices 映射/转换契约之间的适配器。
/// </summary>
internal static class MiniExcelValueAdapter
{
    /// <summary>
    /// 用于解析日期和日期时间值的共享校验规则实例。
    /// </summary>
    private static readonly DateTimeExcelValidationRule DateRule = new DateTimeExcelValidationRule();

    /// <summary>
    /// 将 MiniExcel 原始值转换为固定列目标属性值。
    /// </summary>
    /// <param name="raw">原始单元格值。</param>
    /// <param name="column">固定列映射。</param>
    /// <param name="property">目标属性。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">一基物理行号。</param>
    /// <param name="columnIndex">一基物理列号。</param>
    /// <param name="culture">值转换使用的区域性设置。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="cell">可复用的 Excel 单元格值。</param>
    /// <returns>转换后的目标属性值。</returns>
    public static object ConvertFrom(object raw, IExcelMappingColumn column, PropertyInfo property,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture, bool isDate1904 = false,
        ExcelCellValue cell = null)
    {
        var text = ToText(raw, culture);
        var context = new ExcelConversionContext(raw, column.Name, property.PropertyType, sheetName,
            rowIndex, columnIndex, culture, cell ?? CreateCell(raw, text, isDate1904));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            if (converter.TryConvertFrom(context, out var converted))
                return converted;
        }

        if (column.ValueMap != null && column.ValueMap.TryGetValue(text, out var mapped))
            return ConvertRaw(mapped, mapped, property.PropertyType, culture, isDate1904,
                property.GetCustomAttribute<ExcelDateAttribute>());
        return ConvertRaw(raw, text, property.PropertyType, culture, isDate1904,
            property.GetCustomAttribute<ExcelDateAttribute>(), cell);
    }

    /// <summary>
    /// 将 MiniExcel 原始值转换为动态列目标值。
    /// </summary>
    /// <param name="raw">原始单元格值。</param>
    /// <param name="column">动态列映射。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">一基物理行号。</param>
    /// <param name="columnIndex">一基物理列号。</param>
    /// <param name="culture">值转换使用的区域性设置。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="cell">可复用的 Excel 单元格值。</param>
    /// <returns>按动态列数据类型转换后的值。</returns>
    public static object ConvertDynamicFrom(object raw, IExcelDynamicMappingColumn column,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture, bool isDate1904 = false,
        ExcelCellValue cell = null)
    {
        var type = ResolveDynamicType(column.DataTypeName);
        var text = ToText(raw, culture);
        var context = new ExcelConversionContext(raw, column.Key, type, sheetName,
            rowIndex, columnIndex, culture, cell ?? CreateCell(raw, text, isDate1904));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            if (converter.TryConvertFrom(context, out var converted))
                return converted;
        }
        return ConvertRaw(raw, text, type, culture, isDate1904,
            type == typeof(DateTime) || type == typeof(DateTimeOffset) ? new ExcelDateAttribute() : null, cell);
    }

    /// <summary>
    /// 将实体属性值转换为固定列导出值。
    /// </summary>
    /// <param name="raw">实体属性原始值。</param>
    /// <param name="column">固定列映射。</param>
    /// <param name="property">源属性。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">一基物理行号。</param>
    /// <param name="columnIndex">一基物理列号。</param>
    /// <param name="culture">值转换使用的区域性设置。</param>
    /// <returns>供 MiniExcel 写入的导出值。</returns>
    public static object ConvertTo(object raw, IExcelMappingColumn column, PropertyInfo property,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var context = new ExcelConversionContext(raw, column.Name, property.PropertyType, sheetName,
            rowIndex, columnIndex, culture, CreateCell(raw, ToText(raw, culture)));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            if (converter.TryConvertTo(context, out var converted))
                return converted;
        }

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
        if (raw is DateTimeOffset dateTimeOffset)
            return dateTimeOffset.ToString("O", CultureInfo.InvariantCulture);
        return raw;
    }

    /// <summary>
    /// 将实体值转换为动态列导出值。
    /// </summary>
    /// <param name="raw">实体属性原始值。</param>
    /// <param name="column">动态列映射。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">一基物理行号。</param>
    /// <param name="columnIndex">一基物理列号。</param>
    /// <param name="culture">值转换使用的区域性设置。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <returns>供 MiniExcel 写入的动态列值。</returns>
    public static object ConvertDynamicTo(object raw, IExcelDynamicMappingColumn column,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture, bool isDate1904 = false)
    {
        var type = ResolveDynamicType(column.DataTypeName);
        var context = new ExcelConversionContext(raw, column.Key, type, sheetName,
            rowIndex, columnIndex, culture, CreateCell(raw, ToText(raw, culture), isDate1904));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            if (converter.TryConvertTo(context, out var converted))
                return converted;
        }
        if (raw is DateTimeOffset dateTimeOffset)
        {
            if (!string.IsNullOrWhiteSpace(column.NumberFormat))
                return dateTimeOffset.ToString(column.NumberFormat, culture);
            return dateTimeOffset.ToString("O", CultureInfo.InvariantCulture);
        }
        return ConvertRaw(raw, ToText(raw, culture), type, culture, isDate1904,
            type == typeof(DateTime) || type == typeof(DateTimeOffset) ? new ExcelDateAttribute() : null);
    }

    /// <summary>
    /// 将原始值转换为区域性相关的文本。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <param name="culture">文本转换使用的区域性设置。</param>
    /// <returns>转换后的文本；空值返回空字符串。</returns>
    public static string ToText(object value, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;
        return value is IFormattable formattable
            ? formattable.ToString(null, culture)
            : Convert.ToString(value, culture) ?? string.Empty;
    }

    /// <summary>
    /// 按目标类型转换原始值，并处理日期、枚举和可空类型。
    /// </summary>
    /// <param name="raw">原始值。</param>
    /// <param name="text">规范化后的文本值。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <param name="culture">值转换使用的区域性设置。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="dateAttribute">目标日期属性的日期规则。</param>
    /// <param name="cell">可复用的 Excel 单元格值。</param>
    /// <returns>转换后的目标类型值。</returns>
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
            var sourceCell = cell ?? CreateCell(raw, text, isDate1904);
            if (DateRule.TryParseValue(sourceCell, text, propertyType, culture,
                dateAttribute ?? new ExcelDateAttribute(), out var parsed))
                return NormalizeDateResult(parsed, targetType);
            throw new InvalidCastException($"日期值转换失败。输入值: {text}，目标类型: {propertyType.FullName}");
        }
        if (targetType.IsInstanceOfType(raw))
            return raw;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, text, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(text);
        if (targetType == typeof(Version))
            return new Version(text);
        return Convert.ChangeType(raw, targetType, culture);
    }

    /// <summary>
    /// 将日期解析结果规范化为目标日期类型。
    /// </summary>
    /// <param name="value">日期解析结果。</param>
    /// <param name="targetType">目标日期类型。</param>
    /// <returns>符合目标类型和时区约定的日期值。</returns>
    private static object NormalizeDateResult(object value, Type targetType)
    {
        if (targetType == typeof(DateTime))
        {
            if (value is DateTime date)
                return DateTime.SpecifyKind(date, DateTimeKind.Unspecified);
            if (value is DateTimeOffset offset)
                return DateTime.SpecifyKind(offset.DateTime, DateTimeKind.Unspecified);
        }
        else if (targetType == typeof(DateTimeOffset))
        {
            if (value is DateTimeOffset offset)
                return offset;
            if (value is DateTime date)
                return new DateTimeOffset(DateTime.SpecifyKind(date, DateTimeKind.Unspecified),
                    TimeSpan.Zero);
        }
        return value;
    }

    /// <summary>
    /// 将动态列类型名称解析为 CLR 类型。
    /// </summary>
    /// <param name="name">动态列类型名称。</param>
    /// <returns>对应的 CLR 类型。</returns>
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
            default: throw new NotSupportedException($"MiniExcel 动态列类型不受支持: {name}");
        }
    }

    /// <summary>
    /// 根据 MiniExcel 值创建带日期系统标记的 Excel 单元格值。
    /// </summary>
    /// <param name="value">原始值。</param>
    /// <param name="text">原始文本。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <returns>描述值类型和日期系统的 Excel 单元格值。</returns>
    public static ExcelCellValue CreateCell(object value, string text, bool isDate1904 = false)
    {
        // MiniExcel 已将带日期样式的数值转换为 DateTime；恢复 OLE Automation serial 后，
        // 共享 Core 解析器才能继续应用 1900/1904 日期系统规则。
        if (value is DateTime dateTime)
        {
            try
            {
                return new ExcelCellValue(dateTime.ToOADate(), text, ExcelCellKind.Number,
                    isDate1904: isDate1904);
            }
            catch (ArgumentException)
            {
                // 超出 OADate 范围时保留原生日期，让 Core 返回其标准日期结果或错误。
            }
        }
        else if (value is DateTimeOffset dateTimeOffset)
        {
            try
            {
                return new ExcelCellValue(dateTimeOffset.DateTime.ToOADate(), text, ExcelCellKind.Number,
                    isDate1904: isDate1904);
            }
            catch (ArgumentException)
            {
                // 超出 OADate 范围时保留原生日期，让 Core 返回其标准日期结果或错误。
            }
        }
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
    /// 将映射校验绑定转换为导入错误代码。
    /// </summary>
    /// <param name="binding">待转换的校验绑定。</param>
    /// <returns>对应的导入错误代码。</returns>
    public static ExcelImportErrorCode GetValidationCode(IExcelValidationBinding binding)
    {
        return binding.Kind switch
        {
            ExcelValidationBindingKind.MaxLength => ExcelImportErrorCode.MaxLength,
            ExcelValidationBindingKind.MaxValue => ExcelImportErrorCode.MaxValue,
            ExcelValidationBindingKind.Unique => ExcelImportErrorCode.Validation,
            _ => ExcelImportErrorCode.Validation
        };
    }
}
