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
    private static readonly DateTimeExcelValidationRule DateRule = new DateTimeExcelValidationRule();

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

    public static string ToText(object value, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;
        return value is IFormattable formattable
            ? formattable.ToString(null, culture)
            : Convert.ToString(value, culture) ?? string.Empty;
    }

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
