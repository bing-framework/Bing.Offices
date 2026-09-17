using System.Globalization;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;

namespace Bing.Offices.Dates;

/// <summary>无 offset 文本转换为 DateTimeOffset 时使用的策略。</summary>
public enum ExcelDateOffsetPolicy
{
    /// <summary>必须由输入文本显式提供 offset。</summary>
    RequireExplicitOffset,
    /// <summary>使用配置的固定 offset。</summary>
    UseFixedOffset
}

/// <summary>Excel 与 CSV 共用的确定性日期解析器。</summary>
internal static class ExcelDateParser
{
    /// <summary>未指定格式时使用的默认日期格式（yyyy-MM-dd）。</summary>
    private const string DefaultDateFormat = "yyyy-MM-dd";

    /// <summary>Excel 1900 日期系统的零点。</summary>
    private static readonly DateTime Excel1900Epoch = new DateTime(1899, 12, 31);

    /// <summary>一天包含的毫秒数，用于跨目标框架保持 serial 精度。</summary>
    private const double MillisecondsPerDay = 86400000d;

    /// <summary>
    /// 按日期特性配置将单元格值转换为目标日期类型。
    /// </summary>
    /// <param name="cell">原始单元格描述。</param>
    /// <param name="text">规范化文本。</param>
    /// <param name="targetType">目标类型。</param>
    /// <param name="culture">请求区域性。</param>
    /// <param name="attribute">可选日期特性。</param>
    /// <param name="value">转换结果。</param>
    /// <returns>能够转换时为 true。</returns>
    public static bool TryParse(ExcelCellValue cell, string text, Type targetType, CultureInfo culture,
        ExcelDateAttribute attribute, out object value)
    {
        value = null;
        var effectiveType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (effectiveType != typeof(DateTime) && effectiveType != typeof(DateTimeOffset))
            return false;
        if (cell?.Value is DateTime dateTime)
        {
            var unspecified = DateTime.SpecifyKind(dateTime, DateTimeKind.Unspecified);
            if (effectiveType == typeof(DateTime))
            {
                value = unspecified;
                return true;
            }
            if (effectiveType == typeof(DateTimeOffset))
            {
                if (!TryGetFixedOffset(attribute, out var offset))
                    return false;
                try
                {
                    value = new DateTimeOffset(unspecified, offset);
                    return true;
                }
                catch (ArgumentException)
                {
                    return false;
                }
            }
        }
        if (cell?.Value is DateTimeOffset dateTimeOffset)
        {
            if (effectiveType == typeof(DateTimeOffset))
            {
                value = dateTimeOffset;
                return true;
            }
            if (effectiveType == typeof(DateTime))
            {
                value = DateTime.SpecifyKind(dateTimeOffset.DateTime, DateTimeKind.Unspecified);
                return true;
            }
        }
        if (cell?.Value is double serial
            && (effectiveType == typeof(DateTime) || effectiveType == typeof(DateTimeOffset)))
        {
            try
            {
                var date = DateTime.SpecifyKind(FromExcelSerial(serial, cell.IsDate1904),
                    DateTimeKind.Unspecified);
                // DateTime 可隐式转换为 DateTimeOffset；显式分支避免条件表达式把无时区值转换成本地 offset。
                if (effectiveType == typeof(DateTime))
                    value = date;
                else
                {
                    if (!TryGetFixedOffset(attribute, out var offset))
                        return false;
                    value = new DateTimeOffset(date, offset);
                }
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
        var effectiveCulture = culture ?? CultureInfo.InvariantCulture;
        if (!string.IsNullOrWhiteSpace(attribute?.CultureName))
            effectiveCulture = CultureInfo.GetCultureInfo(attribute.CultureName);
        var format = attribute?.Format;
        if (effectiveType == typeof(DateTime))
        {
            if (DateTime.TryParseExact(text, format ?? DefaultDateFormat, effectiveCulture,
                    DateTimeStyles.None, out var parsed))
            {
                value = DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified);
                return true;
            }
            return false;
        }

        var offsetPolicy = attribute?.OffsetPolicy ?? ExcelDateOffsetPolicy.RequireExplicitOffset;
        var offsetMinutes = attribute?.OffsetMinutes;
        if (offsetPolicy == ExcelDateOffsetPolicy.UseFixedOffset && !offsetMinutes.HasValue)
            return false;
        if (HasExplicitOffset(text))
        {
            var formats = format == null
                ? new[] { "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK", "yyyy-MM-dd'T'HH:mm:ssK" }
                : new[] { format };
            foreach (var candidate in formats)
            {
                if (DateTimeOffset.TryParseExact(text, candidate, effectiveCulture,
                        DateTimeStyles.None, out var parsed))
                {
                    value = parsed;
                    return true;
                }
            }
            return false;
        }
        if (offsetPolicy != ExcelDateOffsetPolicy.UseFixedOffset)
            return false;
        var localFormat = format ?? DefaultDateFormat;
        if (!DateTime.TryParseExact(text, localFormat, effectiveCulture, DateTimeStyles.None,
                out var localDate))
            return false;
        try
        {
            value = new DateTimeOffset(DateTime.SpecifyKind(localDate, DateTimeKind.Unspecified),
                TimeSpan.FromMinutes(offsetMinutes.Value));
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (OverflowException)
        {
            return false;
        }
    }

    /// <summary>
    /// 按 Workbook Data Validation 的日期或时间语义解析值。
    /// </summary>
    /// <param name="cell">原始单元格值。</param>
    /// <param name="text">单元格或约束的文本值。</param>
    /// <param name="timeOnly">是否只解析时间部分。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="value">解析后的无时区日期时间。</param>
    /// <returns>解析成功时为 true。</returns>
    internal static bool TryParseValidation(ExcelCellValue cell, string text, bool timeOnly,
        bool isDate1904, out DateTime value)
    {
        value = default;
        if (cell?.Value is DateTime date)
        {
            value = DateTime.SpecifyKind(timeOnly ? DateTime.MinValue.Add(date.TimeOfDay) : date,
                DateTimeKind.Unspecified);
            return true;
        }
        if (cell?.Value is double serial)
            return TryFromExcelSerial(serial, timeOnly, isDate1904, out value);
        if (string.IsNullOrWhiteSpace(text))
            return false;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var textSerial))
            return TryFromExcelSerial(textSerial, timeOnly, isDate1904, out value);
        if (timeOnly)
        {
            if (!DateTime.TryParseExact(text, new[] { "HH:mm:ss.FFFFFFF", "HH:mm:ss", "H:mm:ss" },
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var time))
                return false;
            value = DateTime.MinValue.Add(time.TimeOfDay);
            return true;
        }
        if (!TryParse(new ExcelCellValue(text, text, ExcelCellKind.Text, isDate1904: isDate1904), text,
                typeof(DateTime), CultureInfo.InvariantCulture, null, out var parsed))
            return false;
        value = (DateTime)parsed;
        return true;
    }

    /// <summary>判断日期文本是否显式包含时区偏移或 UTC 标记。</summary>
    /// <param name="text">待检查的日期文本。</param>
    /// <returns>文本包含可识别的显式时区信息时为 true。</returns>
    private static bool HasExplicitOffset(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;
        if (text.EndsWith("Z", StringComparison.OrdinalIgnoreCase))
            return true;
        var separator = text.LastIndexOfAny(new[] { '+', '-' });
        return separator > 10 && text.Length - separator >= 6
            && text[separator + 3] == ':';
    }

    /// <summary>获取无显式 offset 日期使用的固定偏移。</summary>
    /// <param name="attribute">日期输入配置。</param>
    /// <param name="offset">配置的时间偏移。</param>
    /// <returns>配置为固定偏移且包含偏移分钟时为 true。</returns>
    private static bool TryGetFixedOffset(ExcelDateAttribute attribute, out TimeSpan offset)
    {
        offset = default;
        if ((attribute?.OffsetPolicy ?? ExcelDateOffsetPolicy.RequireExplicitOffset)
            != ExcelDateOffsetPolicy.UseFixedOffset
            || !attribute.OffsetMinutes.HasValue)
            return false;
        try
        {
            offset = TimeSpan.FromMinutes(attribute.OffsetMinutes.Value);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (OverflowException)
        {
            return false;
        }
    }

    /// <summary>将 Excel 序列日期转换为无时区日期时间。</summary>
    /// <param name="serial">Excel 序列日期数值。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <returns>对应的无时区日期时间。</returns>
    private static DateTime FromExcelSerial(double serial, bool isDate1904)
    {
        // 负数不能通过 FromOADate 往返：OA 日期会把负数的小数部分按相反方向折叠。
        // 1904 日期系统没有伪闰日，所有 serial 均按工作簿零点线性计算。
        if (isDate1904)
            return AddLinearExcelDays(new DateTime(1904, 1, 1), serial);

        if (serial < 0d)
            return AddLinearExcelDays(Excel1900Epoch, serial);

        // Excel 1900 系统保留序列 60 的伪闰日；其余序列映射到 OLE Automation 日期时，
        // 60 之前整体向后补一天，60 本身由 FromOADate 保持为 1900-02-28。
        return DateTime.FromOADate(serial < 60d ? serial + 1d : serial);
    }

    /// <summary>按 Excel serial 线性加日期并固定到毫秒精度。</summary>
    /// <param name="epoch">工作簿日期系统零点。</param>
    /// <param name="serial">待增加的 serial 天数。</param>
    /// <returns>线性转换后的日期时间。</returns>
    private static DateTime AddLinearExcelDays(DateTime epoch, double serial)
    {
        if (double.IsNaN(serial) || double.IsInfinity(serial))
            throw new ArgumentException("Excel serial 必须是有限数值。", nameof(serial));
        var milliseconds = serial * MillisecondsPerDay;
        if (milliseconds > long.MaxValue || milliseconds < long.MinValue)
            throw new ArgumentException("Excel serial 超出支持范围。", nameof(serial));
        return epoch.AddMilliseconds(Math.Round(milliseconds, MidpointRounding.AwayFromZero));
    }

    /// <summary>尝试按 Excel 日期系统将序列值转换为日期或时间。</summary>
    /// <param name="serial">Excel 序列日期数值。</param>
    /// <param name="timeOnly">是否只保留序列值中的时间部分。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="value">转换得到的无时区日期时间。</param>
    /// <returns>序列值在支持范围内时为 true。</returns>
    private static bool TryFromExcelSerial(double serial, bool timeOnly, bool isDate1904,
        out DateTime value)
    {
        try
        {
            var date = FromExcelSerial(serial, isDate1904);
            value = DateTime.SpecifyKind(timeOnly ? DateTime.MinValue.Add(date.TimeOfDay) : date,
                DateTimeKind.Unspecified);
            return true;
        }
        catch (ArgumentException)
        {
            value = default;
            return false;
        }
    }
}
