using System.Reflection;
using Bing.Offices.Attributes;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 内部扩展
/// </summary>
internal static class InternalExtensions
{
    /// <summary>
    /// 获取字符串值
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="dto">数据传输对象</param>
    /// <param name="propertyName">属性名</param>
    /// <param name="format">格式化字符串</param>
    public static string GetStringValue<T>(this T dto, string propertyName, string format = "")
    {
        var value = string.Empty;
        var prop = dto.GetType().GetProperties().SingleOrDefault(p => p.Name.Equals(propertyName));
        if (prop != null)
            value = Format(prop.GetValue(dto), format);
        return value;
    }

    /// <summary>
    /// 获取值
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="dto">数据传输对象</param>
    /// <param name="propertyName">属性名</param>
    public static object GetValue<T>(this T dto, string propertyName)
    {
        var prop = dto.GetType().GetProperties().SingleOrDefault(p => p.Name.Equals(propertyName));
        return prop?.GetValue(dto);
    }

    /// <summary>
    /// 格式化
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="format">格式化字符串</param>
    /// <param name="formatProvider">格式化提供程序</param>
    private static string Format(object value, string format, IFormatProvider formatProvider = null)
    {
        if (value == null)
            return string.Empty;
        if (string.IsNullOrWhiteSpace(format))
            return value.ToString();
        if (value is IFormattable formattable)
            return formattable.ToString(format, formatProvider);
        throw new ArgumentException(nameof(value));
    }

    /// <summary>
    /// 获取扩展字典
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="dto">数据传输对象</param>
    /// <param name="propertyName">属性名</param>
    public static IDictionary<string, object> GetExtendDictionary<T>(this T dto, string propertyName)
    {
        var value = new Dictionary<string, object>();
        var prop = dto.GetType().GetProperties().SingleOrDefault(p =>
            p.Name.Equals(propertyName) && p.GetCustomAttribute<DynamicColumnAttribute>() != null);
        if (prop != null)
            value = prop.GetValue(dto) == null
                ? new Dictionary<string, object>()
                : (Dictionary<string, object>)prop.GetValue(dto);
        return value;
    }
}
