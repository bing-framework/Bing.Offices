using System.Reflection;
using Bing.Offices.Attributes;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
    /// Excel 导出和动态列处理使用的内部反射扩展。
/// </summary>
internal static class InternalExtensions
{
    /// <summary>
    /// 读取指定属性并按可选格式转换为字符串。
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="dto">数据传输对象</param>
    /// <param name="propertyName">属性名</param>
    /// <param name="format">格式化字符串</param>
    /// <returns>属性格式化文本；属性不存在或值为空时返回空字符串。</returns>
    public static string GetStringValue<T>(this T dto, string propertyName, string format = "")
    {
        var value = string.Empty;
        var prop = dto.GetType().GetProperties().SingleOrDefault(p => p.Name.Equals(propertyName));
        if (prop != null)
            value = Format(prop.GetValue(dto), format);
        return value;
    }

    /// <summary>
    /// 通过属性名称读取实体属性值。
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="dto">数据传输对象</param>
    /// <param name="propertyName">属性名</param>
    /// <returns>属性值；属性不存在时返回 null。</returns>
    public static object GetValue<T>(this T dto, string propertyName)
    {
        var prop = dto.GetType().GetProperties().SingleOrDefault(p => p.Name.Equals(propertyName));
        return prop?.GetValue(dto);
    }

    /// <summary>
    /// 使用指定格式和区域性提供程序将值转换为文本。
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="format">格式化字符串</param>
    /// <param name="formatProvider">格式化提供程序</param>
    /// <returns>格式化后的文本；值为空时返回空字符串。</returns>
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
    /// 读取标记为动态列的字典属性；未找到有效字典时返回新的空字典。
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="dto">数据传输对象</param>
    /// <param name="propertyName">属性名</param>
    /// <returns>动态列值字典。</returns>
    public static IDictionary<string, object> GetExtendDictionary<T>(this T dto, string propertyName)
    {
        var prop = dto.GetType().GetProperties().SingleOrDefault(p =>
            p.Name.Equals(propertyName) && p.GetCustomAttribute<DynamicColumnAttribute>() != null);
        return prop?.GetValue(dto) as IDictionary<string, object>
               ?? new Dictionary<string, object>();
    }
}
