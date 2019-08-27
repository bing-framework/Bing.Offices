using System.Linq;

namespace Bing.Offices.Npoi.Extensions
{
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
        public static string GetStringValue<T>(this T dto, string propertyName)
        {
            var value = string.Empty;
            var prop = dto.GetType().GetProperties().SingleOrDefault(p => p.Name.Equals(propertyName));
            if (prop != null)
                value = prop.GetValue(dto) == null ? string.Empty : prop.GetValue(dto).ToString();
            return value;
        }
    }
}
