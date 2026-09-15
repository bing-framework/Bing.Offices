using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using Bing.Offices.Attributes;

namespace Bing.Offices.Extensions;

/// <summary>
/// 属性信息(<see cref="PropertyInfo"/>) 扩展
/// </summary>
internal static class PropertyInfoExtensions
{
    /// <summary>
    /// 是否有忽略属性
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    /// <returns>属性声明 <see cref="NotMappedAttribute" /> 或 <see cref="ExcelIgnoreAttribute" /> 时为 <see langword="true" />。</returns>
    internal static bool HasIgnore(this PropertyInfo propertyInfo)
    {
        if (propertyInfo.IsDefined(typeof(NotMappedAttribute)))
            return true;
        if (propertyInfo.IsDefined(typeof(ExcelIgnoreAttribute)))
            return true;
        return false;
    }
}
