using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using Bing.Offices.Attributes;

namespace Bing.Offices.Extensions;

/// <summary>
/// 提供实体属性 Excel 映射元数据的判断扩展。
/// </summary>
internal static class PropertyInfoExtensions
{
    /// <summary>
    /// 判断属性是否声明为不参与 Excel 映射。
    /// </summary>
    /// <param name="propertyInfo">要检查映射特性的属性元数据。</param>
    /// <returns>属性声明 <see cref="NotMappedAttribute" /> 或 <see cref="ExcelIgnoreAttribute" /> 时为 <see langword="true" />；否则为 <see langword="false" />。</returns>
    internal static bool HasIgnore(this PropertyInfo propertyInfo)
    {
        if (propertyInfo.IsDefined(typeof(NotMappedAttribute)))
            return true;
        if (propertyInfo.IsDefined(typeof(ExcelIgnoreAttribute)))
            return true;
        return false;
    }
}
