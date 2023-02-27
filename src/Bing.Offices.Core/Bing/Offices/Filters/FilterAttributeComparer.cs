using System.Collections.Generic;
using Bing.Offices.Attributes;

namespace Bing.Offices.Filters;

/// <summary>
/// 过滤器特性比较器
/// </summary>
internal class FilterAttributeComparer : IEqualityComparer<FilterAttributeBase>
{
    /// <summary>
    /// 是否相等
    /// </summary>
    /// <param name="x">对象A</param>
    /// <param name="y">对象B</param>
    public bool Equals(FilterAttributeBase x, FilterAttributeBase y) => x.GetType() == y.GetType();

    /// <summary>
    /// 获取哈希编码
    /// </summary>
    /// <param name="obj">对象</param>
    public int GetHashCode(FilterAttributeBase obj) => obj.GetType().GetHashCode();
}