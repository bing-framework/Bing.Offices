using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// 绑定过滤器特性
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class BindFilterAttribute : Attribute
{
    /// <summary>
    /// 过滤器类型
    /// </summary>
    public Type FilterType { get; set; }

    /// <summary>
    /// 初始化一个<see cref="BindFilterAttribute"/>类型的实例
    /// </summary>
    /// <param name="filterType">过滤器类型</param>
    public BindFilterAttribute(Type filterType)
    {
        if (!filterType.IsSubclassOf(typeof(FilterAttributeBase)))
            throw new ArgumentOutOfRangeException(nameof(filterType),
                $@"{filterType.Name} 不是 {nameof(FilterAttributeBase)}的子类");
        FilterType = filterType;
    }
}