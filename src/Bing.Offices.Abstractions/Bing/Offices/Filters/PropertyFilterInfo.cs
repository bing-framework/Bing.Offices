﻿using Bing.Offices.Attributes;

namespace Bing.Offices.Filters;

/// <summary>
/// 属性过滤器信息
/// </summary>
public class PropertyFilterInfo
{
    /// <summary>
    /// 属性名
    /// </summary>
    public string PropertyName { get; set; }

    /// <summary>
    /// 过滤器列表
    /// </summary>
    public IList<FilterAttributeBase> Filters { get; set; } = new List<FilterAttributeBase>();
}
