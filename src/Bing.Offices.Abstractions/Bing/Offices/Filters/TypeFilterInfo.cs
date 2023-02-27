using System.Collections.Generic;

namespace Bing.Offices.Filters;

/// <summary>
/// 类型过滤器信息
/// </summary>
public class TypeFilterInfo
{
    /// <summary>
    /// 属性过滤器信息列表
    /// </summary>
    public IList<PropertyFilterInfo> PropertyFilterInfos { get; set; } = new List<PropertyFilterInfo>();
}