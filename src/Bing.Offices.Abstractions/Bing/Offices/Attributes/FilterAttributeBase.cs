using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// 过滤器特性基类
/// </summary>
public abstract class FilterAttributeBase : Attribute
{
    /// <summary>
    /// 错误消息
    /// </summary>
    public virtual string ErrorMsg { get; set; } = "非法";
}