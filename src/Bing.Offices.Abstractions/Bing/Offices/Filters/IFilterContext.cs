namespace Bing.Offices.Filters;

/// <summary>
/// 过滤器上下文
/// </summary>
public interface IFilterContext
{
    /// <summary>
    /// 类型过滤器信息
    /// </summary>
    TypeFilterInfo TypeFilterInfo { get; set; }
}