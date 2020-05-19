namespace Bing.Offices.Filters
{
    /// <summary>
    /// 过滤器上下文
    /// </summary>
    public class FilterContext : IFilterContext
    {
        /// <summary>
        /// 类型过滤器信息
        /// </summary>
        public TypeFilterInfo TypeFilterInfo { get; set; }
    }
}
