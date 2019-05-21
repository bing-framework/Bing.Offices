using System.Collections.Generic;

namespace Bing.Offices.Filters
{
    /// <summary>
    /// 类型过滤器信息
    /// </summary>
    public class TypeFilterInfo
    {
        /// <summary>
        /// 属性过滤器列表
        /// </summary>
        public List<PropertyFilterInfo> PropertyFilters { get; set; }
    }
}
