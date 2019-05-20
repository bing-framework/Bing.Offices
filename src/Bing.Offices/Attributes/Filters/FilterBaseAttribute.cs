using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 过滤器属性基类
    /// </summary>
    public class FilterBaseAttribute : Attribute
    {
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; } = "非法";
    }
}
