using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 过滤器绑定
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class FilterBindAttribute : Attribute
    {
        /// <summary>
        /// 过滤器类型
        /// </summary>
        public Type FilterType { get; set; }

        /// <summary>
        /// 初始化一个<see cref="FilterBindAttribute"/>类型的实例
        /// </summary>
        /// <param name="filterType">过滤器类型</param>
        public FilterBindAttribute(Type filterType)
        {
            if (!filterType.IsSubclassOf(typeof(FilterBaseAttribute)))
            {
                throw new ArgumentOutOfRangeException($"{filterType.ToString()} 不是 FilterBaseAttribute 子类");
            }

            FilterType = filterType;
        }
    }
}
