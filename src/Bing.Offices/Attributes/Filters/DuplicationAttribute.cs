using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 模板数据重复数据校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DuplicationAttribute : FilterBaseAttribute
    {
        /// <summary>
        /// 初始化一个<see cref="DuplicationAttribute"/>类型的实例
        /// </summary>
        public DuplicationAttribute()
        {
            ErrorMessage = "存在重复数据";
        }
    }
}
