using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 必填
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RequiredAttribute : FilterBaseAttribute
    {
        /// <summary>
        /// 初始化一个<see cref="RequiredAttribute"/>类型的实例
        /// </summary>
        public RequiredAttribute()
        {
            ErrorMessage = "必填";
        }
    }
}
