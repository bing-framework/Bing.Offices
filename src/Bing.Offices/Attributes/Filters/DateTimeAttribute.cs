using System;
using System.Collections.Generic;
using System.Text;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 日期
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DateTimeAttribute : FilterBaseAttribute
    {
        /// <summary>
        /// 初始化一个<see cref="DateTimeAttribute"/>类型的实例
        /// </summary>
        public DateTimeAttribute()
        {
            ErrorMessage = "无效日期";
        }
    }
}
