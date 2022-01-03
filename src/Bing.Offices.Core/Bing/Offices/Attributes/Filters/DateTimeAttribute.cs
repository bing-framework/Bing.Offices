using System;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 日期特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DateTimeAttribute : FilterAttributeBase
    {
        /// <summary>
        /// 错误消息
        /// </summary>
        public override string ErrorMsg { get; set; } = "非日期数据";
    }
}
