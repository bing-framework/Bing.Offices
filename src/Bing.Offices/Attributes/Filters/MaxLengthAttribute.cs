using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 最大长度
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MaxLengthAttribute : FilterBaseAttribute
    {
        /// <summary>
        /// 最大长度
        /// </summary>
        public int MaxLength { get; set; }

        /// <summary>
        /// 初始化一个<see cref="MaxLengthAttribute"/>类型的实例
        /// </summary>
        /// <param name="maxLength">最大长度</param>
        public MaxLengthAttribute(int maxLength)
        {
            MaxLength = maxLength;
            ErrorMessage = "超长";
        }
    }
}
