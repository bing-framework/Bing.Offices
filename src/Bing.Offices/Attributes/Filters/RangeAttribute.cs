using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 数值区间
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class RangeAttribute : FilterBaseAttribute
    {
        /// <summary>
        /// 最小值
        /// </summary>
        public int Min { get; set; }

        /// <summary>
        /// 最大值
        /// </summary>
        public int Max { get; set; }

        /// <summary>
        /// 初始化一个<see cref="RangeAttribute"/>类型的实例
        /// </summary>
        /// <param name="min">最小值</param>
        /// <param name="max">最大值</param>
        public RangeAttribute(int min, int max)
        {
            Min = min;
            Max = max;
            ErrorMessage = $"超出范围，仅允许为{min}-{max}";
        }
    }
}
