using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 区间特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class RangeAttribute : FilterAttributeBase
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
        /// 错误消息
        /// </summary>
        public override string ErrorMsg { get; set; }

        /// <summary>
        /// 初始化一个<see cref="RangeAttribute"/>类型的实例
        /// </summary>
        /// <param name="min">最小值</param>
        /// <param name="max">最大值</param>
        public RangeAttribute(int min, int max)
        {
            Min = min;
            Max = max;
            ErrorMsg= $"超限，仅允许为{Min}-{Max}";
        }
    }
}
