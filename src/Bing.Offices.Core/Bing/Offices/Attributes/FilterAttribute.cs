using System;
using Bing.Offices.Settings;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 过滤器 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class,AllowMultiple = true)]
    public sealed class FilterAttribute:Attribute
    {
        /// <summary>
        /// 过滤器设置
        /// </summary>
        internal FilterSetting FilterSetting { get; }

        /// <summary>
        /// 首列索引
        /// </summary>
        public int FirstColumn => FilterSetting.FirstColumn;

        /// <summary>
        /// 最后一列索引
        /// </summary>
        public int? LastColumn => FilterSetting.LastColumn;

        /// <summary>
        /// 初始化一个<see cref="FilterAttribute"/>类型的实例
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        public FilterAttribute(int firstColumn) : this(firstColumn, null) { }

        /// <summary>
        /// 初始化一个<see cref="FilterAttribute"/>类型的实例
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        /// <param name="lastColumn">最后一列索引</param>
        public FilterAttribute(int firstColumn, int? lastColumn) => FilterSetting = new FilterSetting(firstColumn, lastColumn);
    }
}
