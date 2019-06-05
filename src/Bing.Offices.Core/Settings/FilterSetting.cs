using Bing.Offices.Abstractions.Settings;

namespace Bing.Offices.Settings
{
    /// <summary>
    /// 过滤器设置
    /// </summary>
    public sealed class FilterSetting : IFilterSetting
    {
        /// <summary>
        /// 首列索引
        /// </summary>
        public int FirstColumn { get; set; }

        /// <summary>
        /// 最后一列索引
        /// </summary>
        public int? LastColumn { get; set; }

        /// <summary>
        /// 初始化一个<see cref="FilterSetting"/>类型的实例
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        public FilterSetting(int firstColumn) : this(firstColumn, null) { }

        /// <summary>
        /// 初始化一个<see cref="FilterSetting"/>类型的实例
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        /// <param name="lastColumn">最后一列索引</param>
        public FilterSetting(int firstColumn, int? lastColumn)
        {
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }
    }
}
