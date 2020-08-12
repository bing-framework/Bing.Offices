namespace Bing.Offices.Settings
{
    /// <summary>
    /// 过滤器设置
    /// </summary>
    internal sealed class FilterSetting
    {
        /// <summary>
        /// 首列索引
        /// </summary>
        public int FirstColumn { get; }

        /// <summary>
        /// 最后一列索引
        /// </summary>
        public int? LastColumn { get; }

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
