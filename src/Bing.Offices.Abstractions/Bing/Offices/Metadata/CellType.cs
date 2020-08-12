namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 单元格类型
    /// </summary>
    public enum CellType
    {
        /// <summary>
        /// 未知
        /// </summary>
        Unknown = -1,
        /// <summary>
        /// 数值
        /// </summary>
        Numeric = 0,
        /// <summary>
        /// 字符串
        /// </summary>
        String = 1,
        /// <summary>
        /// 公式
        /// </summary>
        Formula = 2,
        /// <summary>
        /// 空白
        /// </summary>
        Blank = 3,
        /// <summary>
        /// 布尔值
        /// </summary>
        Boolean = 4,
        /// <summary>
        /// 错误
        /// </summary>
        Error = 5
    }
}
