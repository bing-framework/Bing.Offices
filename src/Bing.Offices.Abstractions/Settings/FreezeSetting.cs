namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 冻结设置
    /// </summary>
    public sealed class FreezeSetting
    {
        /// <summary>
        /// 冻结单元格的列号
        /// </summary>
        public int ColSplit { get; }

        /// <summary>
        /// 冻结单元格的行号
        /// </summary>
        public int RowSplit { get; }

        /// <summary>
        /// 顶列索引
        /// </summary>
        public int LeftMostColumn { get; }

        /// <summary>
        /// 顶行索引
        /// </summary>
        public int TopRow { get; }

        /// <summary>
        /// 初始化一个<see cref="FreezeSetting"/>类型的实例
        /// </summary>
        /// <param name="colSplit">冻结单元格的列号</param>
        /// <param name="rowSplit">冻结单元格的行号</param>
        public FreezeSetting(int colSplit,int rowSplit) : this(colSplit, rowSplit, 0, 1) { }

        /// <summary>
        /// 初始化一个<see cref="FreezeSetting"/>类型的实例
        /// </summary>
        /// <param name="colSplit">冻结单元格的列号</param>
        /// <param name="rowSplit">冻结单元格的行号</param>
        /// <param name="leftMostColumn">顶列索引</param>
        /// <param name="topRow">顶行索引</param>
        public FreezeSetting(int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            ColSplit = colSplit;
            RowSplit = rowSplit;
            LeftMostColumn = leftMostColumn;
            TopRow = topRow;
        }
    }
}
