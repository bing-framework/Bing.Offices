namespace Bing.Offices.Excel.Settings
{
    /// <summary>
    /// 冻结窗格设置
    /// </summary>
    internal sealed class FreezeSetting
    {
        /// <summary>
        /// 初始化一个<see cref="FreezeSetting"/>类型的实例
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        public FreezeSetting(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1) { }

        /// <summary>
        /// 初始化一个<see cref="FreezeSetting"/>类型的实例
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        /// <param name="leftMostColumn">顶列索引</param>
        /// <param name="topRow">顶行索引</param>
        public FreezeSetting(int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            ColSplit = colSplit;
            RowSplit = rowSplit;
            LeftMostColumn = leftMostColumn;
            TopRow = topRow;
        }

        /// <summary>
        /// 水平分割位置
        /// </summary>
        public int ColSplit { get; }

        /// <summary>
        /// 垂直分割位置
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
    }
}
