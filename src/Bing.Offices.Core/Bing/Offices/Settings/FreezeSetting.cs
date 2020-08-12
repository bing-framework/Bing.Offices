namespace Bing.Offices.Settings
{
    /// <summary>
    /// 冻结设置
    /// </summary>
    internal sealed class FreezeSetting
    {
        /// <summary>
        /// 垂直分割位置
        /// </summary>
        public int ColSplit { get; }

        /// <summary>
        /// 水平分割位置
        /// </summary>
        public int RowSplit { get; }

        /// <summary>
        /// 可见列数
        /// </summary>
        public int LeftMostColumn { get; }

        /// <summary>
        /// 可见行数
        /// </summary>
        public int TopRow { get; }

        /// <summary>
        /// 初始化一个<see cref="FreezeSetting"/>类型的实例
        /// </summary>
        /// <param name="colSplit">垂直分割位置</param>
        /// <param name="rowSplit">水平分割位置</param>
        public FreezeSetting(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1)
        {
        }

        /// <summary>
        /// 初始化一个<see cref="FreezeSetting"/>类型的实例
        /// </summary>
        /// <param name="colSplit">垂直分割位置</param>
        /// <param name="rowSplit">水平分割位置</param>
        /// <param name="leftMostColumn">可见列数</param>
        /// <param name="topRow">可见行数</param>
        public FreezeSetting(int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            ColSplit = colSplit;
            RowSplit = rowSplit;
            LeftMostColumn = leftMostColumn;
            TopRow = topRow;
        }

        /// <summary>
        /// 冻结首行
        /// </summary>
        public static FreezeSetting FirstRow = new FreezeSetting(0, 1, 0, 1);

        /// <summary>
        /// 冻结首列
        /// </summary>
        public static FreezeSetting FirstColumn = new FreezeSetting(1, 0, 1, 0);

        /// <summary>
        /// 冻结首行首列
        /// </summary>
        public static FreezeSetting FirstRowWithColumn = new FreezeSetting(1, 1, 1, 1);
    }
}
