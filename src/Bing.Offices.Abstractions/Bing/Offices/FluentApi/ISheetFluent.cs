namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    public interface ISheetFluent
    {
        /// <summary>
        /// 设置冻结窗格。
        /// 创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都会被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        ISheetFluent HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// 设置冻结窗格。
        /// 创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都会被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        /// <param name="leftMostColumn">可见列</param>
        /// <param name="topRow">可见行</param>
        ISheetFluent HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow);

        /// <summary>
        /// 设置筛选器
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        ISheetFluent HasFilter(int firstColumn);

        /// <summary>
        /// 设置筛选器
        /// </summary>
        /// <param name="fistColumn">首列索引</param>
        /// <param name="lastColumn">尾列索引</param>
        ISheetFluent HasFilter(int fistColumn, int? lastColumn);
    }
}
