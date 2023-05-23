namespace Bing.Offices.Settings;

/// <summary>
/// 冻结设置，表示指定实体的Excel冻结设置
/// </summary>
public sealed class FreezeSetting
{
    /// <summary>
    /// 初始化一个<see cref="FreezeSetting"/>类型的实例
    /// </summary>
    /// <param name="colSplit">冻结的列数</param>
    /// <param name="rowSplit">冻结的行数</param>
    public FreezeSetting(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1)
    {
    }

    /// <summary>
    /// 初始化一个<see cref="FreezeSetting"/>类型的实例
    /// </summary>
    /// <param name="colSplit">冻结的列数</param>
    /// <param name="rowSplit">冻结的行数</param>
    /// <param name="leftMostColumn">冻结后第一列的列号</param>
    /// <param name="topRow">冻结后第一行的行号</param>
    public FreezeSetting(int colSplit, int rowSplit, int leftMostColumn, int topRow)
    {
        ColSplit = colSplit;
        RowSplit = rowSplit;
        LeftMostColumn = leftMostColumn;
        TopRow = topRow;
    }

    /// <summary>
    /// 冻结的列数
    /// </summary>
    public int ColSplit { get; }

    /// <summary>
    /// 冻结的行数
    /// </summary>
    public int RowSplit { get; }

    /// <summary>
    /// 冻结后第一列的列号
    /// </summary>
    public int LeftMostColumn { get; }

    /// <summary>
    /// 冻结后第一行的行号
    /// </summary>
    public int TopRow { get; }
}
