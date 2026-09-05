using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 行(<see cref="IRow"/>) 扩展
/// </summary>
public static class RowExtensions
{
    /// <summary>
    /// 获取指定索引的单元格；不存在时创建新单元格。
    /// </summary>
    /// <param name="row">行</param>
    /// <param name="cellIndex">单元格索引</param>
    /// <returns>指定索引的现有或新建单元格。</returns>
    public static ICell GetOrCreateCell(this IRow row, int cellIndex) => row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);

    /// <summary>
    /// 获取或创建指定单元格，执行可选配置操作后返回当前行。
    /// </summary>
    /// <param name="row">行</param>
    /// <param name="cellIndex">单元格索引</param>
    /// <param name="action">单元格操作</param>
    /// <returns>当前行。</returns>
    public static IRow CreateCell(this IRow row, int cellIndex, Action<ICell> action)
    {
        var cell = row.GetOrCreateCell(cellIndex);
        action?.Invoke(cell);
        return row;
    }

    /// <summary>
    /// 将当前行已有单元格的内容清空为字符串空值。
    /// </summary>
    /// <param name="row">NPOI单元行</param>
    /// <returns>当前行。</returns>
    public static IRow ClearContent(this IRow row)
    {
        foreach (var cell in row.Cells)
            cell.SetCellValue(string.Empty);
        return row;
    }

    /// <summary>
    /// 判断行不存在或所有已有单元格均为空白。
    /// </summary>
    /// <param name="row">行</param>
    /// <returns>行为空或不存在时为 true。</returns>
    public static bool IsEmptyRow(this IRow row) => row == null || row.Cells.All(x => string.IsNullOrWhiteSpace(x?.GetStringValue()));

    /// <summary>
    /// 将值写入指定列，并在提供样式时绑定该样式。
    /// </summary>
    /// <typeparam name="T">数据类型</typeparam>
    /// <param name="row">行</param>
    /// <param name="column">单元格索引</param>
    /// <param name="value">值</param>
    /// <param name="style">行样式</param>
    public static void Value<T>(this IRow row, int column, T value, ICellStyle style = null)
    {
        var cell = row.CreateCell(column);
        cell.SetValue(value);
        if (style != null)
            cell.CellStyle = style;
    }
}
