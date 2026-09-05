using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 工作表行、合并区域和图片相关操作扩展。
/// </summary>
public static partial class SheetExtensions
{
    /// <summary>
    /// 获取工作表全部合并区域的零基边界，列顺序为起始行、起始列、结束行、结束列。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <returns>每行表示一个合并区域的四列整数矩阵；没有合并区域时返回零行矩阵。</returns>
    public static int[,] GetAllMergedRegions(this ISheet sheet)
    {
        // 工作表合并单元格数量
        var mergedRegions = sheet.NumMergedRegions;
        var output = new int[mergedRegions, 4];
        for (var i = 0; i < mergedRegions; i++)
        {
            var cellRangeAddress = sheet.GetMergedRegion(i);
            output[i, 0] = cellRangeAddress.FirstRow;
            output[i, 1] = cellRangeAddress.FirstColumn;
            output[i, 2] = cellRangeAddress.LastRow;
            output[i, 3] = cellRangeAddress.LastColumn;
        }
        return output;
    }

    /// <summary>
    /// 删除指定数量的行，并先移除受影响的合并区域以避免移动行时产生格式错乱。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="deleteRowStartIndex">删除起始行的零基索引。</param>
    /// <param name="count">要删除的行数，必须大于零。</param>
    public static void DeleteRows(this ISheet sheet, int deleteRowStartIndex, int count)
    {
        if (deleteRowStartIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(deleteRowStartIndex));
        if (count <= 0)
            throw new ArgumentOutOfRangeException(nameof(count));
        var mergedRegions = sheet.NumMergedRegions;
        for (var i = mergedRegions - 1; i >= 0; i--)
        {
            var cellRangeAddress = sheet.GetMergedRegion(i);
            if (cellRangeAddress.LastRow >= deleteRowStartIndex &&
                cellRangeAddress.FirstRow < deleteRowStartIndex + count)
                sheet.RemoveMergedRegion(i);
        }
        if (deleteRowStartIndex + count <= sheet.LastRowNum)
            sheet.ShiftRows(deleteRowStartIndex + count, sheet.LastRowNum, -count, true, false);
    }

    /// <summary>
    /// 获取最后一个非空数据行的零基索引；没有数据行时返回 -1。
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <returns>最后一个非空行的零基索引，或 -1。</returns>
    public static int GetHasDataRowNum(this ISheet sheet)
    {
        for (var i = sheet.LastRowNum; i >= 0; i--)
        {
            if (!sheet.GetRow(i).IsEmptyRow())
                return i;
        }
        return -1;
    }

    /// <summary>
    /// 从第一行数据位置开始按顺序创建行，并为每项执行行配置操作。
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="sheet">工作表</param>
    /// <param name="data">数据</param>
    /// <param name="action">操作</param>
    public static void AddRow<T>(this ISheet sheet, IEnumerable<T> data, Action<IRow, T> action)
    {
        var index = 1;
        foreach (var item in data)
        {
            var row = sheet.CreateRow(index);
            row.Height = 20 * 20;
            action(row, item);
            index++;
        }
    }

    /// <summary>
    /// 在指定零基索引处插入一行，并移动后续行。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="rowIndex">行索引</param>
    /// <returns>新建的行；底层工作表未返回时为 null。</returns>
    public static NPOI.SS.UserModel.IRow InsertRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.InsertRows(rowIndex, 1).FirstOrDefault();

    /// <summary>
    /// 在指定零基索引处插入多行，并移动后续行。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="rowIndex">行索引</param>
    /// <param name="rowsCount">插入行数量</param>
    /// <returns>按插入顺序排列的新建行数组。</returns>
    public static NPOI.SS.UserModel.IRow[] InsertRows(this NPOI.SS.UserModel.ISheet sheet, int rowIndex,
        int rowsCount)
    {
        if (rowIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));
        if (rowsCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(rowsCount));
        if (rowIndex <= sheet.LastRowNum)
            sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowsCount, true, false);
        var rows = new List<NPOI.SS.UserModel.IRow>();
        for (var i = 0; i < rowsCount; i++)
        {
            var row = sheet.CreateRow(rowIndex + i);
            rows.Add(row);
        }
        return rows.ToArray();
    }

    /// <summary>
    /// 移除指定零基索引的一行，并同步处理受影响的合并区域和图片锚点。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="rowIndex">行索引</param>
    /// <returns>实际移除的行数，通常为 1。</returns>
    public static int RemoveRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.RemoveRows(rowIndex, rowIndex);

    /// <summary>
    /// 移除指定闭合零基行区间，并同步处理合并区域和图片锚点。
    /// </summary>
    /// <param name="sheet">NPOI工作表</param>
    /// <param name="startRowIndex">起始行索引</param>
    /// <param name="endRowIndex">结束行索引</param>
    /// <returns>实际移除的行数。</returns>
    public static int RemoveRows(this NPOI.SS.UserModel.ISheet sheet, int startRowIndex, int endRowIndex)
    {
        if (startRowIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(startRowIndex));
        if (endRowIndex < startRowIndex)
            throw new ArgumentOutOfRangeException(nameof(endRowIndex));
        var span = endRowIndex - startRowIndex + 1;
        sheet.RemoveMergedRegions(startRowIndex, endRowIndex, null, null);
        sheet.RemovePictures(startRowIndex, endRowIndex, null, null, onlyInternal: false);
        for (var i = endRowIndex; i >= startRowIndex; i--)
        {
            var row = sheet.GetRow(i);
            if (row != null)
                sheet.RemoveRow(row);
        }
        if (endRowIndex + 1 <= sheet.LastRowNum)
        {
            sheet.ShiftRows(endRowIndex + 1, sheet.LastRowNum, -span, true, false);
            sheet.MovePictures(endRowIndex + 1, null, null, null, onlyInternal: false, moveRowCount: -span);
        }
        return span;
    }

    /// <summary>
    /// 判断指定区域是否在内部或交叉
    /// </summary>
    /// <param name="rangeMinRow">区域最小行索引</param>
    /// <param name="rangeMaxRow">区域最大行索引</param>
    /// <param name="rangeMinCol">区域最小列索引</param>
    /// <param name="rangeMaxCol">区域最大列索引</param>
    /// <param name="targetMinRow">目标最小行索引</param>
    /// <param name="targetMaxRow">目标最大行索引</param>
    /// <param name="targetMinCol">目标最小列索引</param>
    /// <param name="targetMaxCol">目标最大列索引</param>
    /// <param name="onlyInternal">仅在内部</param>
    private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol,
        int? rangeMaxCol, int targetMinRow, int targetMaxRow, int targetMinCol, int targetMaxCol, bool onlyInternal)
    {
        var tempMinRow = rangeMinRow ?? targetMinRow;
        var tempMaxRow = rangeMaxRow ?? targetMaxRow;
        var tempMinCol = rangeMinCol ?? targetMinCol;
        var tempMaxCol = rangeMaxCol ?? targetMaxCol;
        if (onlyInternal)
        {
            return tempMinRow <= targetMinRow &&
                   tempMaxRow >= targetMaxRow &&
                   tempMinCol <= targetMinCol &&
                   tempMaxCol >= targetMaxCol;
        }

        return Math.Abs(tempMaxRow - tempMinRow) + Math.Abs(targetMaxRow - targetMinRow) >=
               Math.Abs(tempMaxRow + tempMinRow - targetMaxRow - targetMinRow) &&
               Math.Abs(tempMaxCol - tempMinCol) + Math.Abs(targetMaxCol - targetMinCol) >=
               Math.Abs(tempMaxCol + tempMinCol - targetMaxCol - targetMinCol);
    }
}
