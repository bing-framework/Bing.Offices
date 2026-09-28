using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Bing.Offices.Exports;
using Bing.Offices.Styles;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体列表计算列读取当前行时使用的上下文。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityRowContext<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityRowContext{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <param name="items">本次导出的明细快照。</param>
    /// <param name="index">当前列表项的零基索引。</param>
    /// <param name="rowIndex">当前工作表行的零基索引。</param>
    /// <param name="columnIndex">当前工作表列的零基索引。</param>
    /// <param name="sheetName">当前工作表名称。</param>
    /// <param name="columnKey">当前计算列标识。</param>
    internal ExcelEntityRowContext(TItem item, IReadOnlyList<TItem> items, int index,
        int rowIndex, int columnIndex, string sheetName, string columnKey)
    {
        Item = item ?? throw new ArgumentNullException(nameof(item));
        Items = items ?? throw new ArgumentNullException(nameof(items));
        Index = index;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        ColumnKey = columnKey ?? throw new ArgumentNullException(nameof(columnKey));
    }

    /// <summary>
    /// 获取当前列表项。
    /// </summary>
    public TItem Item { get; }

    /// <summary>
    /// 获取本次导出的明细快照。
    /// </summary>
    public IReadOnlyList<TItem> Items { get; }

    /// <summary>
    /// 获取当前列表项的零基索引。
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// 获取当前工作表行的零基索引。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// 获取当前工作表行的一基行号。
    /// </summary>
    public int RowNumber => RowIndex + 1;

    /// <summary>
    /// 获取当前工作表列的零基索引。
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// 获取当前工作表列的一基列号。
    /// </summary>
    public int ColumnNumber => ColumnIndex + 1;

    /// <summary>
    /// 获取当前工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取当前计算列标识。
    /// </summary>
    public string ColumnKey { get; }
}
