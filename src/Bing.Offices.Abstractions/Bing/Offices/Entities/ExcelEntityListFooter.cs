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
/// 实体列表区域的导出尾部定义。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityListFooter<TItem> : IExcelEntityListFooter where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityListFooter{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <param name="gapRows">明细与尾部之间的空行数量。</param>
    /// <param name="cells">尾部单元格定义。</param>
    /// <param name="merges">尾部相对合并区域。</param>
    /// <param name="markerStyle">标记单元格样式。</param>
    /// <param name="containsContiguousSum">是否包含连续明细求和公式。</param>
    /// <param name="containsDetailSum">是否包含跨小计明细求和公式。</param>
    internal ExcelEntityListFooter(string markerText, int gapRows,
        IReadOnlyList<ExcelEntityListFooterCell<TItem>> cells,
        IReadOnlyList<ExcelEntityCellRange> merges, ExcelCellStyle markerStyle,
        bool containsContiguousSum, bool containsDetailSum)
    {
        MarkerText = markerText ?? throw new ArgumentNullException(nameof(markerText));
        GapRows = gapRows;
        Cells = cells ?? throw new ArgumentNullException(nameof(cells));
        Merges = merges ?? throw new ArgumentNullException(nameof(merges));
        MarkerStyle = markerStyle;
        ContainsContiguousSum = containsContiguousSum;
        ContainsDetailSum = containsDetailSum;
    }

    /// <inheritdoc />
    public string MarkerText { get; }

    /// <inheritdoc />
    public int GapRows { get; }

    /// <summary>
    /// 获取尾部单元格定义。
    /// </summary>
    public IReadOnlyList<ExcelEntityListFooterCell<TItem>> Cells { get; }

    /// <inheritdoc />
    IReadOnlyList<IExcelEntityListFooterCell> IExcelEntityListFooter.Cells =>
        Cells.Cast<IExcelEntityListFooterCell>().ToArray();

    /// <inheritdoc />
    public IReadOnlyList<ExcelEntityCellRange> Merges { get; }

    /// <inheritdoc />
    public ExcelCellStyle MarkerStyle { get; }

    /// <summary>
    /// 获取尾部是否包含连续明细求和公式。
    /// </summary>
    internal bool ContainsContiguousSum { get; }

    /// <summary>
    /// 获取尾部是否包含跨小计明细求和公式。
    /// </summary>
    internal bool ContainsDetailSum { get; }
}
