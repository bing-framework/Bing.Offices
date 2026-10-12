using Bing.Offices.Exceptions;

namespace Bing.Offices.Entities;

/// <summary>
    /// 规划实体列表的明细与分页内容。
/// </summary>
internal sealed class ExcelEntityRowPlan
{
    /// <summary>
    /// 行段类型。
    /// </summary>
    internal enum SegmentKind
    {
        /// <summary>
        /// 明细行。
        /// </summary>
        Detail,
        /// <summary>
        /// 分组小计。
        /// </summary>
        GroupSubtotal,
        /// <summary>
        /// 分页小计。
        /// </summary>
        PageSubtotal,
        /// <summary>
        /// 后续逻辑页的重复页眉。
        /// </summary>
        PageHeader,
        /// <summary>
        /// 当前逻辑页的尾部。
        /// </summary>
        PageFooter
    }

    /// <summary>
    /// 连续明细或分页内容行段。
    /// </summary>
    internal sealed record Segment(SegmentKind Kind, int Row, int StartIndex, int Count);

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityRowPlan" /> 类型的实例。
    /// </summary>
    private ExcelEntityRowPlan(IReadOnlyList<Segment> segments, IReadOnlyList<int> pageBreakRows,
        IReadOnlyList<(int FirstRow, int LastRow)> detailRows, int finalRow)
    {
        Segments = segments;
        PageBreakRows = pageBreakRows;
        DetailRows = detailRows;
        FinalRow = finalRow;
    }

    /// <summary>
    /// 获取明细与分页内容行段。
    /// </summary>
    internal IReadOnlyList<Segment> Segments { get; }

    /// <summary>
    /// 获取分页符所在的零基行号。
    /// </summary>
    internal IReadOnlyList<int> PageBreakRows { get; }

    /// <summary>
    /// 获取最终尾部公式使用的一基明细行段。
    /// </summary>
    internal IReadOnlyList<(int FirstRow, int LastRow)> DetailRows { get; }

    /// <summary>
    /// 获取最终尾部之前的零基行号。
    /// </summary>
    internal int FinalRow { get; }

    /// <summary>
    /// 根据连续分组和每页明细容量生成行规划。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">实体列表区域。</param>
    /// <param name="values">已物化的明细快照。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>稳定的物理行规划。</returns>
    internal static ExcelEntityRowPlan Create<TEntity>(ExcelEntityListRegion<TEntity> region,
        IReadOnlyList<object> values, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (region.PageFooter != null && !region.PageBreakRows.HasValue)
            throw new BingOfficesConfigurationException("每页尾部必须配置 PageBreak。",
                stage: BingOfficesStage.Plan);
        var segments = new List<Segment>();
        var pageBreakRows = new List<int>();
        var detailRows = new List<(int FirstRow, int LastRow)>();
        var subtotal = region.GroupSubtotal;
        var rowsPerPage = region.PageBreakRows;
        var row = (long)region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var groups = new List<(int Start, int End)>();
        if (values.Count > 0)
        {
            if (subtotal == null)
                groups.Add((0, values.Count));
            else
            {
                var start = 0;
                var key = subtotal.GetKey(values[0]);
                for (var index = 1; index < values.Count; index++)
                {
                    if ((index & 255) == 0)
                        cancellationToken.ThrowIfCancellationRequested();
                    var next = subtotal.GetKey(values[index]);
                    if (Equals(key, next))
                        continue;
                    groups.Add((start, index));
                    start = index;
                    key = next;
                }
                groups.Add((start, values.Count));
            }
        }

        foreach (var group in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var index = group.Start;
            while (index < group.End)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var pageEnd = rowsPerPage.HasValue
                    ? (int)Math.Min(values.Count,
                        ((long)index / rowsPerPage.Value + 1) * rowsPerPage.Value)
                    : values.Count;
                var runEnd = Math.Min(group.End, pageEnd);
                var count = runEnd - index;
                segments.Add(new Segment(SegmentKind.Detail, CheckedRow(row), index, count));
                detailRows.Add((CheckedRow(row + 1), CheckedRow(row + count)));
                row += count;
                index = runEnd;

                if (subtotal != null && index == group.End)
                {
                    segments.Add(new Segment(SegmentKind.GroupSubtotal, CheckedRow(row),
                        group.Start, group.End - group.Start));
                    row += (long)subtotal.Footer.GapRows + GetFooterHeight(subtotal.Footer);
                }
                if (rowsPerPage.HasValue && index == pageEnd && index < values.Count)
                {
                    if (region.PageSubtotal != null)
                    {
                        segments.Add(new Segment(SegmentKind.PageSubtotal, CheckedRow(row),
                            index - rowsPerPage.Value, rowsPerPage.Value));
                        row += (long)region.PageSubtotal.GapRows
                            + GetFooterHeight(region.PageSubtotal);
                    }
                    if (region.PageFooter != null)
                    {
                        segments.Add(new Segment(SegmentKind.PageFooter, CheckedRow(row),
                            index - rowsPerPage.Value, rowsPerPage.Value));
                        row += (long)region.PageFooter.GapRows + GetFooterHeight(region.PageFooter);
                    }
                    pageBreakRows.Add(CheckedRow(row - 1));
                    if (region.PageHeader != null)
                    {
                        segments.Add(new Segment(SegmentKind.PageHeader, CheckedRow(row), index, 0));
                        row += region.PageHeader.RowCount;
                    }
                }
            }
        }

        if (region.PageFooter != null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var count = values.Count == 0 ? 0 : (values.Count - 1) % rowsPerPage.Value + 1;
            segments.Add(new Segment(SegmentKind.PageFooter, CheckedRow(row),
                values.Count - count, count));
            row += (long)region.PageFooter.GapRows + GetFooterHeight(region.PageFooter);
        }

        return new ExcelEntityRowPlan(segments, pageBreakRows, detailRows, CheckedRow(row));
    }

    /// <summary>
    /// 获取尾部占用的行数。
    /// </summary>
    /// <param name="footer">尾部定义。</param>
    /// <returns>尾部占用行数。</returns>
    internal static int GetFooterHeight(IExcelEntityListFooter footer)
    {
        var relativeRow = 0;
        foreach (var cell in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
            relativeRow = Math.Max(relativeRow, cell.Reference.Row);
        foreach (var merge in footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
            relativeRow = Math.Max(relativeRow, merge.Last.Row);
        return relativeRow + 1;
    }

    /// <summary>
    /// 校验行号可用于工作表 API。
    /// </summary>
    /// <param name="row">待校验行号。</param>
    /// <returns>转换后的行号。</returns>
    private static int CheckedRow(long row)
    {
        if (row < 0 || row > int.MaxValue)
            throw new BingOfficesConfigurationException("实体列表布局行号超出支持范围。",
                stage: BingOfficesStage.Plan);
        return (int)row;
    }
}
