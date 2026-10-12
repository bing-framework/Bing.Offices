using Bing.Offices.Exceptions;

namespace Bing.Offices.Entities;

/// <summary>
/// 识别实体分页布局的辅助行段。
/// </summary>
internal static class ExcelEntitySubtotalImportScanner
{
    /// <summary>
    /// 辅助行段的标记类别。
    /// </summary>
    private enum MarkerKind { Group, Page, Header, PageFooter }

    /// <summary>标记及其占用的物理行。</summary>
    private sealed record Marker(int StartRow, int EndRow, MarkerKind Kind,
        int Row, int Column);

    /// <summary>
    /// 校验输出布局的保留标记。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="firstRow">首个明细的零基行号。</param>
    /// <param name="endRowExclusive">最终尾部之前的零基行号。</param>
    /// <param name="readMarker">读取规范化列表标记的委托。</param>
    /// <param name="readPageHeaderMarker">读取规范化页眉标记的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <remarks>
    /// 读取已写入的单元格，避免再次执行值工厂；仅检查实际用于导入识别的标记列。
    /// </remarks>
    internal static void ValidateWrittenMarkers<TEntity>(ExcelEntityListRegion<TEntity> region,
        int firstRow, int endRowExclusive, Func<int, string> readMarker,
        Func<int, string> readPageHeaderMarker, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (region.PageFooter == null && region.PageHeaderAnchorPrefix == null)
            return;
        if (region.PageFooter != null && region.PageHeader != null)
            for (var row = region.PageHeader.SourceRange.First.Row;
                 row <= region.PageHeader.SourceRange.Last.Row; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var text = readMarker(row);
                if (string.Equals(text, region.PageFooter.MarkerText, StringComparison.Ordinal)
                    || region.PageSubtotal != null && string.Equals(text,
                        region.PageSubtotal.MarkerText, StringComparison.Ordinal)
                    || region.GroupSubtotal != null && string.Equals(text,
                        region.GroupSubtotal.Footer.MarkerText, StringComparison.Ordinal))
                    throw InvalidMarker(region, row, region.Start.Column);
            }
        FindSpans(region, firstRow, endRowExclusive, readMarker, readPageHeaderMarker,
            cancellationToken);
    }

    /// <summary>
    /// 补齐末页尾部的空白占用行。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="firstRow">首个明细行。</param>
    /// <param name="lastContentRow">区域中最后一个内容行。</param>
    /// <param name="readMarker">读取规范化标记文本的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含尾部完整高度的零基末行。</returns>
    internal static int CompletePageFooterEnd<TEntity>(ExcelEntityListRegion<TEntity> region,
        int firstRow, int lastContentRow, Func<int, string> readMarker,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (region.PageFooter == null)
            return lastContentRow;
        var height = ExcelEntityRowPlan.GetFooterHeight(region.PageFooter);
        for (var row = lastContentRow; row >= Math.Max(firstRow, lastContentRow - height + 1); row--)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(readMarker(row), region.PageFooter.MarkerText, StringComparison.Ordinal))
                continue;
            var end = (long)row + height - 1;
            if (end > int.MaxValue || region.End.HasValue && end > region.End.Value.Row)
                throw InvalidMarker(region, row, region.Start.Column);
            return Math.Max(lastContentRow, (int)end);
        }
        return lastContentRow;
    }

    /// <summary>
    /// 识别并校验实体分页辅助行段。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="firstRow">首个明细的零基行号。</param>
    /// <param name="endRowExclusive">最终尾部之前的零基行号。</param>
    /// <param name="readMarker">读取规范化标记文本的委托。</param>
    /// <param name="readPageHeaderMarker">读取重复页眉标记文本的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>导入时应跳过的辅助行段。</returns>
    internal static IReadOnlyList<(int StartRow, int EndRow)> FindSpans<TEntity>(
        ExcelEntityListRegion<TEntity> region, int firstRow, int endRowExclusive,
        Func<int, string> readMarker, Func<int, string> readPageHeaderMarker,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        cancellationToken.ThrowIfCancellationRequested();
        var markers = new List<Marker>();
        if (region.PageHeader != null && !string.Equals(
                readPageHeaderMarker?.Invoke(region.PageHeader.Marker.Row),
                region.PageHeader.MarkerText, StringComparison.Ordinal))
            throw InvalidMarker(region, region.PageHeader.Marker.Row,
                region.PageHeader.Marker.Column);
        if (region.PageHeader != null)
            for (var row = region.PageHeader.SourceRange.First.Row;
                 row <= region.PageHeader.SourceRange.Last.Row; row++)
            {
                if (((row - region.PageHeader.SourceRange.First.Row) & 255) == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                if (row != region.PageHeader.Marker.Row && string.Equals(
                        readPageHeaderMarker?.Invoke(row), region.PageHeader.MarkerText,
                        StringComparison.Ordinal))
                    throw InvalidMarker(region, row, region.PageHeader.Marker.Column);
            }
        for (var row = firstRow; row < endRowExclusive; row++)
        {
            if ((row - firstRow & 255) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var text = readMarker(row);
            var isGroup = region.GroupSubtotal != null && string.Equals(text,
                region.GroupSubtotal.Footer.MarkerText,
                StringComparison.Ordinal);
            var isPage = region.PageSubtotal != null && string.Equals(text,
                region.PageSubtotal.MarkerText, StringComparison.Ordinal);
            var isHeader = region.PageHeader != null && string.Equals(readPageHeaderMarker?.Invoke(row),
                region.PageHeader.MarkerText, StringComparison.Ordinal);
            var isFooter = region.PageFooter != null && string.Equals(text,
                region.PageFooter.MarkerText, StringComparison.Ordinal);
            if ((isGroup ? 1 : 0) + (isPage ? 1 : 0) + (isHeader ? 1 : 0)
                + (isFooter ? 1 : 0) > 1)
                throw InvalidMarker(region, row, isHeader
                    ? region.PageHeader.Marker.Column : region.Start.Column);
            if (!isGroup && !isPage && !isHeader && !isFooter)
                continue;
            var footer = isFooter ? region.PageFooter : isPage ? region.PageSubtotal
                : isGroup ? region.GroupSubtotal.Footer : null;
            var marker = isHeader
                ? new Marker(row - (region.PageHeader.Marker.Row
                    - region.PageHeader.SourceRange.First.Row),
                    row - (region.PageHeader.Marker.Row
                        - region.PageHeader.SourceRange.First.Row) + region.PageHeader.RowCount - 1,
                    MarkerKind.Header, row, region.PageHeader.Marker.Column)
                : new Marker(row - footer.GapRows,
                    row + ExcelEntityRowPlan.GetFooterHeight(footer) - 1,
                    isFooter ? MarkerKind.PageFooter : isPage ? MarkerKind.Page : MarkerKind.Group,
                    row, region.Start.Column);
            if (marker.StartRow < firstRow || marker.EndRow >= endRowExclusive
                || markers.Count > 0 && marker.StartRow <= markers[markers.Count - 1].EndRow)
                throw InvalidMarker(region, row, marker.Column);
            markers.Add(marker);
        }

        if (region.PageFooter != null)
        {
            ValidatePageFooters(region, markers, firstRow, endRowExclusive, cancellationToken);
            return markers.Select(marker => (marker.StartRow, marker.EndRow)).ToArray();
        }

        var cursor = firstRow;
        var pageDetails = 0;
        var groupDetails = 0;
        var sawPageMarker = false;
        var sawHeader = false;
        var awaitingHeader = false;
        var previousWasPage = false;
        int? expectedHeaderStart = null;
        foreach (var marker in markers)
        {
            if (awaitingHeader && marker.Kind != MarkerKind.Header)
                throw InvalidMarker(region, cursor + region.PageHeader.Marker.Row
                    - region.PageHeader.SourceRange.First.Row,
                    region.PageHeader.Marker.Column);
            var details = marker.StartRow - cursor;
            if (region.PageHeader != null && region.PageSubtotal == null
                && !expectedHeaderStart.HasValue && pageDetails < region.PageBreakRows
                && pageDetails + details >= region.PageBreakRows)
                expectedHeaderStart = pageDetails + details == region.PageBreakRows
                    && marker.Kind == MarkerKind.Group
                    ? marker.EndRow + 1 : cursor + region.PageBreakRows - pageDetails;
            pageDetails += details;
            groupDetails += details;
            if (marker.Kind == MarkerKind.Page)
            {
                if (pageDetails != region.PageBreakRows)
                    throw InvalidMarker(region, marker.Row, marker.Column);
                pageDetails = 0;
                sawPageMarker = true;
                awaitingHeader = region.PageHeader != null;
            }
            else if (marker.Kind == MarkerKind.Group)
            {
                if (groupDetails == 0 || previousWasPage && details == 0)
                    throw InvalidMarker(region, marker.Row, marker.Column);
                groupDetails = 0;
            }
            else
            {
                if (region.PageSubtotal == null && pageDetails != region.PageBreakRows
                    || region.PageSubtotal != null && (!awaitingHeader || pageDetails != 0))
                    throw InvalidMarker(region, marker.Row, marker.Column);
                pageDetails = 0;
                sawHeader = true;
                awaitingHeader = false;
                expectedHeaderStart = null;
            }
            previousWasPage = marker.Kind == MarkerKind.Page;
            cursor = marker.EndRow + 1;
        }
        var trailingDetails = endRowExclusive - cursor;
        if (region.PageHeader != null && region.PageSubtotal == null
            && !expectedHeaderStart.HasValue && pageDetails <= region.PageBreakRows
            && pageDetails + trailingDetails > region.PageBreakRows)
            expectedHeaderStart = cursor + region.PageBreakRows - pageDetails;
        pageDetails += trailingDetails;
        if (awaitingHeader || region.PageSubtotal != null && (pageDetails > region.PageBreakRows
            || sawPageMarker && pageDetails == 0)
            || region.PageHeader != null && (pageDetails > region.PageBreakRows
                || sawHeader && pageDetails == 0))
            throw InvalidMarker(region, awaitingHeader
                ? cursor + region.PageHeader.Marker.Row - region.PageHeader.SourceRange.First.Row
                : expectedHeaderStart.HasValue && pageDetails > region.PageBreakRows
                    ? expectedHeaderStart.Value + region.PageHeader.Marker.Row
                        - region.PageHeader.SourceRange.First.Row
                    : endRowExclusive);
        return markers.Select(marker => (marker.StartRow, marker.EndRow)).ToArray();
    }

    /// <summary>
    /// 校验每页尾部与小计、重复页眉的物理顺序。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="markers">按物理行排序的标记。</param>
    /// <param name="firstRow">首个明细行。</param>
    /// <param name="endRowExclusive">最终尾部之前的行号。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ValidatePageFooters<TEntity>(ExcelEntityListRegion<TEntity> region,
        IReadOnlyList<Marker> markers, int firstRow, int endRowExclusive,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var capacity = region.PageBreakRows.GetValueOrDefault();
        if (capacity <= 0)
            throw InvalidMarker(region, firstRow, region.Start.Column);
        var cursor = firstRow;
        var pageDetails = 0;
        var groupDetails = 0;
        var footerCount = 0;
        var subtotalSeen = false;
        var awaitingHeader = false;
        var finished = false;
        foreach (var marker in markers)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var details = marker.StartRow - cursor;
            if (finished)
                throw InvalidMarker(region, marker.Row, marker.Column);
            if (awaitingHeader)
            {
                if (marker.Kind != MarkerKind.Header || details != 0)
                    throw InvalidMarker(region, cursor + region.PageHeader.Marker.Row
                        - region.PageHeader.SourceRange.First.Row, region.PageHeader.Marker.Column);
                awaitingHeader = false;
                cursor = marker.EndRow + 1;
                continue;
            }
            if (subtotalSeen && (marker.Kind != MarkerKind.PageFooter || details != 0))
                throw InvalidMarker(region, cursor + region.PageFooter.GapRows, region.Start.Column);
            if ((long)pageDetails + details > capacity)
                throw InvalidMarker(region, cursor + capacity - pageDetails
                    + region.PageFooter.GapRows, region.Start.Column);
            pageDetails += details;
            groupDetails += details;
            switch (marker.Kind)
            {
                case MarkerKind.Group:
                    if (groupDetails == 0)
                        throw InvalidMarker(region, marker.Row, marker.Column);
                    groupDetails = 0;
                    break;
                case MarkerKind.Page:
                    if (pageDetails != capacity || subtotalSeen)
                        throw InvalidMarker(region, marker.Row, marker.Column);
                    subtotalSeen = true;
                    break;
                case MarkerKind.Header:
                    throw InvalidMarker(region, cursor + region.PageFooter.GapRows,
                        region.Start.Column);
                case MarkerKind.PageFooter:
                    var last = marker.EndRow + 1 == endRowExclusive;
                    if (pageDetails == 0 && (footerCount > 0 || !last)
                        || !last && pageDetails != capacity
                        || last && subtotalSeen
                        || !last && region.PageSubtotal != null && !subtotalSeen
                        || last && region.GroupSubtotal != null && groupDetails != 0)
                        throw InvalidMarker(region, marker.Row, marker.Column);
                    footerCount++;
                    subtotalSeen = false;
                    pageDetails = 0;
                    finished = last;
                    awaitingHeader = !last && region.PageHeader != null;
                    break;
            }
            cursor = marker.EndRow + 1;
        }
        if (awaitingHeader)
            throw InvalidMarker(region, cursor + region.PageHeader.Marker.Row
                - region.PageHeader.SourceRange.First.Row, region.PageHeader.Marker.Column);
        if (!finished || cursor != endRowExclusive)
        {
            var footerStart = endRowExclusive - region.PageFooter.GapRows
                - ExcelEntityRowPlan.GetFooterHeight(region.PageFooter);
            var expected = footerStart >= cursor && (long)pageDetails + footerStart - cursor <= capacity
                ? footerStart : cursor + Math.Min(capacity - pageDetails,
                    Math.Max(0, endRowExclusive - cursor));
            throw InvalidMarker(region, expected + region.PageFooter.GapRows, region.Start.Column);
        }
    }

    /// <summary>
    /// 创建分页标记位置错误。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="row">错误所在的零基行号。</param>
    /// <param name="column">错误所在的零基列号。</param>
    /// <returns>结构化配置异常。</returns>
    private static BingOfficesConfigurationException InvalidMarker<TEntity>(
        ExcelEntityListRegion<TEntity> region, int row,
        int? column = null) where TEntity : class, new() =>
        new BingOfficesConfigurationException(
            region.PageFooter == null
                ? $"列表区域分组、分页小计或重复页眉标记位置无效: {region.SheetName}!{region.Start.Address}"
                : $"列表区域每页尾部 {region.PageFooter.MarkerText} 或关联标记位置无效: {region.SheetName}!{region.Start.Address}",
            BingOfficesStage.Plan, region.SheetName, row + 1,
            (column ?? region.PageHeader?.Marker.Column ?? region.Start.Column) + 1);
}
