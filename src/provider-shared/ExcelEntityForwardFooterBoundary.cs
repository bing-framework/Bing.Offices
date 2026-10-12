using System;
using System.Collections.Generic;
using System.Threading;
using System.Linq;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Entities;

/// <summary>
/// 为实体 Forward 导入裁剪分页与尾部辅助行并校验其标记。
/// </summary>
internal static class ExcelEntityForwardFooterBoundary
{
    /// <summary>
    /// 仅输出实体列表的明细行，同时校验分页小计、每页尾部和最终尾部的位置。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <typeparam name="TRow">前向读取器的行类型。</typeparam>
    /// <param name="rows">按物理行顺序读取的工作表行。</param>
    /// <param name="region">实体列表区域。</param>
    /// <param name="getRowIndex">获取物理行号的委托。</param>
    /// <param name="getCell">获取指定物理列单元格的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>不包含分页小计、GapRows 和尾部内容的行序列。</returns>
    internal static IEnumerable<TRow> Filter<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));
        if (getRowIndex == null)
            throw new ArgumentNullException(nameof(getRowIndex));
        if (getCell == null)
            throw new ArgumentNullException(nameof(getCell));
        var finalFooterPreFiltered = region?.Footer != null
            && region.PageHeader != null && region.PageFooter != null;
        if (finalFooterPreFiltered)
        {
            rows = FilterFinalFooter(rows, region, getRowIndex, getCell,
                cancellationToken);
        }
        if (region?.PageHeader != null && region.PageFooter != null)
        {
            rows = FilterPageHeaderAndFooter(rows, region, getRowIndex, getCell,
                cancellationToken);
        }
        else if (region?.PageFooter != null)
        {
            rows = FilterPageFooter(rows, region, getRowIndex, getCell, cancellationToken);
        }
        else if (region?.PageHeader != null)
            rows = FilterPageHeader(rows, region, getRowIndex, getCell, cancellationToken);
        if (region?.PageFooter == null && region?.PageSubtotal != null)
            rows = FilterPageSubtotal(rows, region, getRowIndex, getCell, cancellationToken);
        if (region?.PageFooter == null && region?.GroupSubtotal != null)
            rows = FilterGroupSubtotal(rows, region, getRowIndex, getCell, cancellationToken);
        if (region?.Footer == null || finalFooterPreFiltered)
        {
            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return row;
            }
            yield break;
        }

        foreach (var row in FilterFinalFooter(rows, region, getRowIndex, getCell,
            cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>
    /// 裁剪最终尾部及其前置间隔，并校验尾部范围。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <typeparam name="TRow">前向读取器的行类型。</typeparam>
    /// <param name="rows">按物理行顺序读取的工作表行。</param>
    /// <param name="region">实体列表区域。</param>
    /// <param name="getRowIndex">获取物理行号的委托。</param>
    /// <param name="getCell">获取指定物理列单元格的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>不包含最终尾部内容的行序列。</returns>
    private static IEnumerable<TRow> FilterFinalFooter<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var footer = region.Footer;
        if (footer == null)
            throw new InvalidOperationException("最终尾部定义不能为空。");
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var minimumMarkerRow = firstDataRow + footer.GapRows;
        var maximumRow = region.End?.Row;
        var footerHeight = ExcelEntityRowPlan.GetFooterHeight(footer);
        int? markerRow = null;
        var pendingRows = new Queue<TRow>();

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (maximumRow.HasValue && rowIndex > maximumRow.Value)
            {
                if (!markerRow.HasValue)
                    ThrowMissingMarker(region);
                yield break;
            }

            if (rowIndex < firstDataRow)
            {
                yield return row;
                continue;
            }

            if (IsMarker(getCell(row, region.Start.Column), footer.MarkerText))
            {
                if (markerRow.HasValue)
                    ThrowDuplicateMarker(region, rowIndex);
                if (rowIndex < minimumMarkerRow)
                    ThrowEarlyMarker(region, rowIndex);
                if (maximumRow.HasValue
                    && (long)rowIndex + footerHeight - 1 > maximumRow.Value)
                    ThrowFooterOverflow(region, rowIndex, footerHeight);
                while (pendingRows.Count > 0)
                {
                    var pending = pendingRows.Dequeue();
                    if (getRowIndex(pending) < (long)rowIndex - footer.GapRows)
                        yield return pending;
                }
                markerRow = rowIndex;
                continue;
            }

            if (markerRow.HasValue)
                continue;
            pendingRows.Enqueue(row);
            var cutoff = (long)rowIndex - footer.GapRows;
            while (pendingRows.Count > 0 && getRowIndex(pendingRows.Peek()) < cutoff)
                yield return pendingRows.Dequeue();
        }

        var footerStart = markerRow.HasValue
            ? (long)markerRow.Value - footer.GapRows
            : long.MaxValue;
        while (pendingRows.Count > 0)
        {
            var pending = pendingRows.Dequeue();
            if (!markerRow.HasValue || getRowIndex(pending) < footerStart)
                yield return pending;
        }
        if (!markerRow.HasValue)
            ThrowMissingMarker(region);
    }

    /// <summary>
    /// 按页缓冲并裁剪每页尾部、分页小计和分组小计区域。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <typeparam name="TRow">前向读取器的行类型。</typeparam>
    /// <param name="rows">按物理行顺序读取的工作表行。</param>
    /// <param name="region">实体列表区域。</param>
    /// <param name="getRowIndex">获取物理行号的委托。</param>
    /// <param name="getCell">获取指定物理列单元格的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>不包含分页辅助行的列表行序列。</returns>
    private static IEnumerable<TRow> FilterPageFooter<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var pageFooter = region.PageFooter;
        var rowsPerPage = region.PageBreakRows;
        if (pageFooter == null || !rowsPerPage.HasValue || rowsPerPage.Value <= 0)
            throw new BingOfficesConfigurationException(
                $"列表区域每页尾部必须同时配置有效的 PageBreak: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var maximumRow = region.End?.Row;
        var pageRows = new List<TRow>();
        var pageCount = 0;
        var previousPageRequiresSubtotal = false;
        var previousPageHadSubtotal = false;
        long activePageFooterEnd = -1;
        long lastPageFooterEnd = -1;
        var sawPageFooter = false;
        var passThrough = false;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (maximumRow.HasValue && rowIndex > maximumRow.Value)
            {
                if (!passThrough)
                {
                    if (pageRows.Count > 0)
                        ThrowMissingPageFooter(region, rowIndex);
                    if (!sawPageFooter)
                        ThrowMissingPageFooter(region, rowIndex);
                    if (region.Footer != null)
                        ThrowMissingMarker(region);
                }
                yield break;
            }
            if (rowIndex < firstDataRow || passThrough)
            {
                yield return row;
                continue;
            }

            if (activePageFooterEnd >= 0)
            {
                if (rowIndex <= activePageFooterEnd)
                {
                    var footerCell = getCell(row, region.Start.Column);
                    if (IsMarker(footerCell, pageFooter.MarkerText))
                        ThrowDuplicatePageFooter(region, rowIndex);
                    if (region.Footer != null && IsMarker(footerCell, region.Footer.MarkerText)
                        || region.PageSubtotal != null
                        && IsMarker(footerCell, region.PageSubtotal.MarkerText)
                        || region.GroupSubtotal != null
                        && IsMarker(footerCell, region.GroupSubtotal.Footer.MarkerText))
                        ThrowPageFooterInvalid(region, rowIndex);
                    continue;
                }
                activePageFooterEnd = -1;
            }

            pageRows.Add(row);
            var isPageFooter = IsMarker(getCell(row, region.Start.Column),
                pageFooter.MarkerText);
            var isFinalFooter = region.Footer != null
                && IsMarker(getCell(row, region.Start.Column), region.Footer.MarkerText);
            if (isPageFooter && isFinalFooter)
                ThrowMarkerConflict(region, rowIndex);

            if (isFinalFooter)
            {
                var finalFooterStart = (long)rowIndex - region.Footer.GapRows;
                var hasCurrentPageRows = pageRows.Any(item =>
                    getRowIndex(item) < finalFooterStart);
                if (!sawPageFooter || hasCurrentPageRows
                    || region.PageSubtotal != null && previousPageHadSubtotal
                    || lastPageFooterEnd >= 0 && finalFooterStart <= lastPageFooterEnd)
                    ThrowPageFooterInvalid(region, rowIndex);
                foreach (var buffered in pageRows)
                    yield return buffered;
                pageRows.Clear();
                passThrough = true;
                continue;
            }

            if (!isPageFooter)
                continue;

            if (previousPageRequiresSubtotal)
                ThrowMissingPageSubtotal(region, rowIndex);
            var page = ReadPageFooterPage(pageRows, region, firstDataRow,
                rowsPerPage.Value, getRowIndex, getCell, cancellationToken,
                pageCount == 0);
            foreach (var detail in page.Details)
                yield return detail;
            pageRows.Clear();
            pageCount++;
            sawPageFooter = true;
            previousPageRequiresSubtotal = page.RequiresSubtotalIfContinued;
            previousPageHadSubtotal = page.HasSubtotal;
            activePageFooterEnd = (long)rowIndex + ExcelEntityRowPlan.GetFooterHeight(pageFooter) - 1;
            lastPageFooterEnd = activePageFooterEnd;
        }

        if (!passThrough && pageRows.Count > 0)
            ThrowMissingPageFooter(region, getRowIndex(pageRows[pageRows.Count - 1]));
        if (!sawPageFooter)
            ThrowMissingPageFooter(region, firstDataRow);
        if (region.Footer != null && !passThrough)
            ThrowMissingMarker(region);
    }

    /// <summary>
    /// 裁剪同时包含实际重复页眉和每页尾部的列表行。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <typeparam name="TRow">前向读取器的行类型。</typeparam>
    /// <param name="rows">按物理行顺序读取的工作表行。</param>
    /// <param name="region">实体列表区域。</param>
    /// <param name="getRowIndex">获取物理行号的委托。</param>
    /// <param name="getCell">获取指定物理列单元格的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>不包含页眉和每页尾部的明细行序列。</returns>
    private static IEnumerable<TRow> FilterPageHeaderAndFooter<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var header = region.PageHeader;
        var footer = region.PageFooter;
        if (header == null || footer == null)
            throw new InvalidOperationException("实际重复页眉和每页尾部必须同时存在。");
        var rowsPerPage = region.PageBreakRows;
        if (!rowsPerPage.HasValue || rowsPerPage.Value <= 0)
            throw new BingOfficesConfigurationException(
                $"列表区域实际重复页眉和每页尾部必须同时配置有效的 PageBreak: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (!region.End.HasValue)
            throw new BingOfficesConfigurationException(
                $"Forward 实际重复页眉和每页尾部必须声明结束地址: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);

        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var sourceFirstRow = header.SourceRange.First.Row;
        var sourceLastRow = header.SourceRange.Last.Row;
        var markerOffset = header.Marker.Row - sourceFirstRow;
        var listHeaderOffset = region.Start.Row - sourceFirstRow;
        var headerRowCount = header.RowCount;
        var maximumRow = region.End.Value.Row;
        var footerHeight = ExcelEntityRowPlan.GetFooterHeight(footer);
        var firstHeaderValues = new Dictionary<int, string>();
        var pageRows = new List<TRow>();
        var sourceHeaderComplete = false;
        var sourceMarkerSeen = false;
        var copiedHeaderStart = -1;
        var headerRowsRemaining = 0;
        var copiedMarkerSeen = false;
        var footerMarkerRow = -1;
        var footerRowsRemaining = 0;
        var pageIndex = 0;
        var sawFooter = false;
        var lastPageDetails = -1;
        var expectingHeader = false;
        var passThrough = false;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (rowIndex > maximumRow)
                break;

            if (passThrough)
            {
                yield return row;
                continue;
            }

            if (rowIndex < sourceFirstRow)
            {
                yield return row;
                continue;
            }

            if (!sourceHeaderComplete)
            {
                if (rowIndex > sourceLastRow)
                    ThrowPageHeaderMarker(region, sourceLastRow, header.Marker.Column);
                if (rowIndex == header.Marker.Row)
                {
                    sourceMarkerSeen = IsMarker(getCell(row, header.Marker.Column),
                        header.MarkerText);
                    if (!sourceMarkerSeen)
                        ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                }
                if (rowIndex == region.Start.Row && region.IncludeHeader)
                    CaptureHeaderValues(row, getCell, region.Start.Column,
                        region.End.Value.Column, firstHeaderValues);
                yield return row;
                if (rowIndex == sourceLastRow)
                {
                    sourceHeaderComplete = true;
                    if (!sourceMarkerSeen)
                        ThrowPageHeaderMarker(region, header.Marker.Row, header.Marker.Column);
                }
                continue;
            }

            if (footerRowsRemaining > 0)
            {
                var expected = footerMarkerRow + footerHeight - footerRowsRemaining;
                if (rowIndex != expected)
                    ThrowPageFooterInvalid(region, rowIndex);
                var footerCell = getCell(row, region.Start.Column);
                if (IsMarker(footerCell, footer.MarkerText))
                    ThrowDuplicatePageFooter(region, rowIndex);
                if (region.Footer != null && IsMarker(footerCell, region.Footer.MarkerText))
                    ThrowMarkerConflict(region, rowIndex);
                pageRows.Add(row);
                footerRowsRemaining--;
                if (footerRowsRemaining == 0)
                {
                    var page = ReadPageFooterPage(pageRows, region, firstDataRow,
                        rowsPerPage.Value, getRowIndex, getCell, cancellationToken,
                        pageIndex == 0);
                    foreach (var detail in page.Details)
                        yield return detail;
                    lastPageDetails = page.Details.Count;
                    pageRows.Clear();
                    pageIndex++;
                    sawFooter = true;
                    expectingHeader = true;
                }
                continue;
            }

            if (headerRowsRemaining > 0)
            {
                var expected = copiedHeaderStart + headerRowCount - headerRowsRemaining;
                if (rowIndex != expected)
                    ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                if (rowIndex == copiedHeaderStart + markerOffset)
                {
                    copiedMarkerSeen = IsMarker(getCell(row, header.Marker.Column),
                        header.MarkerText);
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                }
                if (region.IncludeHeader && rowIndex == copiedHeaderStart + listHeaderOffset)
                    ValidateCopiedHeader(row, getCell, firstHeaderValues, region, rowIndex);
                headerRowsRemaining--;
                if (headerRowsRemaining == 0)
                {
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, copiedHeaderStart + markerOffset,
                            header.Marker.Column);
                    expectingHeader = false;
                }
                continue;
            }

            if (expectingHeader)
            {
                if (region.Footer != null
                    && IsMarker(getCell(row, region.Start.Column), region.Footer.MarkerText))
                {
                    yield return row;
                    expectingHeader = false;
                    passThrough = true;
                    continue;
                }
                if (lastPageDetails < rowsPerPage.Value)
                    ThrowPageFooterInvalid(region, rowIndex);
                copiedHeaderStart = rowIndex;
                headerRowsRemaining = headerRowCount;
                copiedMarkerSeen = false;
                expectingHeader = false;
                // 继续使用本轮行处理复制页眉，避免丢失首行取消检查。
                var expected = copiedHeaderStart;
                if (rowIndex != expected)
                    ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                if (rowIndex == copiedHeaderStart + markerOffset)
                {
                    copiedMarkerSeen = IsMarker(getCell(row, header.Marker.Column),
                        header.MarkerText);
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                }
                if (region.IncludeHeader && rowIndex == copiedHeaderStart + listHeaderOffset)
                    ValidateCopiedHeader(row, getCell, firstHeaderValues, region, rowIndex);
                headerRowsRemaining--;
                if (headerRowsRemaining == 0)
                    expectingHeader = false;
                continue;
            }

            if (rowIndex < firstDataRow)
                continue;
            var markerCell = getCell(row, region.Start.Column);
            var isFinalFooter = region.Footer != null
                && IsMarker(markerCell, region.Footer.MarkerText);
            var isPageFooter = IsMarker(markerCell, footer.MarkerText);
            if (isPageFooter && isFinalFooter)
                ThrowMarkerConflict(region, rowIndex);
            if (isFinalFooter)
                ThrowPageFooterInvalid(region, rowIndex);
            if (isPageFooter)
            {
                footerMarkerRow = rowIndex;
                pageRows.Add(row);
                footerRowsRemaining = footerHeight - 1;
                if (footerRowsRemaining == 0)
                {
                    var page = ReadPageFooterPage(pageRows, region, firstDataRow,
                        rowsPerPage.Value, getRowIndex, getCell, cancellationToken,
                        pageIndex == 0);
                    foreach (var detail in page.Details)
                        yield return detail;
                    lastPageDetails = page.Details.Count;
                    pageRows.Clear();
                    pageIndex++;
                    sawFooter = true;
                    expectingHeader = true;
                }
                continue;
            }
            pageRows.Add(row);
        }

        if (footerRowsRemaining > 0)
            ThrowPageFooterInvalid(region, footerMarkerRow);
        if (headerRowsRemaining > 0)
            ThrowPageHeaderMarker(region, copiedHeaderStart + markerOffset,
                header.Marker.Column);
        if (pageRows.Count > 0)
            ThrowMissingPageFooter(region, getRowIndex(pageRows[pageRows.Count - 1]));
        if (!sawFooter)
            ThrowMissingPageFooter(region, firstDataRow);
    }

    /// <summary>
    /// 解析一个已包含每页尾部标记的物理页。
    /// </summary>
    private static PageFooterPageResult<TRow> ReadPageFooterPage<TEntity, TRow>(
        IReadOnlyList<TRow> rows, ExcelEntityListRegion<TEntity> region, int firstDataRow,
        int rowsPerPage, Func<TRow, int> getRowIndex,
        Func<TRow, int, ExcelCellValue> getCell, CancellationToken cancellationToken,
        bool firstPage) where TEntity : class, new()
    {
        var footer = region.PageFooter;
        var footerHeight = ExcelEntityRowPlan.GetFooterHeight(footer);
        var footerMarkers = rows.Where(row => IsMarker(getCell(row, region.Start.Column),
            footer.MarkerText)).ToArray();
        if (footerMarkers.Length == 0)
            ThrowMissingPageFooter(region, getRowIndex(rows[rows.Count - 1]));
        if (footerMarkers.Length > 1)
            ThrowDuplicatePageFooter(region, getRowIndex(footerMarkers[1]));

        var footerMarkerRow = getRowIndex(footerMarkers[0]);
        var footerStart = (long)footerMarkerRow - footer.GapRows;
        var footerEnd = (long)footerMarkerRow + footerHeight - 1;
        if (footerStart < firstDataRow || footerEnd > int.MaxValue
            || region.End.HasValue && footerEnd > region.End.Value.Row)
            ThrowPageFooterInvalid(region, footerMarkerRow);

        var spans = new List<(int Start, int End)>();
        var markers = new List<(int Row, bool IsGroup, bool IsPage)>();
        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (rowIndex >= footerStart)
            {
                var gapText = getCell(row, region.Start.Column)?.Text?.Trim();
                if (rowIndex < footerMarkerRow
                    && (region.PageSubtotal != null
                        && string.Equals(gapText, region.PageSubtotal.MarkerText,
                            StringComparison.Ordinal)
                        || region.GroupSubtotal != null
                        && string.Equals(gapText, region.GroupSubtotal.Footer.MarkerText,
                            StringComparison.Ordinal)))
                    ThrowPageFooterInvalid(region, rowIndex);
                continue;
            }

            var text = getCell(row, region.Start.Column)?.Text?.Trim();
            var isGroup = region.GroupSubtotal != null
                && string.Equals(text, region.GroupSubtotal.Footer.MarkerText,
                    StringComparison.Ordinal);
            var isPage = region.PageSubtotal != null
                && string.Equals(text, region.PageSubtotal.MarkerText,
                    StringComparison.Ordinal);
            var isFooter = string.Equals(text, footer.MarkerText, StringComparison.Ordinal);
            if ((isGroup ? 1 : 0) + (isPage ? 1 : 0) + (isFooter ? 1 : 0) > 1)
                ThrowMarkerConflict(region, rowIndex);
            if (isFooter)
                ThrowPageFooterInvalid(region, rowIndex);
            if (!isGroup && !isPage)
                continue;

            var listFooter = isGroup ? region.GroupSubtotal.Footer : region.PageSubtotal;
            var height = ExcelEntityRowPlan.GetFooterHeight(listFooter);
            var start = (long)rowIndex - listFooter.GapRows;
            var end = (long)rowIndex + height - 1;
            if (start < firstDataRow || end >= footerStart || end > int.MaxValue
                || spans.Any(span => start <= span.End))
                ThrowPageFooterInvalid(region, rowIndex);
            spans.Add(((int)start, (int)end));
            markers.Add((rowIndex, isGroup, isPage));
        }

        var details = new List<TRow>();
        var detailCount = 0;
        var groupDetails = 0;
        var pageSubtotalSeen = false;
        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (rowIndex >= footerStart)
                continue;
            var spanIndex = -1;
            for (var index = 0; index < spans.Count; index++)
            {
                if (rowIndex < spans[index].Start || rowIndex > spans[index].End)
                    continue;
                spanIndex = index;
                break;
            }
            if (spanIndex >= 0)
            {
                var markerIndex = -1;
                for (var index = 0; index < markers.Count; index++)
                {
                    if (markers[index].Row == rowIndex)
                    {
                        markerIndex = index;
                        break;
                    }
                }
                if (markerIndex >= 0)
                {
                    var marker = markers[markerIndex];
                    if (marker.IsGroup)
                    {
                        if (groupDetails == 0 || pageSubtotalSeen)
                            ThrowPageFooterInvalid(region, rowIndex);
                        groupDetails = 0;
                    }
                    else
                    {
                        if (pageSubtotalSeen || detailCount != rowsPerPage)
                            ThrowPageFooterInvalid(region, rowIndex);
                        pageSubtotalSeen = true;
                    }
                }
                continue;
            }

            details.Add(row);
            detailCount++;
            groupDetails++;
            if (detailCount > rowsPerPage)
                ThrowPageFooterInvalid(region, rowIndex);
        }

        if (detailCount == 0 && !firstPage || pageSubtotalSeen && detailCount != rowsPerPage)
            ThrowPageFooterInvalid(region, footerMarkerRow);
        return new PageFooterPageResult<TRow>(details,
            requiresSubtotalIfContinued: region.PageSubtotal != null && !pageSubtotalSeen,
            hasSubtotal: pageSubtotalSeen);
    }

    /// <summary>
    /// 裁剪 Forward 导入中的实际重复页眉。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <typeparam name="TRow">前向读取器的行类型。</typeparam>
    /// <param name="rows">按物理行顺序读取的工作表行。</param>
    /// <param name="region">实体列表区域。</param>
    /// <param name="getRowIndex">获取物理行号的委托。</param>
    /// <param name="getCell">获取指定物理列单元格的委托。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>不包含后续逻辑页页眉的行序列。</returns>
    /// <remarks>
    /// Forward 当前只接受独立的分页页眉。页眉与小计、尾部或逐页尾部的组合由 Provider
    /// 预检拒绝；这样可以在不缓存整个工作表的情况下保持页眉与明细的物理边界。
    /// </remarks>
    private static IEnumerable<TRow> FilterPageHeader<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var header = region.PageHeader;
        if (header == null)
        {
            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return row;
            }
            yield break;
        }

        var rowsPerPage = region.PageBreakRows;
        if (!rowsPerPage.HasValue || rowsPerPage.Value <= 0)
            throw new BingOfficesConfigurationException(
                $"列表区域重复页眉必须同时配置有效的 PageBreak: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (!region.End.HasValue)
            throw new BingOfficesConfigurationException(
                $"Forward 重复页眉必须声明结束地址: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);

        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var sourceFirstRow = header.SourceRange.First.Row;
        var markerOffset = header.Marker.Row - sourceFirstRow;
        var listHeaderOffset = region.Start.Row - sourceFirstRow;
        var headerRowCount = header.RowCount;
        var maximumRow = region.End.Value.Row;
        var detailsInPage = 0;
        var previousDataRow = -1;
        var headerRowsRemaining = 0;
        var copiedHeaderStart = -1;
        var copiedMarkerSeen = false;
        var firstHeaderValues = new Dictionary<int, string>();
        var firstMarkerSeen = false;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (rowIndex > maximumRow)
                break;

            if (rowIndex < sourceFirstRow)
            {
                yield return row;
                continue;
            }

            if (rowIndex <= header.SourceRange.Last.Row)
            {
                if (rowIndex == header.Marker.Row)
                {
                    firstMarkerSeen = IsMarker(getCell(row, header.Marker.Column),
                        header.MarkerText);
                    if (!firstMarkerSeen)
                        ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                }
                if (rowIndex == region.Start.Row && region.IncludeHeader)
                    CaptureHeaderValues(row, getCell, region.Start.Column,
                        region.End.Value.Column, firstHeaderValues);
                yield return row;
                continue;
            }

            if (headerRowsRemaining > 0)
            {
                if (rowIndex != copiedHeaderStart + header.RowCount - headerRowsRemaining)
                    ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                if (rowIndex == copiedHeaderStart + markerOffset)
                {
                    copiedMarkerSeen = IsMarker(getCell(row, header.Marker.Column),
                        header.MarkerText);
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                }
                if (region.IncludeHeader && rowIndex == copiedHeaderStart + listHeaderOffset)
                    ValidateCopiedHeader(row, getCell, firstHeaderValues,
                        region, rowIndex);
                headerRowsRemaining--;
                if (headerRowsRemaining == 0)
                {
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, copiedHeaderStart + markerOffset,
                            header.Marker.Column);
                    detailsInPage = 0;
                    previousDataRow = -1;
                }
                continue;
            }

            if (detailsInPage >= rowsPerPage.Value)
            {
                copiedHeaderStart = rowIndex;
                if (previousDataRow >= 0 && rowIndex != previousDataRow + 1)
                    ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                headerRowsRemaining = headerRowCount;
                copiedMarkerSeen = false;
                if (rowIndex != copiedHeaderStart)
                    ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                if (rowIndex == copiedHeaderStart + markerOffset)
                {
                    copiedMarkerSeen = IsMarker(getCell(row, header.Marker.Column),
                        header.MarkerText);
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, rowIndex, header.Marker.Column);
                }
                if (region.IncludeHeader && rowIndex == copiedHeaderStart + listHeaderOffset)
                    ValidateCopiedHeader(row, getCell, firstHeaderValues,
                        region, rowIndex);
                headerRowsRemaining--;
                if (headerRowsRemaining == 0)
                {
                    if (!copiedMarkerSeen)
                        ThrowPageHeaderMarker(region, copiedHeaderStart + markerOffset,
                            header.Marker.Column);
                    detailsInPage = 0;
                    previousDataRow = -1;
                }
                continue;
            }

            if (rowIndex < firstDataRow)
            {
                yield return row;
                continue;
            }

            detailsInPage++;
            previousDataRow = rowIndex;
            yield return row;
        }
        if (headerRowsRemaining > 0)
            ThrowPageHeaderMarker(region, copiedHeaderStart + markerOffset,
                header.Marker.Column);
        if (!firstMarkerSeen)
            ThrowPageHeaderMarker(region, header.Marker.Row, header.Marker.Column);
    }

    /// <summary>
    /// 保存首页列表表头的文本快照。
    /// </summary>
    private static void CaptureHeaderValues<TRow>(TRow row,
        Func<TRow, int, ExcelCellValue> getCell, int firstColumn, int lastColumn,
        IDictionary<int, string> values)
    {
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            var cell = getCell(row, column);
            if (cell == null)
                continue;
            var text = cell.Text?.Trim() ?? string.Empty;
            if (text.Length > 0)
                values[column] = text;
        }
    }

    /// <summary>
    /// 校验复制页的列表表头标题。
    /// </summary>
    private static void ValidateCopiedHeader<TEntity, TRow>(TRow row,
        Func<TRow, int, ExcelCellValue> getCell, IReadOnlyDictionary<int, string> expected,
        ExcelEntityListRegion<TEntity> region,
        int rowIndex) where TEntity : class, new()
    {
        foreach (var pair in expected)
        {
            var actual = getCell(row, pair.Key)?.Text?.Trim() ?? string.Empty;
            if (string.Equals(actual, pair.Value, StringComparison.Ordinal))
                continue;
            throw new BingOfficesConfigurationException("重复页眉的列表表头与映射标题不一致。",
                BingOfficesStage.Plan, region.SheetName, rowIndex + 1, pair.Key + 1);
        }
    }

    private static void ThrowPageHeaderMarker<TEntity>(ExcelEntityListRegion<TEntity> region,
        int row, int column) where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域重复页眉标记位置无效: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan, sheetName: region.SheetName, rowIndex: row + 1,
            columnIndex: column + 1);

    /// <summary>
    /// 裁剪连续分组小计区域，并校验每个小计前存在明细。
    /// </summary>
    private static IEnumerable<TRow> FilterGroupSubtotal<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var subtotal = region.GroupSubtotal?.Footer;
        if (subtotal == null)
            yield break;

        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var maximumRow = region.End?.Row;
        var subtotalHeight = ExcelEntityRowPlan.GetFooterHeight(subtotal);
        var cursor = firstDataRow;
        var pending = new Queue<TRow>();
        var finalFooterSeen = false;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (maximumRow.HasValue && rowIndex > maximumRow.Value)
                break;
            if (rowIndex < firstDataRow)
            {
                yield return row;
                continue;
            }

            var isSubtotal = IsMarker(getCell(row, region.Start.Column), subtotal.MarkerText);
            var isFinalFooter = region.Footer != null
                && IsMarker(getCell(row, region.Start.Column), region.Footer.MarkerText);
            if (isSubtotal && isFinalFooter)
                ThrowMarkerConflict(region, rowIndex);
            if (rowIndex < cursor)
            {
                if (isSubtotal || isFinalFooter)
                    ThrowGroupSubtotalMisplaced(region, rowIndex);
                continue;
            }
            if (isFinalFooter)
            {
                while (pending.Count > 0)
                    yield return pending.Dequeue();
                yield return row;
                finalFooterSeen = true;
                continue;
            }
            if (isSubtotal)
            {
                if ((long)rowIndex - subtotal.GapRows <= cursor)
                    ThrowGroupSubtotalMisplaced(region, rowIndex);
                if (maximumRow.HasValue
                    && (long)rowIndex + subtotalHeight - 1 > maximumRow.Value)
                    ThrowGroupSubtotalOverflow(region, rowIndex, subtotalHeight);
                while (pending.Count > 0
                    && getRowIndex(pending.Peek()) < (long)rowIndex - subtotal.GapRows)
                    yield return pending.Dequeue();
                pending.Clear();
                cursor = checked(rowIndex + subtotalHeight);
                continue;
            }
            if (finalFooterSeen)
            {
                yield return row;
                continue;
            }
            pending.Enqueue(row);
            var cutoff = (long)rowIndex - subtotal.GapRows;
            while (pending.Count > 0 && getRowIndex(pending.Peek()) < cutoff)
                yield return pending.Dequeue();
        }

        while (pending.Count > 0)
            yield return pending.Dequeue();
    }

    /// <summary>
    /// 裁剪完整分页后的分页小计区域，并校验每页明细容量。
    /// </summary>
    private static IEnumerable<TRow> FilterPageSubtotal<TEntity, TRow>(
        IEnumerable<TRow> rows, ExcelEntityListRegion<TEntity> region,
        Func<TRow, int> getRowIndex, Func<TRow, int, ExcelCellValue> getCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var subtotal = region.PageSubtotal;
        var rowsPerPage = region.PageBreakRows;
        if (subtotal == null || !rowsPerPage.HasValue || rowsPerPage.Value <= 0)
            throw new BingOfficesConfigurationException(
                $"列表区域分页小计必须同时配置有效的 PageBreak: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);

        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var maximumRow = region.End?.Row;
        var subtotalHeight = ExcelEntityRowPlan.GetFooterHeight(subtotal);
        var cursor = firstDataRow;
        var expectedMarker = cursor + rowsPerPage.Value + subtotal.GapRows;
        var pending = new Queue<TRow>();
        var sawSubtotal = false;
        var detailAfterSubtotal = false;
        var finalFooterSeen = false;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rowIndex = getRowIndex(row);
            if (maximumRow.HasValue && rowIndex > maximumRow.Value)
                break;
            if (rowIndex < firstDataRow)
            {
                yield return row;
                continue;
            }

            var isSubtotal = IsMarker(getCell(row, region.Start.Column), subtotal.MarkerText);
            var isFinalFooter = region.Footer != null
                && IsMarker(getCell(row, region.Start.Column), region.Footer.MarkerText);
            if (isSubtotal && isFinalFooter)
                ThrowMarkerConflict(region, rowIndex);
            if (rowIndex < cursor)
            {
                if (isSubtotal || isFinalFooter)
                    ThrowMarkerConflict(region, rowIndex);
                continue;
            }
            if (isSubtotal || isFinalFooter)
            {
                if (rowIndex < cursor)
                    ThrowSubtotalEarly(region, rowIndex);
                if (isSubtotal && sawSubtotal && !detailAfterSubtotal)
                    ThrowSubtotalAfterLastDetail(region, rowIndex);
                if (isSubtotal && rowIndex < cursor + rowsPerPage.Value + subtotal.GapRows)
                    ThrowSubtotalEarly(region, rowIndex);
                if (rowIndex < cursor + rowsPerPage.Value + subtotal.GapRows
                    && !isFinalFooter)
                    ThrowSubtotalEarly(region, rowIndex);

                if (isFinalFooter)
                {
                    var detailEnd = (long)rowIndex - region.Footer.GapRows;
                    var details = detailEnd - cursor;
                    if (details < 0)
                        ThrowSubtotalEarly(region, rowIndex);
                    if (details > rowsPerPage.Value)
                        ThrowMissingSubtotal(region, rowIndex);
                    if (sawSubtotal && !detailAfterSubtotal)
                        ThrowSubtotalAfterLastDetail(region, rowIndex);
                    while (pending.Count > 0)
                        yield return pending.Dequeue();
                    yield return row;
                    finalFooterSeen = true;
                    continue;
                }

                if (rowIndex != expectedMarker)
                    ThrowSubtotalMisplaced(region, rowIndex);
                if (maximumRow.HasValue
                    && (long)rowIndex + subtotalHeight - 1 > maximumRow.Value)
                    ThrowSubtotalOverflow(region, rowIndex, subtotalHeight);

                var detailEndRow = cursor + rowsPerPage.Value;
                while (pending.Count > 0
                    && getRowIndex(pending.Peek()) < detailEndRow)
                    yield return pending.Dequeue();
                pending.Clear();
                cursor = checked(rowIndex + subtotalHeight);
                expectedMarker = checked(cursor + rowsPerPage.Value + subtotal.GapRows);
                sawSubtotal = true;
                detailAfterSubtotal = false;
                continue;
            }

            if (finalFooterSeen)
            {
                yield return row;
                continue;
            }
            if (rowIndex >= expectedMarker)
                ThrowMissingSubtotal(region, rowIndex);
            if (sawSubtotal)
                detailAfterSubtotal = true;
            pending.Enqueue(row);
            var dataEnd = cursor + rowsPerPage.Value;
            while (pending.Count > 0 && getRowIndex(pending.Peek()) < dataEnd)
                yield return pending.Dequeue();
        }

        var lastPageDataEnd = cursor + rowsPerPage.Value;
        if (!finalFooterSeen && pending.Any(item => getRowIndex(item) >= lastPageDataEnd))
            ThrowMissingSubtotal(region, getRowIndex(pending.Last()));
        while (pending.Count > 0)
            yield return pending.Dequeue();
        if (sawSubtotal && !detailAfterSubtotal && !finalFooterSeen)
            ThrowSubtotalAfterLastDetail(region, cursor);
    }

    /// <summary>
    /// 判断行首单元格是否包含尾部结束标记。
    /// </summary>
    private static bool IsMarker(ExcelCellValue cell, string markerText) =>
        string.Equals(cell?.Text?.Trim(), markerText, StringComparison.Ordinal);

    private static void ThrowMarkerConflict<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分页小计与最终尾部标记冲突: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowMissingMarker<TEntity>(ExcelEntityListRegion<TEntity> region)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域缺少尾部标记: {region.SheetName}!{region.Start.Address}",
            stage: BingOfficesStage.Plan);

    private static void ThrowDuplicateMarker<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域尾部标记重复: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowEarlyMarker<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域尾部标记位于明细起始位置之前: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowFooterOverflow<TEntity>(ExcelEntityListRegion<TEntity> region,
        int row, int height) where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域尾部超出结束地址: {region.SheetName}!{region.Start.Address}，行 {row + height}",
            stage: BingOfficesStage.Plan);

    private static void ThrowSubtotalEarly<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分页小计标记位置无效（位于明细起始位置之前或当前页容量未满）: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowSubtotalMisplaced<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分页小计标记位置无效: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowMissingSubtotal<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域缺少分页小计标记: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    /// <summary>
    /// 抛出缺少每页尾部标记的配置异常。
    /// </summary>
    private static void ThrowMissingPageFooter<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域缺少每页尾部标记: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan, sheetName: region.SheetName,
            rowIndex: row + 1, columnIndex: region.Start.Column + 1);

    /// <summary>
    /// 抛出缺少每页分页小计标记的配置异常。
    /// </summary>
    private static void ThrowMissingPageSubtotal<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域缺少分页小计标记: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan, sheetName: region.SheetName,
            rowIndex: row + 1, columnIndex: region.Start.Column + 1);

    /// <summary>
    /// 抛出重复每页尾部标记的配置异常。
    /// </summary>
    private static void ThrowDuplicatePageFooter<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域每页尾部标记重复: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan, sheetName: region.SheetName,
            rowIndex: row + 1, columnIndex: region.Start.Column + 1);

    /// <summary>
    /// 抛出每页尾部或关联标记位置错误的配置异常。
    /// </summary>
    private static void ThrowPageFooterInvalid<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域每页尾部或关联标记位置无效: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan, sheetName: region.SheetName,
            rowIndex: row + 1, columnIndex: region.Start.Column + 1);

    private static void ThrowSubtotalAfterLastDetail<TEntity>(ExcelEntityListRegion<TEntity> region, int row)
        where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分页小计之后缺少明细: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowSubtotalOverflow<TEntity>(ExcelEntityListRegion<TEntity> region,
        int row, int height) where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分页小计超出结束地址: {region.SheetName}!{region.Start.Address}，行 {row + height}",
            stage: BingOfficesStage.Plan);

    private static void ThrowGroupSubtotalMisplaced<TEntity>(
        ExcelEntityListRegion<TEntity> region, int row) where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分组小计标记位置无效: {region.SheetName}!{region.Start.Address}，行 {row + 1}",
            stage: BingOfficesStage.Plan);

    private static void ThrowGroupSubtotalOverflow<TEntity>(
        ExcelEntityListRegion<TEntity> region, int row, int height) where TEntity : class, new() =>
        throw new BingOfficesConfigurationException(
            $"列表区域分组小计超出结束地址: {region.SheetName}!{region.Start.Address}，行 {row + height}",
            stage: BingOfficesStage.Plan);

    /// <summary>
    /// 保存一页明细及其后续分页小计要求。
    /// </summary>
    /// <typeparam name="TRow">前向读取器的行类型。</typeparam>
    private sealed class PageFooterPageResult<TRow>
    {
        /// <summary>
        /// 初始化每页尾部解析结果。
        /// </summary>
        /// <param name="details">当前页明细行。</param>
        /// <param name="requiresSubtotalIfContinued">后续仍有页面时是否必须出现分页小计。</param>
        /// <param name="hasSubtotal">当前页是否包含分页小计。</param>
        internal PageFooterPageResult(IReadOnlyList<TRow> details,
            bool requiresSubtotalIfContinued, bool hasSubtotal)
        {
            Details = details ?? throw new ArgumentNullException(nameof(details));
            RequiresSubtotalIfContinued = requiresSubtotalIfContinued;
            HasSubtotal = hasSubtotal;
        }

        /// <summary>
        /// 获取当前页明细行。
        /// </summary>
        internal IReadOnlyList<TRow> Details { get; }

        /// <summary>
        /// 获取后续仍有页面时是否必须出现分页小计。
        /// </summary>
        internal bool RequiresSubtotalIfContinued { get; }

        /// <summary>
        /// 获取当前页是否包含分页小计。
        /// </summary>
        internal bool HasSubtotal { get; }
    }
}
