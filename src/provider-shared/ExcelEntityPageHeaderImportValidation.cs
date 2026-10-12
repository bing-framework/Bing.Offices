using Bing.Offices.Exceptions;

namespace Bing.Offices.Entities;

/// <summary>校验实体导入时的重复页眉定位。</summary>
internal static class ExcelEntityPageHeaderImportValidation
{
    /// <summary>校验解析后的列表起点紧邻首页页眉。</summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="region">已解析命名起点的列表区域。</param>
    internal static void ValidateStart<TEntity>(ExcelEntityListRegion<TEntity> region)
        where TEntity : class, new()
    {
        if (region.PageHeader == null)
            return;
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        if (region.PageHeader.SourceRange.Last.Row == firstDataRow - 1)
            return;
        throw new BingOfficesConfigurationException(
            $"重复页眉源区域未紧邻明细: {region.SheetName}!{region.Start.Address}",
            BingOfficesStage.Plan, region.SheetName, region.Start.Row + 1,
            region.Start.Column + 1);
    }

    /// <summary>校验无表头明细与页眉标记列不重叠。</summary>
    internal static void ValidateMarkerColumn<TEntity>(ExcelEntityListRegion<TEntity> region,
        int columnCount) where TEntity : class, new()
    {
        if (region.PageHeader == null || region.IncludeHeader
            || region.PageHeader.Marker.Column < region.Start.Column
            || region.PageHeader.Marker.Column >= region.Start.Column + columnCount)
            return;
        throw new BingOfficesConfigurationException("无表头列表的重复页眉标记列与明细列重叠。",
            BingOfficesStage.Plan, region.SheetName, region.PageHeader.Marker.Row + 1,
            region.PageHeader.Marker.Column + 1);
    }

    /// <summary>校验复制页的列表表头与映射标题一致。</summary>
    internal static void ValidateCopiedHeaders<TEntity>(ExcelEntityListRegion<TEntity> region,
        IReadOnlyList<(int StartRow, int EndRow)> spans, int columnCount,
        Func<int, string> getTitle, Func<int, int, string> readCell,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (region.PageHeader == null || !region.IncludeHeader)
            return;
        var header = region.PageHeader;
        var offset = region.Start.Row - header.SourceRange.First.Row;
        var markerOffset = header.Marker.Row - header.SourceRange.First.Row;
        foreach (var span in spans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (span.EndRow - span.StartRow + 1 != header.RowCount
                || !string.Equals(readCell(span.StartRow + markerOffset, header.Marker.Column),
                    header.MarkerText, StringComparison.Ordinal))
                continue;
            for (var index = 0; index < columnCount; index++)
            {
                if ((index & 255) == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                var row = span.StartRow + offset;
                var column = region.Start.Column + index;
                if (string.Equals(readCell(row, column), getTitle(index), StringComparison.Ordinal))
                    continue;
                throw new BingOfficesConfigurationException("重复页眉的列表表头与映射标题不一致。",
                    BingOfficesStage.Plan, region.SheetName, row + 1, column + 1);
            }
        }
    }
}
