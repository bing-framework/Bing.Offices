using System.Collections;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 将完整实体布局导入结果投影为列表区域批次。
/// </summary>
internal static class ExcelEntityBatchProjection
{
    /// <summary>
    /// 同步交付实体布局批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="result">完整实体导入结果。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="batchSize">每个列表区域的目标批次大小。</param>
    /// <param name="onBatch">同步批次回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>实体布局批次摘要。</returns>
    internal static ExcelEntityBatchImportSummary<TEntity> Deliver<TEntity>(
        ExcelEntityImportResult<TEntity> result, ExcelEntityLayout<TEntity> layout, int batchSize,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        ValidateInputs(result, layout, onBatch, batchSize);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateRegions(result, layout, cancellationToken);
        var globalErrors = CreateGlobalErrors(result);
        var summaries = new List<ExcelEntityBatchRegionSummary>(layout.ListRegions.Count);
        var batchNumber = 1;
        if (globalErrors.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            onBatch(CreateDiagnosticBatch(result.Entity, globalErrors, batchNumber++));
        }

        for (var index = 0; index < layout.ListRegions.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // 每次只保留当前列表区域的投影，避免多区域结果同时复制到内存。
            var projection = CreateProjection(result, layout, index, cancellationToken);
            var summary = DeliverRegion(projection, result.Entity, batchSize, ref batchNumber,
                onBatch, cancellationToken);
            summaries.Add(new ExcelEntityBatchRegionSummary(projection.RegionIndex,
                projection.SheetName, projection.PropertyName, projection.ItemType, summary));
        }

        return new ExcelEntityBatchImportSummary<TEntity>(result, summaries, globalErrors);
    }

    /// <summary>
    /// 异步交付实体布局批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="result">完整实体导入结果。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="batchSize">每个列表区域的目标批次大小。</param>
    /// <param name="onBatch">异步批次回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>实体布局批次摘要任务。</returns>
    internal static async Task<ExcelEntityBatchImportSummary<TEntity>> DeliverAsync<TEntity>(
        ExcelEntityImportResult<TEntity> result, ExcelEntityLayout<TEntity> layout, int batchSize,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (result == null)
            throw new ArgumentNullException(nameof(result));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        if (batchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(batchSize));

        cancellationToken.ThrowIfCancellationRequested();
        ValidateRegions(result, layout, cancellationToken);
        var globalErrors = CreateGlobalErrors(result);
        var summaries = new List<ExcelEntityBatchRegionSummary>(layout.ListRegions.Count);
        var batchNumber = 1;
        if (globalErrors.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await onBatch(CreateDiagnosticBatch(result.Entity, globalErrors, batchNumber++),
                cancellationToken).ConfigureAwait(false);
        }

        for (var index = 0; index < layout.ListRegions.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // 每次只保留当前列表区域的投影，避免多区域结果同时复制到内存。
            var projection = CreateProjection(result, layout, index, cancellationToken);
            var state = new BatchNumberState(batchNumber);
            var summary = await DeliverRegionAsync(projection, result.Entity, batchSize, state,
                onBatch, cancellationToken).ConfigureAwait(false);
            batchNumber = state.Value;
            summaries.Add(new ExcelEntityBatchRegionSummary(projection.RegionIndex,
                projection.SheetName, projection.PropertyName, projection.ItemType, summary));
        }

        return new ExcelEntityBatchImportSummary<TEntity>(result, summaries, globalErrors);
    }

    /// <summary>
    /// 校验投影入口参数。
    /// </summary>
    private static void ValidateInputs<TEntity>(ExcelEntityImportResult<TEntity> result,
        ExcelEntityLayout<TEntity> layout, Action<ExcelEntityImportBatch<TEntity>> onBatch,
        int batchSize) where TEntity : class, new()
    {
        if (result == null)
            throw new ArgumentNullException(nameof(result));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        if (batchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(batchSize));
    }

    /// <summary>
    /// 预检实体布局区域结果，确保后续惰性投影不会在已交付批次后才发现结构错误。
    /// </summary>
    private static void ValidateRegions<TEntity>(ExcelEntityImportResult<TEntity> result,
        ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (result.Sheets.Count != layout.ListRegions.Count)
            throw new BingOfficesConfigurationException("实体布局列表区域结果数量不一致。",
                stage: BingOfficesStage.Plan);

        for (var index = 0; index < layout.ListRegions.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var region = layout.ListRegions[index];
            var sheet = result.Sheets[index];
            if (sheet == null || !string.Equals(sheet.Name, region.SheetName,
                    StringComparison.OrdinalIgnoreCase) || sheet.ItemType != region.ItemType)
                throw new BingOfficesConfigurationException(
                    $"实体布局列表区域结果无法匹配: {region.SheetName}.{region.PropertyName}",
                    stage: BingOfficesStage.Plan);

            var itemCount = CountItems(region.Getter(result.Entity), cancellationToken);
            var sourceRowCount = sheet.SourceRows?.Count ?? 0;
            if (itemCount != sourceRowCount)
                throw new BingOfficesConfigurationException(
                    $"实体布局列表区域实体数量与源行数量不一致: {region.PropertyName}",
                    stage: BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 分离无法归属列表区域的实体级错误。
    /// </summary>
    private static IReadOnlyList<ExcelImportError> CreateGlobalErrors<TEntity>(
        ExcelEntityImportResult<TEntity> result) where TEntity : class, new()
    {
        var sheets = result.Sheets ?? Array.Empty<ExcelSheetImportResult>();
        var hasRegionalErrors = false;
        foreach (var sheet in sheets)
        {
            if (sheet?.Errors != null && sheet.Errors.Count > 0)
            {
                hasRegionalErrors = true;
                break;
            }
        }

        var allErrors = result.Errors ?? Array.Empty<ExcelImportError>();
        if (!hasRegionalErrors)
            return allErrors;

        var regionalErrors = new HashSet<ExcelImportError>();
        foreach (var sheet in sheets)
        {
            foreach (var error in sheet?.Errors ?? Array.Empty<ExcelImportError>())
                regionalErrors.Add(error);
        }

        return allErrors.Where(error => !regionalErrors.Contains(error)).ToArray();
    }

    /// <summary>
    /// 创建单个列表区域的临时投影。
    /// </summary>
    private static RegionProjection CreateProjection<TEntity>(ExcelEntityImportResult<TEntity> result,
        ExcelEntityLayout<TEntity> layout, int index, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var region = layout.ListRegions[index];
        var sheet = result.Sheets[index];
        var items = CopyItems(region.Getter(result.Entity), cancellationToken);
        // 源行和区域错误在本次内部投影期间只读，直接引用结果列表，避免重复数组化。
        var sourceRows = sheet.SourceRows ?? Array.Empty<int>();
        var errors = sheet.Errors ?? Array.Empty<ExcelImportError>();
        return new RegionProjection(index, region.SheetName, region.PropertyName, region.ItemType,
            region.Start.Row, region.IncludeHeader, sourceRows, items, errors, result.ErrorsTruncated);
    }

    /// <summary>
    /// 获取实体列表的元素数量。
    /// </summary>
    private static int CountItems(object value, CancellationToken cancellationToken)
    {
        if (value is ICollection collection)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return collection.Count;
        }

        if (value is not IEnumerable enumerable)
            return 0;

        var count = 0;
        foreach (var _ in enumerable)
        {
            cancellationToken.ThrowIfCancellationRequested();
            count = checked(count + 1);
        }

        return count;
    }

    /// <summary>
    /// 按项复制列表区域实体并响应取消请求。
    /// </summary>
    private static IReadOnlyList<object> CopyItems(object value, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (value is not IEnumerable enumerable)
            return Array.Empty<object>();

        var capacity = value is ICollection collection ? collection.Count : 0;
        var items = new List<object>(capacity);
        foreach (var item in enumerable)
        {
            cancellationToken.ThrowIfCancellationRequested();
            items.Add(item);
        }

        return items.Count == 0 ? Array.Empty<object>() : items;
    }

    /// <summary>
    /// 同步交付单个列表区域的批次。
    /// </summary>
    private static ExcelBatchImportSummary DeliverRegion<TEntity>(RegionProjection projection,
        TEntity entity, int batchSize, ref int batchNumber,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var rowsRead = 0;
        var failedRows = 0;
        var items = new List<object>();
        var errors = new List<ExcelImportError>();
        var firstRow = 0;
        foreach (var rowEvent in EnumerateEvents(projection, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (rowEvent.Row >= projection.DataRow)
            {
                rowsRead++;
                if (rowEvent.Errors.Count > 0)
                    failedRows++;
            }
            firstRow = firstRow == 0 ? rowEvent.Row : firstRow;
            items.AddRange(rowEvent.Items);
            errors.AddRange(rowEvent.Errors);
            if (items.Count >= batchSize || errors.Count >= batchSize)
            {
                // 将已满缓冲交给批次对象，回调期间改用新缓冲，避免清空导致批次内容被复用。
                var batchItems = items;
                var batchErrors = errors;
                items = new List<object>();
                errors = new List<ExcelImportError>();
                onBatch(CreateRegionBatch(entity, projection, batchNumber++, firstRow,
                    batchItems, batchErrors));
                firstRow = 0;
            }
        }

        if (items.Count > 0 || errors.Count > 0)
            onBatch(CreateRegionBatch(entity, projection, batchNumber++, firstRow, items, errors));
        return projection.CreateSummary(rowsRead, failedRows);
    }

    /// <summary>
    /// 异步交付单个列表区域的批次。
    /// </summary>
    private static async Task<ExcelBatchImportSummary> DeliverRegionAsync<TEntity>(
        RegionProjection projection, TEntity entity, int batchSize, BatchNumberState batchState,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var rowsRead = 0;
        var failedRows = 0;
        var items = new List<object>();
        var errors = new List<ExcelImportError>();
        var firstRow = 0;
        foreach (var rowEvent in EnumerateEvents(projection, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (rowEvent.Row >= projection.DataRow)
            {
                rowsRead++;
                if (rowEvent.Errors.Count > 0)
                    failedRows++;
            }
            firstRow = firstRow == 0 ? rowEvent.Row : firstRow;
            items.AddRange(rowEvent.Items);
            errors.AddRange(rowEvent.Errors);
            if (items.Count >= batchSize || errors.Count >= batchSize)
            {
                // 将已满缓冲交给批次对象，回调期间改用新缓冲，避免清空导致批次内容被复用。
                var batchItems = items;
                var batchErrors = errors;
                items = new List<object>();
                errors = new List<ExcelImportError>();
                await onBatch(CreateRegionBatch(entity, projection, batchState.Value++, firstRow,
                    batchItems, batchErrors), cancellationToken).ConfigureAwait(false);
                firstRow = 0;
            }
        }

        if (items.Count > 0 || errors.Count > 0)
            await onBatch(CreateRegionBatch(entity, projection, batchState.Value++, firstRow,
                items, errors), cancellationToken).ConfigureAwait(false);
        return projection.CreateSummary(rowsRead, failedRows);
    }

    /// <summary>
    /// 创建实体列表区域批次。
    /// </summary>
    private static ExcelEntityImportBatch<TEntity> CreateRegionBatch<TEntity>(TEntity entity,
        RegionProjection projection, int batchNumber, int firstRow, IReadOnlyList<object> items,
        IReadOnlyList<ExcelImportError> errors) where TEntity : class, new() =>
        new ExcelEntityImportBatch<TEntity>(entity, projection.SheetName, projection.PropertyName,
            projection.ItemType, projection.RegionIndex, batchNumber, firstRow,
            items, errors, false);

    /// <summary>
    /// 创建实体级诊断批次。
    /// </summary>
    private static ExcelEntityImportBatch<TEntity> CreateDiagnosticBatch<TEntity>(TEntity entity,
        IReadOnlyList<ExcelImportError> errors, int batchNumber) where TEntity : class, new()
    {
        var first = errors.FirstOrDefault();
        return new ExcelEntityImportBatch<TEntity>(entity, first?.SheetName, null, typeof(TEntity), -1,
            batchNumber, first?.RowIndex > 0 ? first.RowIndex : 0, Array.Empty<object>(), errors, true);
    }

    /// <summary>
    /// 将成功项和区域错误按物理行合并为批次事件。
    /// </summary>
    private static IEnumerable<RowEvent> EnumerateEvents(RegionProjection projection,
        CancellationToken cancellationToken)
    {
        if (HasOrderedRows(projection))
        {
            foreach (var rowEvent in EnumerateOrderedEvents(projection, cancellationToken))
                yield return rowEvent;
            yield break;
        }

        foreach (var rowEvent in EnumerateIndexedEvents(projection, cancellationToken))
            yield return rowEvent;
    }

    /// <summary>
    /// 判断成功行和错误是否按源行顺序排列。
    /// </summary>
    private static bool HasOrderedRows(RegionProjection projection)
    {
        var previous = int.MinValue;
        for (var index = 0; index < projection.SourceRows.Count; index++)
        {
            var row = projection.SourceRows[index] + 1;
            if (row < previous)
                return false;
            previous = row;
        }

        previous = int.MinValue;
        for (var index = 0; index < projection.Errors.Count; index++)
        {
            var row = NormalizeErrorRow(projection.Errors[index], projection.DataRow);
            if (row < previous)
                return false;
            previous = row;
        }

        return true;
    }

    /// <summary>
    /// 按源行顺序合并成功项和错误，避免为整个区域创建索引字典。
    /// </summary>
    private static IEnumerable<RowEvent> EnumerateOrderedEvents(RegionProjection projection,
        CancellationToken cancellationToken)
    {
        var sourceIndex = 0;
        var errorIndex = 0;
        while (sourceIndex < projection.SourceRows.Count
            || errorIndex < projection.Errors.Count)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceRow = sourceIndex < projection.SourceRows.Count
                ? projection.SourceRows[sourceIndex] + 1 : int.MaxValue;
            var errorRow = errorIndex < projection.Errors.Count
                ? NormalizeErrorRow(projection.Errors[errorIndex], projection.DataRow)
                : int.MaxValue;
            var row = Math.Min(sourceRow, errorRow);
            var items = new List<object>();
            while (sourceIndex < projection.SourceRows.Count
                && projection.SourceRows[sourceIndex] + 1 == row)
            {
                cancellationToken.ThrowIfCancellationRequested();
                items.Add(projection.Items[sourceIndex++]);
            }

            var errors = new List<ExcelImportError>();
            while (errorIndex < projection.Errors.Count
                && NormalizeErrorRow(projection.Errors[errorIndex], projection.DataRow) == row)
            {
                cancellationToken.ThrowIfCancellationRequested();
                errors.Add(projection.Errors[errorIndex++]);
            }

            yield return new RowEvent(row, items, errors);
        }
    }

    /// <summary>
    /// 按任意源行顺序创建兼容投影，处理需要回看的结果集合。
    /// </summary>
    private static IEnumerable<RowEvent> EnumerateIndexedEvents(RegionProjection projection,
        CancellationToken cancellationToken)
    {
        var rows = new SortedSet<int>();
        var itemsByRow = new Dictionary<int, List<object>>();
        for (var index = 0; index < projection.SourceRows.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = projection.SourceRows[index] + 1;
            rows.Add(row);
            if (!itemsByRow.TryGetValue(row, out var items))
                itemsByRow[row] = items = new List<object>();
            items.Add(projection.Items[index]);
        }

        var errorsByRow = new Dictionary<int, List<ExcelImportError>>();
        foreach (var error in projection.Errors)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = error.RowIndex > 0 ? error.RowIndex : projection.DataRow;
            if (!errorsByRow.TryGetValue(row, out var rowErrors))
                errorsByRow[row] = rowErrors = new List<ExcelImportError>();
            rowErrors.Add(error);
        }
        foreach (var row in errorsByRow.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rows.Add(row);
        }

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return new RowEvent(row,
                itemsByRow.TryGetValue(row, out var items) ? items : Array.Empty<object>(),
                errorsByRow.TryGetValue(row, out var errors) ? errors : Array.Empty<ExcelImportError>());
        }
    }

    /// <summary>
    /// 获取错误对应的一基源行号。
    /// </summary>
    private static int NormalizeErrorRow(ExcelImportError error, int dataRow) =>
        error.RowIndex > 0 ? error.RowIndex : dataRow;

    /// <summary>
    /// 保存一个列表区域的批次投影。
    /// </summary>
    private sealed class RegionProjection
    {
        internal RegionProjection(int regionIndex, string sheetName, string propertyName,
            Type itemType, int startRow, bool includeHeader, IReadOnlyList<int> sourceRows,
            IReadOnlyList<object> items, IReadOnlyList<ExcelImportError> errors, bool errorsTruncated)
        {
            RegionIndex = regionIndex;
            SheetName = sheetName;
            PropertyName = propertyName;
            ItemType = itemType;
            SourceRows = sourceRows;
            Items = items;
            Errors = errors;
            ErrorsTruncated = errorsTruncated;
            DataRow = startRow + (includeHeader ? 2 : 1);
        }

        internal int RegionIndex { get; }
        internal string SheetName { get; }
        internal string PropertyName { get; }
        internal Type ItemType { get; }
        internal IReadOnlyList<int> SourceRows { get; }
        internal IReadOnlyList<object> Items { get; }
        internal IReadOnlyList<ExcelImportError> Errors { get; }
        internal bool ErrorsTruncated { get; }
        internal int DataRow { get; }

        internal ExcelBatchImportSummary CreateSummary(int rowsRead, int failedRows)
        {
            return new ExcelBatchImportSummary(rowsRead, Items.Count, failedRows,
                Errors.Count, ErrorsTruncated,
                Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit));
        }
    }

    /// <summary>
    /// 保存一个物理行上的成功项和错误。
    /// </summary>
    private sealed class RowEvent
    {
        internal RowEvent(int row, IReadOnlyList<object> items, IReadOnlyList<ExcelImportError> errors)
        {
            Row = row;
            Items = items;
            Errors = errors;
        }

        internal int Row { get; }
        internal IReadOnlyList<object> Items { get; }
        internal IReadOnlyList<ExcelImportError> Errors { get; }
    }

    /// <summary>
    /// 保存异步批次编号。
    /// </summary>
    private sealed class BatchNumberState
    {
        internal BatchNumberState(int value)
        {
            Value = value;
        }

        internal int Value { get; set; }
    }
}
