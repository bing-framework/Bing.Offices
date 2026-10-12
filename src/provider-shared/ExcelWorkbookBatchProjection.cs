using System.Collections;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Imports;

/// <summary>
/// 将完整 Workbook 导入结果投影为批次回调。
/// </summary>
internal static class ExcelWorkbookBatchProjection
{
    /// <summary>
    /// 按物理工作表顺序交付工作簿批次。
    /// </summary>
    /// <remarks>
    /// 通过同步回调交付批次。
    /// </remarks>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="result">完整 Workbook 导入结果。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="physicalSheetNames">源工作簿中的物理工作表名称。</param>
    /// <param name="onBatch">批次回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>Workbook 批次摘要。</returns>
    internal static ExcelWorkbookBatchImportSummary Deliver<TWorkbook>(
        ExcelWorkbookImportResult<TWorkbook> result,
        ExcelWorkbookImportRequest<TWorkbook> request,
        IReadOnlyList<string> physicalSheetNames,
        Action<ExcelWorkbookImportBatch> onBatch,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        if (result == null)
            throw new ArgumentNullException(nameof(result));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));

        var projections = CreateProjections(result, request, physicalSheetNames);
        var summaries = new List<ExcelBatchImportSheetSummary>(projections.Count);
        var batchNumber = 1;
        for (var index = 0; index < projections.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var projection = projections[index];
            var summary = DeliverSheet(projection, ref batchNumber, onBatch, cancellationToken);
            summaries.Add(new ExcelBatchImportSheetSummary(projection.Name, projection.ItemType, summary));
        }
        return new ExcelWorkbookBatchImportSummary(summaries);
    }

    /// <summary>
    /// 按物理工作表顺序交付工作簿批次。
    /// </summary>
    /// <remarks>
    /// 通过异步回调交付批次。
    /// </remarks>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="result">完整 Workbook 导入结果。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="physicalSheetNames">源工作簿中的物理工作表名称。</param>
    /// <param name="onBatch">批次回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>工作簿批次摘要任务。</returns>
    internal static async Task<ExcelWorkbookBatchImportSummary> DeliverAsync<TWorkbook>(
        ExcelWorkbookImportResult<TWorkbook> result,
        ExcelWorkbookImportRequest<TWorkbook> request,
        IReadOnlyList<string> physicalSheetNames,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        if (result == null)
            throw new ArgumentNullException(nameof(result));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));

        var projections = CreateProjections(result, request, physicalSheetNames);
        var summaries = new List<ExcelBatchImportSheetSummary>(projections.Count);
        var batchNumber = 1;
        for (var index = 0; index < projections.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var projection = projections[index];
            var batchState = new BatchNumberState(batchNumber);
            var summary = await DeliverSheetAsync(projection, batchState, onBatch,
                cancellationToken).ConfigureAwait(false);
            batchNumber = batchState.Value;
            summaries.Add(new ExcelBatchImportSheetSummary(projection.Name, projection.ItemType, summary));
        }
        return new ExcelWorkbookBatchImportSummary(summaries);
    }

    /// <summary>
    /// 根据完整导入结果创建各工作表投影并排序。
    /// </summary>
    private static IReadOnlyList<SheetProjection> CreateProjections<TWorkbook>(
        ExcelWorkbookImportResult<TWorkbook> result,
        ExcelWorkbookImportRequest<TWorkbook> request,
        IReadOnlyList<string> physicalSheetNames)
        where TWorkbook : class, new()
    {
        var comparer = request.SheetNameComparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var projections = new List<SheetProjection>(result.Sheets.Count);
        for (var index = 0; index < result.Sheets.Count; index++)
        {
            var sheet = result.Sheets[index];
            var sheetRequest = FindRequest(request, sheet, physicalSheetNames, comparer, index);
            var sourceRows = sheet.SourceRows?.ToArray() ?? Array.Empty<int>();
            var items = ReadImportedItems(result.Workbook, sheetRequest, sourceRows.Length, sheet.Name);
            var errors = result.Errors
                .Where(error => string.Equals(error.SheetName, sheet.Name, comparer))
                .ToArray();
            projections.Add(new SheetProjection(sheet.Name, sheet.ItemType,
                sheetRequest, sourceRows, items, errors, result.ErrorsTruncated));
        }

        return projections
            .Select((projection, index) => new { Projection = projection, Index = index,
                PhysicalIndex = FindPhysicalIndex(physicalSheetNames, projection.Name, comparer) })
            .OrderBy(item => item.PhysicalIndex < 0 ? int.MaxValue : item.PhysicalIndex)
            .ThenBy(item => item.Index)
            .Select(item => item.Projection)
            .ToArray();
    }

    /// <summary>
    /// 查找完整导入结果对应的 Sheet 请求。
    /// </summary>
    private static ExcelSheetImportRequest FindRequest<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request,
        ExcelSheetImportResult sheet,
        IReadOnlyList<string> physicalSheetNames,
        StringComparison comparison,
        int resultIndex)
        where TWorkbook : class, new()
    {
        foreach (var candidate in request.Sheets)
        {
            if (candidate.ItemType != sheet.ItemType)
                continue;
            if (candidate.Selector.Kind == ExcelSheetSelectorKind.ByName
                && string.Equals(candidate.Selector.Name, sheet.Name, comparison))
                return candidate;
            if (candidate.Selector.Kind == ExcelSheetSelectorKind.ByIndex
                && candidate.Selector.Index >= 0
                && candidate.Selector.Index < physicalSheetNames.Count
                && string.Equals(physicalSheetNames[candidate.Selector.Index.Value], sheet.Name, comparison))
                return candidate;
        }

        return resultIndex >= 0 && resultIndex < request.Sheets.Count
            ? request.Sheets[resultIndex] : null;
    }

    /// <summary>
    /// 读取请求目标集合中本次导入追加的实体。
    /// </summary>
    private static IReadOnlyList<object> ReadImportedItems<TWorkbook>(TWorkbook workbook,
        ExcelSheetImportRequest request, int count, string sheetName)
        where TWorkbook : class, new()
    {
        if (count == 0)
            return Array.Empty<object>();
        if (request == null)
            throw new BingOfficesConfigurationException($"无法定位工作表批次目标集合: {sheetName}",
                stage: BingOfficesStage.Plan);
        var target = request.Target(workbook) as IEnumerable;
        if (target == null)
            throw new BingOfficesConfigurationException($"Workbook 导入目标集合不可枚举: {sheetName}",
                stage: BingOfficesStage.Plan);
        var all = target.Cast<object>().ToArray();
        if (all.Length < count)
            throw new BingOfficesConfigurationException($"Workbook 导入目标集合未包含完整批次: {sheetName}",
                stage: BingOfficesStage.Plan);
        return all.Skip(all.Length - count).ToArray();
    }

    /// <summary>
    /// 同步交付单个工作表的批次。
    /// </summary>
    private static ExcelBatchImportSummary DeliverSheet(SheetProjection projection, ref int batchNumber,
        Action<ExcelWorkbookImportBatch> onBatch, CancellationToken cancellationToken)
    {
        var events = CreateEvents(projection);
        var items = new List<object>();
        var errors = new List<ExcelImportError>();
        var firstRow = 0;
        foreach (var item in events)
        {
            cancellationToken.ThrowIfCancellationRequested();
            firstRow = firstRow == 0 ? item.Row : firstRow;
            items.AddRange(item.Items);
            errors.AddRange(item.Errors);
            if (items.Count >= projection.BatchSize || errors.Count >= projection.BatchSize)
            {
                onBatch(new ExcelWorkbookImportBatch(projection.Name, projection.ItemType, batchNumber++,
                    firstRow, items.ToArray(), errors.ToArray()));
                items.Clear();
                errors.Clear();
                firstRow = 0;
            }
        }
        if (items.Count > 0 || errors.Count > 0)
        {
            onBatch(new ExcelWorkbookImportBatch(projection.Name, projection.ItemType, batchNumber++,
                firstRow, items.ToArray(), errors.ToArray()));
        }
        return projection.CreateSummary(events);
    }

    /// <summary>
    /// 异步交付单个工作表的批次。
    /// </summary>
    private static async Task<ExcelBatchImportSummary> DeliverSheetAsync(SheetProjection projection,
        BatchNumberState batchState, Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken)
    {
        var events = CreateEvents(projection);
        var items = new List<object>();
        var errors = new List<ExcelImportError>();
        var firstRow = 0;
        foreach (var item in events)
        {
            cancellationToken.ThrowIfCancellationRequested();
            firstRow = firstRow == 0 ? item.Row : firstRow;
            items.AddRange(item.Items);
            errors.AddRange(item.Errors);
            if (items.Count >= projection.BatchSize || errors.Count >= projection.BatchSize)
            {
                await onBatch(new ExcelWorkbookImportBatch(projection.Name, projection.ItemType,
                    batchState.Value++, firstRow, items.ToArray(), errors.ToArray()), cancellationToken)
                    .ConfigureAwait(false);
                items.Clear();
                errors.Clear();
                firstRow = 0;
            }
        }
        if (items.Count > 0 || errors.Count > 0)
        {
            await onBatch(new ExcelWorkbookImportBatch(projection.Name, projection.ItemType,
                batchState.Value++, firstRow, items.ToArray(), errors.ToArray()), cancellationToken)
                .ConfigureAwait(false);
        }
        return projection.CreateSummary(events);
    }

    /// <summary>
    /// 将成功实体和错误按源行合并为事件。
    /// </summary>
    private static IReadOnlyList<RowEvent> CreateEvents(SheetProjection projection)
    {
        var rows = new SortedSet<int>();
        var itemsByRow = new Dictionary<int, List<object>>();
        for (var index = 0; index < projection.SourceRows.Count; index++)
        {
            var row = projection.SourceRows[index] + 1;
            rows.Add(row);
            if (!itemsByRow.TryGetValue(row, out var items))
                itemsByRow[row] = items = new List<object>();
            items.Add(projection.Items[index]);
        }

        var errorsByRow = projection.Errors
            .GroupBy(error => error.RowIndex > 0 ? error.RowIndex : projection.HeaderRow)
            .ToDictionary(group => group.Key, group => (IReadOnlyList<ExcelImportError>)group.ToArray());
        foreach (var row in errorsByRow.Keys)
            rows.Add(row);

        return rows.Select(row => new RowEvent(row,
            itemsByRow.TryGetValue(row, out var items) ? items : Array.Empty<object>(),
            errorsByRow.TryGetValue(row, out var errors) ? errors : Array.Empty<ExcelImportError>()))
            .ToArray();
    }

    /// <summary>
    /// 查找工作表在源工作簿中的物理索引。
    /// </summary>
    private static int FindPhysicalIndex(IReadOnlyList<string> physicalSheetNames,
        string sheetName, StringComparison comparison)
    {
        if (physicalSheetNames == null)
            return -1;
        for (var index = 0; index < physicalSheetNames.Count; index++)
        {
            if (string.Equals(physicalSheetNames[index], sheetName, comparison))
                return index;
        }
        return -1;
    }

    /// <summary>
    /// 保存单个工作表的批次投影。
    /// </summary>
    private sealed class SheetProjection
    {
        /// <summary>
        /// 初始化一个 <see cref="SheetProjection" /> 类型的实例。
        /// </summary>
        /// <param name="name">工作表名称。</param>
        /// <param name="itemType">工作表实体类型。</param>
        /// <param name="request">工作表导入请求。</param>
        /// <param name="sourceRows">成功读取实体对应的源行号。</param>
        /// <param name="items">成功导入的实体。</param>
        /// <param name="errors">工作表错误。</param>
        /// <param name="errorsTruncated">错误是否因预算而截断。</param>
        internal SheetProjection(string name, Type itemType, ExcelSheetImportRequest request,
            IReadOnlyList<int> sourceRows, IReadOnlyList<object> items,
            IReadOnlyList<ExcelImportError> errors, bool errorsTruncated)
        {
            Name = name;
            ItemType = itemType;
            Request = request;
            SourceRows = sourceRows;
            Items = items;
            Errors = errors;
            ErrorsTruncated = errorsTruncated;
            BatchSize = request?.BatchSize > 0 ? request.BatchSize : 1000;
            HeaderRow = request == null ? 1 : request.HeaderRowIndex + 1;
            DataRow = request == null ? 2 : request.DataRowStartIndex + 1;
        }

        /// <summary>
        /// 获取工作表名称。
        /// </summary>
        internal string Name { get; }
        /// <summary>
        /// 获取工作表实体类型。
        /// </summary>
        internal Type ItemType { get; }
        /// <summary>
        /// 获取工作表导入请求。
        /// </summary>
        internal ExcelSheetImportRequest Request { get; }
        /// <summary>
        /// 获取成功读取实体对应的源行号。
        /// </summary>
        internal IReadOnlyList<int> SourceRows { get; }
        /// <summary>
        /// 获取成功导入的实体。
        /// </summary>
        internal IReadOnlyList<object> Items { get; }
        /// <summary>
        /// 获取工作表错误。
        /// </summary>
        internal IReadOnlyList<ExcelImportError> Errors { get; }
        /// <summary>
        /// 获取错误是否因预算而截断。
        /// </summary>
        internal bool ErrorsTruncated { get; }
        /// <summary>
        /// 获取单批实体数量上限。
        /// </summary>
        internal int BatchSize { get; }
        /// <summary>
        /// 获取一基表头行号。
        /// </summary>
        internal int HeaderRow { get; }
        /// <summary>
        /// 获取一基数据起始行号。
        /// </summary>
        internal int DataRow { get; }

        /// <summary>
        /// 根据源行事件创建工作表导入摘要。
        /// </summary>
        /// <param name="events">按源行合并的导入事件。</param>
        /// <returns>工作表导入摘要。</returns>
        internal ExcelBatchImportSummary CreateSummary(IReadOnlyList<RowEvent> events)
        {
            var dataRows = events.Where(item => item.Row >= DataRow).Select(item => item.Row).Distinct().ToArray();
            var failedRows = events.Where(item => item.Row >= DataRow && item.Errors.Count > 0)
                .Select(item => item.Row).Distinct().Count();
            return new ExcelBatchImportSummary(dataRows.Length, Items.Count, failedRows, Errors.Count,
                ErrorsTruncated, Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit));
        }
    }

    /// <summary>
    /// 保存一个源行上的成功实体和错误。
    /// </summary>
    private sealed class RowEvent
    {
        /// <summary>
        /// 初始化一个 <see cref="RowEvent" /> 类型的实例。
        /// </summary>
        /// <param name="row">一基源行号。</param>
        /// <param name="items">该行成功导入的实体。</param>
        /// <param name="errors">该行错误。</param>
        internal RowEvent(int row, IReadOnlyList<object> items, IReadOnlyList<ExcelImportError> errors)
        {
            Row = row;
            Items = items;
            Errors = errors;
        }

        /// <summary>
        /// 获取一基源行号。
        /// </summary>
        internal int Row { get; }
        /// <summary>
        /// 获取该行成功导入的实体。
        /// </summary>
        internal IReadOnlyList<object> Items { get; }
        /// <summary>
        /// 获取该行错误。
        /// </summary>
        internal IReadOnlyList<ExcelImportError> Errors { get; }
    }

    /// <summary>
    /// 保存异步批次编号状态。
    /// </summary>
    private sealed class BatchNumberState
    {
        /// <summary>
        /// 初始化一个 <see cref="BatchNumberState" /> 类型的实例。
        /// </summary>
        /// <param name="value">当前批次号。</param>
        internal BatchNumberState(int value)
        {
            Value = value;
        }

        /// <summary>
        /// 获取或设置当前批次号。
        /// </summary>
        internal int Value { get; set; }
    }
}
