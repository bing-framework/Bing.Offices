using System.Collections;
using System.Collections.Concurrent;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Bing.Offices.Entities;

/// <summary>
/// NPOI 单个实体导出执行器。
/// </summary>
internal sealed class NpoiEntityExportExecutor
{
    /// <summary>
    /// 记录实体列表区域在运行时占用的工作表矩形。
    /// </summary>
    private sealed record EntityListFootprint(string SheetName, long FirstRow, long LastRow,
        long FirstColumn, long LastColumn)
    {
        /// <summary>
        /// 判断当前矩形是否与另一个矩形相交。
        /// </summary>
        /// <param name="other">待比较的矩形。</param>
        /// <returns>两个矩形相交时返回 <see langword="true" />。</returns>
        public bool Intersects(EntityListFootprint other) =>
            string.Equals(SheetName, other.SheetName, StringComparison.OrdinalIgnoreCase)
            && FirstRow <= other.LastRow && other.FirstRow <= LastRow
            && FirstColumn <= other.LastColumn && other.FirstColumn <= LastColumn;
    }

    /// <summary>
    /// NPOI 导出映射计划构建器。
    /// </summary>
    private readonly NpoiExportPlanBuilder _planBuilder;

    /// <summary>
    /// NPOI 工作表数据写入器。
    /// </summary>
    private readonly NpoiExportSheetWriter _sheetWriter = new NpoiExportSheetWriter();

    /// <summary>
    /// 实体属性值转换器集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityExportExecutor" /> 类型的实例。
    /// </summary>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="valueConverters">实体属性值转换器集合。</param>
    internal NpoiEntityExportExecutor(IExcelMappingPlanFactory mappingPlanFactory,
        IReadOnlyList<IExcelValueConverter> valueConverters)
    {
        _planBuilder = new NpoiExportPlanBuilder(mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory)));
        _valueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
    }

    /// <summary>
    /// 按实体布局将固定单元格、列表区域和合并区域写入工作簿。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局定义。</param>
    /// <param name="template">是否在既有模板上执行写入。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    internal void Write<TEntity>(IWorkbook workbook, TEntity entity, ExcelEntityLayout<TEntity> layout,
        bool template, CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (entity == null)
            throw new ArgumentNullException(nameof(entity));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        var sheetNames = layout.Cells.Select(item => item.SheetName)
            .Concat(layout.ListRegions.Select(item => item.SheetName))
            .Concat(layout.Merges.Select(item => item.SheetName))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        foreach (var name in sheetNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, name, template);
            NpoiEntityLayoutSupport.PreflightMerges(sheet, layout.Merges, true, false);
        }
        foreach (var binding in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, binding.SheetName, template);
            var resolvedBinding = binding;
            if (binding.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, binding.SheetName,
                    binding.AnchorName);
                sheet = named.Sheet;
                resolvedBinding = binding.WithResolvedReference(named.Reference);
            }
            var anchor = NpoiEntityLayoutSupport.ResolveAnchor(sheet, resolvedBinding.Reference);
            var cell = NpoiEntityLayoutSupport.GetCell(sheet, anchor, true);
            var value = resolvedBinding.Getter(entity);
            var column = NpoiEntityLayoutSupport.CreateCellPlan(_planBuilder.MappingPlanFactory, resolvedBinding,
                MappingDirection.Export);
            var converted = column.ConvertTo(value, sheet.SheetName, anchor.Row + 1, anchor.Column + 1,
                CultureInfo.InvariantCulture);
            NpoiEntityLayoutSupport.ValidateCellExport(column, value, converted, sheet.SheetName, anchor);
            column.WriteValue(cell, converted);
        }
        var footprints = new List<EntityListFootprint>();
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, region.SheetName, template);
            var resolvedRegion = region;
            if (region.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, region.SheetName,
                    region.AnchorName);
                sheet = named.Sheet;
                resolvedRegion = region.WithResolvedStart(named.Reference);
            }
            NpoiEntityLayoutSupport.ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            var values = resolvedRegion.Getter(entity)?.Cast<object>().ToArray() ?? Array.Empty<object>();
            WriteRegion(workbook, sheet, entity, resolvedRegion, values, footprints, cancellationToken);
        }
    }

    /// <summary>
    /// 根据列表区域的运行时元素类型分派强类型写入逻辑。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="values">列表元素集合。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values,
        ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(WriteTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, region, values, footprints,
            cancellationToken });
    }

    /// <summary>
    /// 将强类型列表元素写入布局声明的工作表区域。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="values">列表元素集合。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values,
        ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        var data = values.Cast<TItem>().ToArray();
        var groups = region.DynamicColumnGroups ?? Array.Empty<IExcelEntityDynamicColumnGroup>();
        var groupDefinitions = groups.SelectMany(group => group.Definitions).ToArray();
        var request = ExcelExport.Workbook(builder => builder.AddSheet(region.SheetName, data, configure =>
        {
            configure.HeaderRowIndex(0).DataRowStartIndex(1);
            if (region.MappingDocument != null)
                configure.Mapping(region.MappingDocument);
            else if (region.MappingConfiguration != null)
                configure.Mapping(region.MappingConfiguration);
            if (groupDefinitions.Length > 0)
                configure.DynamicColumns(CreateDynamicGetter<TItem>(groups), groupDefinitions)
                    .UnknownDynamicValues(region.UnknownDynamicValues);
        }));
        var sheetRequest = request.Sheets[0];
        var mapping = _planBuilder.Create(request)[sheetRequest];
        var dynamicProperty = groups.Count == 0 ? GetDynamicProperty<TItem>(mapping) : null;
        var dynamicDefinitions = groupDefinitions.Length > 0
            ? groupDefinitions : CreateDynamicDefinitions(mapping);
        var columns = groups.Count == 0
            ? NpoiExportColumnPlanner.CreateColumns<TItem>(mapping, dynamicDefinitions,
                region.CalculatedColumns)
            : NpoiExportColumnPlanner.CreateEntityColumns<TItem>(mapping, dynamicDefinitions, groups,
                region.CalculatedColumns);
        if (region.GroupSubtotal != null)
        {
            WriteGroupedTypedRegion(workbook, sheet, region, data, groups, dynamicProperty,
                dynamicDefinitions, columns, footprints, cancellationToken);
            return;
        }
        NpoiEntityLayoutSupport.ValidateListRegionBounds(region, columns.Count, data.Length, region.IncludeHeader);
        var dataEndRow = region.PageSubtotal != null
            ? NpoiEntityLayoutSupport.ValidatePagedFooterBounds(workbook, region, columns.Count,
                data.Length)
            : region.Start.Row + (region.IncludeHeader ? 1 : 0) + data.Length;
        NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, region.Footer, dataEndRow);
        RegisterRuntimeFootprint(region, columns.Count, data.Length, footprints);
        cancellationToken.ThrowIfCancellationRequested();
        if (region.PageSubtotal != null)
        {
            WritePagedTypedRegion(workbook, sheet, region, data, groups, dynamicProperty,
                dynamicDefinitions, columns, cancellationToken);
            return;
        }
        if (region.IncludeHeader)
        {
            if (dynamicDefinitions.Count > 0 && groups.Count == 0)
            {
                var dynamicGetter = CreateDynamicGetter<TItem>(dynamicProperty);
                request = ExcelExport.Workbook(builder => builder.AddSheet(region.SheetName, data, configure =>
                {
                    configure.HeaderRowIndex(0).DataRowStartIndex(1);
                    if (region.MappingDocument != null)
                        configure.Mapping(region.MappingDocument);
                    else if (region.MappingConfiguration != null)
                        configure.Mapping(region.MappingConfiguration);
                    configure.DynamicColumns(dynamicGetter, dynamicDefinitions)
                        .UnknownDynamicValues(region.UnknownDynamicValues);
                }));
                sheetRequest = request.Sheets[0];
            }
            _sheetWriter.Write<TItem>(workbook, sheetRequest, cancellationToken, mapping, columns,
                region.Start.Row, region.Start.Column);
            WriteFooter(workbook, sheet, region.Footer, values,
                dataEndRow, region.Start.Column, cancellationToken, region.FooterAnchorName);
            NpoiEntityLayoutSupport.ApplyPageBreaks(sheet, region, data.Length);
            return;
        }
        var rowIndex = region.Start.Row;
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray()
            : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        for (var itemIndex = 0; itemIndex < data.Length; itemIndex++)
        {
            var item = data[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = groups.Count == 0
                ? dynamicProperty?.GetValue(item) as IDictionary<string, object>
                : MergeDynamicValues(item, groups);
            if (region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail)
            {
                var knownKeys = new HashSet<string>(dynamicDefinitions.Select(definition => definition.Key),
                    StringComparer.OrdinalIgnoreCase);
                var unknown = dynamicValues?.Keys.FirstOrDefault(key => !knownKeys.Contains(key));
                if (unknown != null)
                    throw new BingOfficesConfigurationException($"动态值未声明: {unknown}",
                        stage: BingOfficesStage.Validate);
            }
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var cell = row.GetCell(region.Start.Column + index) ?? row.CreateCell(region.Start.Column + index);
                NpoiExportSheetWriter.WriteCell(cell, item, column, dynamicValues, sheet.SheetName,
                    rowIndex + 1, region.Start.Column + index + 1, CultureInfo.InvariantCulture,
                    itemIndex, calculatedEvaluators.TryGetValue(column.Key, out var evaluator)
                        ? evaluator : null);
                if (column.BodyStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, column.BodyStyle);
            }
            rowIndex++;
        }
        WriteFooter(workbook, sheet, region.Footer, values, rowIndex, region.Start.Column,
            cancellationToken, region.FooterAnchorName);
        NpoiEntityLayoutSupport.ApplyPageBreaks(sheet, region, data.Length);
    }

    /// <summary>
    /// 将带分页小计的强类型列表写入 NPOI 工作表。
    /// </summary>
    private static void WritePagedTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<TItem> data,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups, PropertyInfo dynamicProperty,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicDefinitions,
        IReadOnlyList<ExcelColumnPlan> columns, CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        if (region.PageSubtotal == null || region.PageBreakRows is not int rowsPerPage)
            throw new BingOfficesConfigurationException("分页小计缺少分页配置。",
                stage: BingOfficesStage.Plan);
        var plannedRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var plannedIndex = 0;
        while (plannedIndex < data.Count)
        {
            var pageCount = Math.Min(rowsPerPage, data.Count - plannedIndex);
            plannedRow += pageCount;
            plannedIndex += pageCount;
            if (plannedIndex < data.Count)
            {
                NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, region.PageSubtotal,
                    plannedRow, region.Start.Column);
                plannedRow += region.PageSubtotal.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.PageSubtotal);
            }
        }
        NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, region.Footer, plannedRow,
            region.Start.Column);
        if (region.IncludeHeader)
        {
            var header = sheet.GetRow(region.Start.Row) ?? sheet.CreateRow(region.Start.Row);
            for (var index = 0; index < columns.Count; index++)
            {
                var cell = header.GetCell(region.Start.Column + index)
                    ?? header.CreateCell(region.Start.Column + index);
                cell.SetCellValue(columns[index].Title);
                if (columns[index].HeaderStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle,
                        columns[index].HeaderStyle);
            }
        }
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray() : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        var rowIndex = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var itemIndex = 0;
        var detailRows = new List<(int FirstRow, int LastRow)>();
        while (itemIndex < data.Count)
        {
            var firstDetailRow = rowIndex + 1;
            var pageStart = itemIndex;
            var pageCount = Math.Min(rowsPerPage, data.Count - itemIndex);
            for (var pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                var item = data[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                if (item == null)
                    throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                        stage: BingOfficesStage.Plan);
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var dynamicValues = groups.Count == 0
                    ? dynamicProperty?.GetValue(item) as IDictionary<string, object>
                    : MergeDynamicValues(item, groups);
                if (region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail)
                {
                    var knownKeys = new HashSet<string>(dynamicDefinitions.Select(definition => definition.Key),
                        StringComparer.OrdinalIgnoreCase);
                    var unknown = dynamicValues?.Keys.FirstOrDefault(key => !knownKeys.Contains(key));
                    if (unknown != null)
                        throw new BingOfficesConfigurationException($"动态值未声明: {unknown}",
                            stage: BingOfficesStage.Validate);
                }
                for (var index = 0; index < columns.Count; index++)
                {
                    var column = columns[index];
                    var cell = row.GetCell(region.Start.Column + index)
                        ?? row.CreateCell(region.Start.Column + index);
                    NpoiExportSheetWriter.WriteCell(cell, item, column, dynamicValues, sheet.SheetName,
                        rowIndex + 1, region.Start.Column + index + 1, CultureInfo.InvariantCulture,
                        itemIndex, calculatedEvaluators.TryGetValue(column.Key, out var evaluator)
                            ? evaluator : null);
                    if (column.BodyStyle != null)
                        cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, column.BodyStyle);
                }
                rowIndex++;
                itemIndex++;
            }
            detailRows.Add((firstDetailRow, rowIndex));
            if (itemIndex < data.Count)
            {
                var pageItems = data.Skip(pageStart).Take(pageCount).Cast<object>().ToArray();
                WriteFooter(workbook, sheet, region.PageSubtotal, pageItems, rowIndex,
                    region.Start.Column, cancellationToken);
                rowIndex += region.PageSubtotal.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.PageSubtotal);
                NpoiEntityLayoutSupport.ApplyPageBreak(sheet, rowIndex - 1);
            }
        }
        WriteFooter(workbook, sheet, region.Footer, data.Cast<object>().ToArray(), rowIndex,
            region.Start.Column, cancellationToken, region.FooterAnchorName, detailRows);
    }

    /// <summary>
    /// 将带连续分组小计的列表项写入 NPOI 工作表。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="data">列表元素集合。</param>
    /// <param name="groups">动态列组集合。</param>
    /// <param name="dynamicProperty">单字典动态属性。</param>
    /// <param name="dynamicDefinitions">动态列定义。</param>
    /// <param name="columns">最终列计划。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void WriteGroupedTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<TItem> data,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups, PropertyInfo dynamicProperty,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicDefinitions,
        IReadOnlyList<ExcelColumnPlan> columns, ICollection<EntityListFootprint> footprints,
        CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        var subtotal = region.GroupSubtotal;
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var rowIndex = firstDataRow;
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray() : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        var groupRanges = GetGroupedRanges(data, subtotal);
        var previewRow = firstDataRow;
        foreach (var range in groupRanges)
        {
            var dataEndRow = previewRow + range.Count;
            NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, subtotal.Footer, dataEndRow);
            previewRow = dataEndRow + subtotal.Footer.GapRows
                + NpoiEntityLayoutSupport.GetFooterHeight(subtotal.Footer);
        }
        NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, region.Footer, previewRow);
        var lastRow = previewRow + (region.Footer == null ? -1
            : region.Footer.GapRows + NpoiEntityLayoutSupport.GetFooterHeight(region.Footer) - 1);
        if (region.End.HasValue && (lastRow > region.End.Value.Row
            || region.Start.Column + columns.Count - 1 > region.End.Value.Column))
            throw new BingOfficesConfigurationException(
                $"列表区域分组小计超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        RegisterRuntimeFootprint(region, columns.Count, data.Count, footprints, data.Cast<object>().ToArray());
        cancellationToken.ThrowIfCancellationRequested();

        if (region.IncludeHeader)
        {
            var header = sheet.GetRow(region.Start.Row) ?? sheet.CreateRow(region.Start.Row);
            for (var index = 0; index < columns.Count; index++)
            {
                var cell = header.GetCell(region.Start.Column + index)
                    ?? header.CreateCell(region.Start.Column + index);
                cell.SetCellValue(columns[index].Title);
                if (columns[index].HeaderStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle,
                        columns[index].HeaderStyle);
            }
        }

        void WriteItem(TItem item, int itemIndex)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = groups.Count == 0
                ? dynamicProperty?.GetValue(item) as IDictionary<string, object>
                : MergeDynamicValues(item, groups);
            if (region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail)
            {
                var knownKeys = new HashSet<string>(dynamicDefinitions.Select(definition => definition.Key),
                    StringComparer.OrdinalIgnoreCase);
                var unknown = dynamicValues?.Keys.FirstOrDefault(key => !knownKeys.Contains(key));
                if (unknown != null)
                    throw new BingOfficesConfigurationException($"动态值未声明: {unknown}",
                        stage: BingOfficesStage.Validate);
            }
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var cell = row.GetCell(region.Start.Column + index)
                    ?? row.CreateCell(region.Start.Column + index);
                NpoiExportSheetWriter.WriteCell(cell, item, column, dynamicValues, sheet.SheetName,
                    rowIndex + 1, region.Start.Column + index + 1, CultureInfo.InvariantCulture,
                    itemIndex, calculatedEvaluators.TryGetValue(column.Key, out var evaluator)
                        ? evaluator : null);
                if (column.BodyStyle != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, column.BodyStyle);
            }
            rowIndex++;
        }

        var itemIndexOffset = 0;
        var detailRows = new List<(int FirstRow, int LastRow)>();
        foreach (var range in groupRanges)
        {
            var firstDetailRow = rowIndex + 1;
            var groupItems = data.Skip(range.Start).Take(range.Count).Cast<object>().ToArray();
            for (var itemIndex = 0; itemIndex < range.Count; itemIndex++)
                WriteItem(data[range.Start + itemIndex], itemIndexOffset + itemIndex);
            detailRows.Add((firstDetailRow, rowIndex));
            WriteFooter(workbook, sheet, subtotal.Footer, groupItems, rowIndex,
                region.Start.Column, cancellationToken);
            rowIndex += subtotal.Footer.GapRows + NpoiEntityLayoutSupport.GetFooterHeight(subtotal.Footer);
            itemIndexOffset += range.Count;
        }
        WriteFooter(workbook, sheet, region.Footer, data.Cast<object>().ToArray(), rowIndex,
            region.Start.Column, cancellationToken, region.FooterAnchorName, detailRows);
    }

    /// <summary>
    /// 按相邻分组键计算列表元素的连续区间。
    /// </summary>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="data">列表元素集合。</param>
    /// <param name="subtotal">分组小计定义。</param>
    /// <returns>连续分组的起始索引和元素数量。</returns>
    private static IReadOnlyList<(int Start, int Count)> GetGroupedRanges<TItem>(
        IReadOnlyList<TItem> data, IExcelEntityGroupSubtotal subtotal) where TItem : class, new()
    {
        var ranges = new List<(int Start, int Count)>();
        if (data == null || data.Count == 0)
            return ranges;
        var start = 0;
        var key = subtotal.GetKey(data[0]);
        for (var index = 1; index < data.Count; index++)
        {
            var next = subtotal.GetKey(data[index]);
            if (Equals(key, next))
                continue;
            ranges.Add((start, index - start));
            start = index;
            key = next;
        }
        ranges.Add((start, data.Count - start));
        return ranges;
    }

    /// <summary>
    /// 预检并登记列表区域的运行时占用范围。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="columnCount">映射计划的列数。</param>
    /// <param name="itemCount">实际明细行数。</param>
    /// <param name="footprints">已登记的列表区域范围。</param>
    /// <param name="values">列表项快照；配置分组小计时用于计算运行时行数。</param>
    private static void RegisterRuntimeFootprint<TEntity>(ExcelEntityListRegion<TEntity> region,
        int columnCount, int itemCount, ICollection<EntityListFootprint> footprints,
        IReadOnlyList<object> values = null)
        where TEntity : class, new()
    {
        var dataRows = (long)itemCount + (region.IncludeHeader ? 1 : 0);
        var hasData = dataRows > 0;
        var firstRow = (long)region.Start.Row;
        var lastRow = hasData ? firstRow + dataRows - 1 : firstRow - 1;
        var firstColumn = (long)region.Start.Column;
        var lastColumn = hasData ? firstColumn + Math.Max(0, columnCount - 1) : firstColumn - 1;
        if (region.GroupSubtotal != null)
        {
            var row = (long)region.Start.Row + (region.IncludeHeader ? 1 : 0);
            foreach (var range in GetGroupedRanges(values ?? Array.Empty<object>(), region.GroupSubtotal))
                row += range.Count + region.GroupSubtotal.Footer.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.GroupSubtotal.Footer);
            lastRow = Math.Max(lastRow, row - 1);
            foreach (var cell in region.GroupSubtotal.Footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
            foreach (var merge in region.GroupSubtotal.Footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
                lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
        }
        if (region.PageSubtotal != null && region.PageBreakRows is int rowsPerPage)
        {
            var row = (long)region.Start.Row + (region.IncludeHeader ? 1 : 0);
            var written = 0;
            while (written < itemCount)
            {
                var count = Math.Min(rowsPerPage, itemCount - written);
                row += count;
                written += count;
                if (written >= itemCount)
                    break;
                row += region.PageSubtotal.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(region.PageSubtotal);
                lastRow = Math.Max(lastRow, row - 1);
                foreach (var cell in region.PageSubtotal.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                    lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
                foreach (var merge in region.PageSubtotal.Merges ?? Array.Empty<ExcelEntityCellRange>())
                    lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            }
        }
        if (region.Footer != null)
        {
            var footerBaseRow = region.GroupSubtotal != null || region.PageSubtotal != null
                ? lastRow + 1 : firstRow + dataRows;
            var footerRow = footerBaseRow + region.Footer.GapRows;
            lastRow = Math.Max(lastRow, footerRow);
            lastColumn = Math.Max(lastColumn, firstColumn);
            foreach (var cell in region.Footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
            {
                lastRow = Math.Max(lastRow, footerRow + cell.Reference.Row);
                lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
            }
            foreach (var merge in region.Footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
            {
                lastRow = Math.Max(lastRow, footerRow + merge.Last.Row);
                lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            }
        }
        if (!hasData && region.Footer == null)
            return;
        var current = new EntityListFootprint(region.SheetName, firstRow, lastRow, firstColumn, lastColumn);
        if (footprints.Any(existing => current.Intersects(existing)))
            throw new BingOfficesConfigurationException(
                $"实体列表区域运行时范围与已有区域重叠: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        footprints.Add(current);
    }

    /// <summary>
    /// 将列表尾部标记、聚合值和相对合并区域写入 NPOI 工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="footer">列表尾部定义。</param>
    /// <param name="items">已物化的列表项快照。</param>
    /// <param name="dataEndRow">明细结束后的零基行索引。</param>
    /// <param name="startColumn">列表区域起始零基列索引。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="anchorName">可选的最终尾部命名锚点。</param>
    /// <param name="detailRows">最终尾部使用的一基明细行段。</param>
    private static void WriteFooter(IWorkbook workbook, ISheet sheet, IExcelEntityListFooter footer,
        IReadOnlyList<object> items, int dataEndRow, int startColumn, CancellationToken cancellationToken,
        string anchorName = null, IReadOnlyList<(int FirstRow, int LastRow)> detailRows = null)
    {
        if (footer == null)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        var footerRow = dataEndRow + footer.GapRows;
        var marker = sheet.GetRow(footerRow) ?? sheet.CreateRow(footerRow);
        var markerCell = marker.GetCell(startColumn) ?? marker.CreateCell(startColumn);
        markerCell.SetCellValue(footer.MarkerText, footer.MarkerStyle?.NumberFormat);
        if (footer.MarkerStyle != null)
            markerCell.CellStyle = NpoiStyleCache.Compose(workbook, markerCell.CellStyle, footer.MarkerStyle);
        foreach (var definition in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(footerRow + definition.Reference.Row)
                ?? sheet.CreateRow(footerRow + definition.Reference.Row);
            var cell = row.GetCell(startColumn + definition.Reference.Column)
                ?? row.CreateCell(startColumn + definition.Reference.Column);
            var value = definition.Evaluate(items);
            var formulaA1 = (value as ExcelEntityFooterFormulaValue)?.FormulaA1;
            if (value is ExcelEntityFooterDetailSumValue detailSum)
            {
                var rows = detailRows ?? (items.Count == 0
                    ? Array.Empty<(int FirstRow, int LastRow)>()
                    : new[] { (dataEndRow - items.Count + 1, dataEndRow) });
                formulaA1 = detailSum.ToFormulaA1(rows);
            }
            if (formulaA1 != null)
            {
                cell.SetCellFormula(formulaA1);
                if (definition.Style != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, definition.Style);
                if (!string.IsNullOrWhiteSpace(definition.NumberFormat))
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle,
                        new ExcelCellStyle { NumberFormat = definition.NumberFormat });
            }
            else
            {
                cell.SetCellValue(value, definition.NumberFormat ?? definition.Style?.NumberFormat);
                if (definition.Style != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, definition.Style);
            }
        }
        var merges = (footer.Merges ?? Array.Empty<ExcelEntityCellRange>()).Select(merge =>
            new CellRangeAddress(footerRow + merge.First.Row, footerRow + merge.Last.Row,
                startColumn + merge.First.Column, startColumn + merge.Last.Column)).ToArray();
        NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, merges);
        if (anchorName != null)
            UpdateFooterAnchor(workbook, sheet, anchorName, footerRow, startColumn);
    }

    /// <summary>
    /// 将工作簿名称定位到最终尾部标记单元格。
    /// </summary>
    /// <param name="workbook">待导出的工作簿。</param>
    /// <param name="sheet">尾部所在工作表。</param>
    /// <param name="anchorName">命名锚点名称。</param>
    /// <param name="row">标记的零基行索引。</param>
    /// <param name="column">标记的零基列索引。</param>
    private static void UpdateFooterAnchor(IWorkbook workbook, ISheet sheet, string anchorName,
        int row, int column)
    {
        var matches = workbook.GetAllNames().Where(name =>
            string.Equals(name.NameName, anchorName, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length > 1)
            throw new BingOfficesConfigurationException(
                $"最终尾部命名锚点不唯一: {sheet.SheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (matches.Length == 1)
            NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, sheet.SheetName, anchorName);
        var name = matches.Length == 1 ? matches[0] : workbook.CreateName();
        if (matches.Length == 0)
            name.NameName = anchorName;
        var escapedSheet = sheet.SheetName.Replace("'", "''");
        name.RefersToFormula = $"'{escapedSheet}'!{new CellReference(row, column, true, true).FormatAsString()}";
    }

    /// <summary>
    /// 获取列表项上声明动态列的字典属性。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>动态字典属性；没有动态列时为 <see langword="null" />。</returns>
    private static PropertyInfo GetDynamicProperty<TItem>(IExcelMappingPlan mapping)
        where TItem : class, new()
    {
        var properties = (mapping?.Columns ?? Array.Empty<IExcelMappingColumn>())
            .Where(column => column.IsDynamicColumn).ToArray();
        if (properties.Length == 0)
            return null;
        if (properties.Length > 1)
            throw new BingOfficesConfigurationException(
                $"实体列表项 {typeof(TItem).FullName} 只能声明一个动态列字典属性。",
                stage: BingOfficesStage.Plan);
        var property = typeof(TItem).GetProperty(properties[0].Name,
            BindingFlags.Instance | BindingFlags.Public);
        if (property == null || !typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
            throw new BingOfficesConfigurationException(
                $"动态列属性必须实现 IDictionary<string, object>: {properties[0].Name}",
                stage: BingOfficesStage.Plan);
        return property;
    }

    /// <summary>
    /// 将映射计划中的动态列转换为实体布局导出定义。
    /// </summary>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>按映射顺序排列的动态列定义。</returns>
    private static IReadOnlyList<ExcelDynamicColumnDefinition> CreateDynamicDefinitions(
        IExcelMappingPlan mapping) => (mapping?.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>())
        .Select(column => NpoiExportColumnPlanner.CreateDynamicDefinition(column, null, mapping))
        .ToArray();

    /// <summary>
    /// 创建列表项动态字典的表达式读取器。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="property">动态字典属性。</param>
    /// <returns>可供标准导出请求使用的动态值表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicGetter<TItem>(
        PropertyInfo property) where TItem : class, new()
    {
        if (property == null)
            throw new ArgumentNullException(nameof(property));
        var parameter = Expression.Parameter(typeof(TItem), "item");
        Expression body = Expression.Property(parameter, property);
        if (body.Type != typeof(IDictionary<string, object>))
            body = Expression.Convert(body, typeof(IDictionary<string, object>));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 创建多个实体动态列组的合并字典读取器。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>将各组动态值合并到同一字典的表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicGetter<TItem>(
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        if (groups == null || groups.Count == 0)
            throw new ArgumentException("动态列组不能为空。", nameof(groups));
        var parameter = Expression.Parameter(typeof(TItem), "item");
        var method = typeof(NpoiEntityExportExecutor).GetMethod(nameof(MergeDynamicValues),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(typeof(TItem));
        var body = Expression.Call(method, parameter, Expression.Constant(groups));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 合并一个列表项的多个动态值字典。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">列表项。</param>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>按动态列键合并的字典。</returns>
    private static IDictionary<string, object> MergeDynamicValues<TItem>(TItem item,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        var values = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            var source = group.Getter(item);
            if (source == null)
                continue;
            foreach (var pair in source)
                values[pair.Key] = pair.Value;
        }
        return values;
    }

    /// <summary>
    /// 调用反射方法，并保留内部异常的原始堆栈后重新抛出。
    /// </summary>
    /// <param name="method">待调用的方法。</param>
    /// <param name="target">方法目标实例。</param>
    /// <param name="arguments">方法参数。</param>
    private static void InvokeUnwrapped(MethodInfo method, object target, object[] arguments)
    {
        try
        {
            method.Invoke(target, arguments);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }
}
