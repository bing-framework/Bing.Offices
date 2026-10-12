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

namespace Bing.Offices.Npoi.Entities;

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
            if (resolvedRegion.PageHeader != null && sheet.RepeatingRows != null)
                throw new BingOfficesConfigurationException(
                    $"工作表打印标题与实际重复页眉互斥: {sheet.SheetName}",
                    stage: BingOfficesStage.Plan);
            NpoiEntityLayoutSupport.ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            var values = resolvedRegion.Getter(entity)?.Cast<object>().ToArray() ?? Array.Empty<object>();
            if (resolvedRegion.PageHeader != null
                && values.Length > resolvedRegion.PageBreakRows.GetValueOrDefault()
                && layout.DataValidations.TryGetValue(sheet.SheetName, out var validations))
                ValidateConfiguredPageHeaderValidations(resolvedRegion, validations);
            WriteRegion(workbook, sheet, entity, resolvedRegion, values, layout.Cells, footprints,
                cancellationToken);
        }
        foreach (var pair in layout.PrintTitleRows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = NpoiEntityLayoutSupport.ResolveSheet(workbook, pair.Key, template);
            var last = int.Parse(pair.Value.Split(':')[1], CultureInfo.InvariantCulture);
            if (last > workbook.SpreadsheetVersion.LastRowIndex + 1)
                throw new BingOfficesConfigurationException("重复打印标题行超出当前格式行数。",
                    stage: BingOfficesStage.Plan);
            sheet.RepeatingRows = CellRangeAddress.ValueOf(pair.Value);
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
    /// <param name="fixedCells">布局中的固定单元格。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values,
        IReadOnlyList<ExcelEntityCellBinding<TEntity>> fixedCells,
        ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(WriteTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, region, values, fixedCells,
            footprints, cancellationToken });
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
    /// <param name="fixedCells">布局中的固定单元格。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void WriteTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<object> values,
        IReadOnlyList<ExcelEntityCellBinding<TEntity>> fixedCells,
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
        var dynamicProperties = groups.Count == 0
            ? NpoiEntityLayoutSupport.GetDynamicProperties<TItem>(mapping)
            : Array.Empty<PropertyInfo>();
        var dynamicProperty = dynamicProperties.Count == 1 ? dynamicProperties[0] : null;
        var implicitDynamicGetter = dynamicProperties.Count > 1
            ? CreateDynamicGetter<TItem>(dynamicProperties)
            : dynamicProperty == null ? null : CreateDynamicGetter<TItem>(dynamicProperty);
        var dynamicGetter = groups.Count > 0
            ? CreateDynamicGetter<TItem>(groups)
            : implicitDynamicGetter;
        var dynamicValuesGetter = dynamicGetter?.Compile();
        var dynamicDefinitions = groupDefinitions.Length > 0
            ? groupDefinitions : CreateDynamicDefinitions(mapping);
        var columns = groups.Count == 0
            ? NpoiExportColumnPlanner.CreateColumns<TItem>(mapping, dynamicDefinitions,
                region.CalculatedColumns)
            : NpoiExportColumnPlanner.CreateEntityColumns<TItem>(mapping, dynamicDefinitions, groups,
                region.CalculatedColumns);
        if (region.GroupSubtotal != null || region.PageHeader != null || region.PageFooter != null)
        {
            WriteGroupedTypedRegion(workbook, sheet, region, data, dynamicValuesGetter,
                dynamicDefinitions, columns, fixedCells, footprints, cancellationToken);
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
            WritePagedTypedRegion(workbook, sheet, region, data, dynamicValuesGetter,
                dynamicDefinitions, columns, cancellationToken);
            return;
        }
        if (region.IncludeHeader)
        {
            if (dynamicDefinitions.Count > 0 && groups.Count == 0)
            {
                request = ExcelExport.Workbook(builder => builder.AddSheet(region.SheetName, data, configure =>
                {
                    configure.HeaderRowIndex(0).DataRowStartIndex(1);
                    if (region.MappingDocument != null)
                        configure.Mapping(region.MappingDocument);
                    else if (region.MappingConfiguration != null)
                        configure.Mapping(region.MappingConfiguration);
                    configure.DynamicColumns(implicitDynamicGetter, dynamicDefinitions)
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
            var dynamicValues = dynamicValuesGetter?.Invoke(item);
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
        Func<TItem, IDictionary<string, object>> dynamicValuesGetter,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicDefinitions,
        IReadOnlyList<ExcelColumnPlan> columns, CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        if (region.PageSubtotal == null || region.PageBreakRows is not int rowsPerPage)
            throw new BingOfficesConfigurationException("分页小计缺少分页配置。",
                stage: BingOfficesStage.Plan);
        var pageAnchorNames = PreparePageAnchorNames(workbook, sheet, region.PageSubtotalAnchorPrefix,
            data.Count < 2 ? 0 : (data.Count - 1) / rowsPerPage);
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
        var pageAnchorIndex = 0;
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
                var dynamicValues = dynamicValuesGetter?.Invoke(item);
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
                    region.Start.Column, cancellationToken, pageAnchorNames.Count == 0
                        ? null : pageAnchorNames[pageAnchorIndex++]);
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
    /// <param name="dynamicValuesGetter">读取当前列表项动态值的委托。</param>
    /// <param name="dynamicDefinitions">动态列定义。</param>
    /// <param name="columns">最终列计划。</param>
    /// <param name="fixedCells">布局中的固定单元格。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void WriteGroupedTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<TItem> data,
        Func<TItem, IDictionary<string, object>> dynamicValuesGetter,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicDefinitions,
        IReadOnlyList<ExcelColumnPlan> columns,
        IReadOnlyList<ExcelEntityCellBinding<TEntity>> fixedCells,
        ICollection<EntityListFootprint> footprints,
        CancellationToken cancellationToken)
        where TEntity : class, new() where TItem : class, new()
    {
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var rowIndex = firstDataRow;
        var calculatedItems = columns.Any(column => column.IsCalculated)
            ? data.Cast<object>().ToArray() : null;
        var calculatedEvaluators = columns.Where(column => column.IsCalculated)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, sheet.SheetName),
                StringComparer.OrdinalIgnoreCase);
        var rowPlan = ExcelEntityRowPlan.Create(region, data.Cast<object>().ToArray(), cancellationToken);
        var pageAnchorNames = PreparePageAnchorNames(workbook, sheet, region.PageSubtotalAnchorPrefix,
            rowPlan.Segments.Count(segment => segment.Kind == ExcelEntityRowPlan.SegmentKind.PageSubtotal));
        var pageCount = data.Count == 0 ? 1 : 1 + (data.Count - 1) / region.PageBreakRows.GetValueOrDefault(int.MaxValue);
        var headerAnchors = PrepareIndependentPageAnchorNames(workbook, sheet,
            region.PageHeaderAnchorPrefix, pageCount, region.PageHeaderAnchorStaleNamePolicy,
            cancellationToken);
        var footerAnchors = PrepareIndependentPageAnchorNames(workbook, sheet,
            region.PageFooterAnchorPrefix, pageCount, region.PageFooterAnchorStaleNamePolicy,
            cancellationToken);
        NpoiEntityLayoutSupport.ValidateListRegionBounds(region, columns.Count, data.Count,
            region.IncludeHeader);
        foreach (var segment in rowPlan.Segments)
        {
            if (segment.Kind == ExcelEntityRowPlan.SegmentKind.Detail)
            {
                if ((long)segment.Row + segment.Count - 1 > workbook.SpreadsheetVersion.LastRowIndex
                    || region.End.HasValue && (long)segment.Row + segment.Count - 1 > region.End.Value.Row)
                    throw new BingOfficesConfigurationException(
                        $"列表区域明细超出工作表边界: {region.SheetName}!{region.Start.Address}",
                        stage: BingOfficesStage.Plan);
                continue;
            }
            if (segment.Kind == ExcelEntityRowPlan.SegmentKind.PageHeader)
                continue;
            var footer = segment.Kind == ExcelEntityRowPlan.SegmentKind.GroupSubtotal
                ? region.GroupSubtotal.Footer : segment.Kind == ExcelEntityRowPlan.SegmentKind.PageFooter
                    ? region.PageFooter : region.PageSubtotal;
            NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, footer, segment.Row);
            NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, footer, segment.Row,
                region.Start.Column);
            if (segment.Kind == ExcelEntityRowPlan.SegmentKind.PageFooter)
            {
                var lastColumn = region.Start.Column + footer.Cells.Select(cell => cell.Reference.Column)
                    .Concat(footer.Merges.Select(merge => merge.Last.Column)).DefaultIfEmpty(0).Max();
                var lastFooterRow = (long)segment.Row + footer.GapRows
                    + NpoiEntityLayoutSupport.GetFooterHeight(footer) - 1;
                foreach (var binding in fixedCells)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var target = binding.Reference;
                    var targetSheet = binding.SheetName;
                    if (binding.AnchorName != null)
                    {
                        var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook,
                            binding.SheetName, binding.AnchorName);
                        target = named.Reference;
                        targetSheet = named.Sheet.SheetName;
                    }
                    if (string.Equals(targetSheet, sheet.SheetName, StringComparison.OrdinalIgnoreCase)
                        && target.Row >= segment.Row && target.Row <= lastFooterRow
                        && target.Column >= region.Start.Column && target.Column <= lastColumn)
                        throw new BingOfficesConfigurationException("每页尾部与固定单元格冲突。",
                            BingOfficesStage.Plan, region.SheetName, target.Row + 1, target.Column + 1);
                }
            }
        }
        NpoiEntityLayoutSupport.ValidateFooterBounds(workbook, region, region.Footer, rowPlan.FinalRow);
        NpoiEntityLayoutSupport.PreflightFooterMerges(sheet, region.Footer, rowPlan.FinalRow,
            region.Start.Column);
        ValidatePageHeader(workbook, sheet, region, rowPlan, columns.Count, fixedCells);
        var lastRow = (long)rowPlan.FinalRow + (region.Footer == null ? -1
            : region.Footer.GapRows + NpoiEntityLayoutSupport.GetFooterHeight(region.Footer) - 1);
        if (region.End.HasValue && (lastRow > region.End.Value.Row
            || region.Start.Column + columns.Count - 1 > region.End.Value.Column))
            throw new BingOfficesConfigurationException(
                $"列表区域分组小计超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        RegisterRuntimeFootprint(region, columns.Count, data.Count, footprints, rowPlan);
        cancellationToken.ThrowIfCancellationRequested();
        foreach (var obsolete in headerAnchors.Obsolete.Concat(footerAnchors.Obsolete))
        {
            cancellationToken.ThrowIfCancellationRequested();
            workbook.RemoveName(obsolete);
        }

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
        if (region.PageHeader != null)
        {
            WritePageHeader(workbook, sheet, region.PageHeader,
                region.PageHeader.SourceRange.First.Row, 1,
                pageCount,
                data.Take(region.PageBreakRows.Value).Cast<object>().ToArray(), copySource: false,
                cancellationToken);
            ValidateWrittenPageHeader(sheet, region, columns,
                region.PageHeader.SourceRange.First.Row);
            if (headerAnchors.Names.Count > 0)
                UpdateFooterAnchor(workbook, sheet, headerAnchors.Names[0],
                    region.PageHeader.Marker.Row, region.PageHeader.Marker.Column);
        }

        void WriteItem(TItem item, int itemIndex)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = dynamicValuesGetter?.Invoke(item);
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
            ValidateWrittenDetailMarker(sheet, region, rowIndex);
            rowIndex++;
        }

        var pageAnchorIndex = 0;
        var pageNumber = 1;
        foreach (var segment in rowPlan.Segments)
        {
            if (segment.Kind == ExcelEntityRowPlan.SegmentKind.Detail)
            {
                for (var index = 0; index < segment.Count; index++)
                    WriteItem(data[segment.StartIndex + index], segment.StartIndex + index);
                continue;
            }
            if (segment.Kind == ExcelEntityRowPlan.SegmentKind.PageHeader)
            {
                pageNumber++;
                WritePageHeader(workbook, sheet, region.PageHeader, rowIndex, pageNumber,
                    pageCount,
                    data.Skip(segment.StartIndex).Take(region.PageBreakRows.Value)
                        .Cast<object>().ToArray(), copySource: true, cancellationToken);
                ValidateWrittenPageHeader(sheet, region, columns, rowIndex);
                if (headerAnchors.Names.Count > 0)
                    UpdateFooterAnchor(workbook, sheet, headerAnchors.Names[pageNumber - 1],
                        rowIndex + region.PageHeader.Marker.Row - region.PageHeader.SourceRange.First.Row,
                        region.PageHeader.Marker.Column);
                rowIndex += region.PageHeader.RowCount;
                continue;
            }
            var footer = segment.Kind == ExcelEntityRowPlan.SegmentKind.GroupSubtotal
                ? region.GroupSubtotal.Footer : segment.Kind == ExcelEntityRowPlan.SegmentKind.PageFooter
                    ? region.PageFooter : region.PageSubtotal;
            var items = data.Skip(segment.StartIndex).Take(segment.Count).Cast<object>().ToArray();
            WriteFooter(workbook, sheet, footer, items, rowIndex, region.Start.Column,
                cancellationToken, segment.Kind == ExcelEntityRowPlan.SegmentKind.PageSubtotal
                    && pageAnchorNames.Count > 0 ? pageAnchorNames[pageAnchorIndex++]
                    : segment.Kind == ExcelEntityRowPlan.SegmentKind.PageFooter && footerAnchors.Names.Count > 0
                        ? footerAnchors.Names[segment.StartIndex / region.PageBreakRows.Value] : null);
            rowIndex += footer.GapRows + NpoiEntityLayoutSupport.GetFooterHeight(footer);
        }
        WriteFooter(workbook, sheet, region.Footer, data.Cast<object>().ToArray(), rowIndex,
            region.Start.Column, cancellationToken, region.FooterAnchorName, rowPlan.DetailRows);
        ExcelEntitySubtotalImportScanner.ValidateWrittenMarkers(region, firstDataRow, rowPlan.FinalRow,
            physicalRow => NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(
                sheet.GetRow(physicalRow)?.GetCell(region.Start.Column)), ExcelWhitespacePolicy.Trim),
            physicalRow => region.PageHeader == null ? null : NpoiExcelImporter.NormalizeText(
                NpoiExcelImporter.GetRawStringValue(sheet.GetRow(physicalRow)?.GetCell(
                    region.PageHeader.Marker.Column)), ExcelWhitespacePolicy.Trim), cancellationToken);
        foreach (var breakRow in rowPlan.PageBreakRows)
            NpoiEntityLayoutSupport.ApplyPageBreak(sheet, breakRow);
    }

    /// <summary>预检实际重复页眉的源区域和目标行。</summary>
    private static void ValidatePageHeader<TEntity>(IWorkbook workbook, ISheet sheet,
        ExcelEntityListRegion<TEntity> region, ExcelEntityRowPlan rowPlan, int columnCount,
        IReadOnlyList<ExcelEntityCellBinding<TEntity>> fixedCells)
        where TEntity : class, new()
    {
        var header = region.PageHeader;
        if (header == null)
            return;
        ExcelEntityPageHeaderImportValidation.ValidateMarkerColumn(region, columnCount);
        if (rowPlan.Segments.Any(item => item.Kind == ExcelEntityRowPlan.SegmentKind.PageHeader))
            ValidatePageHeaderSourceMetadata(sheet, header);
        var source = header.SourceRange;
        var firstDetail = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        if (source.Last.Row != firstDetail - 1 || source.Last.Row > workbook.SpreadsheetVersion.LastRowIndex
            || source.Last.Column > workbook.SpreadsheetVersion.LastColumnIndex)
            throw new BingOfficesConfigurationException("重复页眉源区域未紧邻明细或超出格式边界。",
                stage: BingOfficesStage.Plan);
        if (region.IncludeHeader && (source.First.Column > region.Start.Column
            || source.Last.Column < region.Start.Column + columnCount - 1))
            throw new BingOfficesConfigurationException("重复页眉源区域未包含完整列表表头。",
                stage: BingOfficesStage.Plan);
        if (region.IncludeHeader && header.Marker.Row == region.Start.Row
            && header.Marker.Column >= region.Start.Column
            && header.Marker.Column < region.Start.Column + columnCount)
            throw new BingOfficesConfigurationException("重复页眉标记与列表表头重叠。",
                stage: BingOfficesStage.Plan);
        foreach (var merged in sheet.MergedRegions)
        {
            var overlaps = merged.FirstRow <= source.Last.Row && merged.LastRow >= source.First.Row
                && merged.FirstColumn <= source.Last.Column && merged.LastColumn >= source.First.Column;
            if (overlaps && (merged.FirstRow < source.First.Row || merged.LastRow > source.Last.Row
                || merged.FirstColumn < source.First.Column || merged.LastColumn > source.Last.Column))
                throw new BingOfficesConfigurationException("重复页眉源区域包含跨界合并。",
                    stage: BingOfficesStage.Plan);
            if (merged.IsInRange(header.Marker.Row, header.Marker.Column)
                && (merged.FirstRow != header.Marker.Row
                    || merged.FirstColumn != header.Marker.Column))
                throw new BingOfficesConfigurationException("重复页眉标记必须位于合并区域左上角。",
                    stage: BingOfficesStage.Plan);
        }
        foreach (var merge in header.Merges)
        {
            var absolute = new CellRangeAddress(source.First.Row + merge.First.Row,
                source.First.Row + merge.Last.Row,
                source.First.Column + merge.First.Column,
                source.First.Column + merge.Last.Column);
            if (region.IncludeHeader && absolute.FirstRow <= region.Start.Row
                && absolute.LastRow >= region.Start.Row
                && absolute.FirstColumn <= region.Start.Column + columnCount - 1
                && absolute.LastColumn >= region.Start.Column)
                throw new BingOfficesConfigurationException("重复页眉合并区域与列表表头重叠。",
                    BingOfficesStage.Plan, region.SheetName, region.Start.Row + 1,
                    Math.Max(absolute.FirstColumn, region.Start.Column) + 1);
            if (sheet.MergedRegions.Any(existing => existing.Intersects(absolute)
                && !existing.Equals(absolute)))
                throw new BingOfficesConfigurationException("重复页眉合并区域与模板合并冲突。",
                    stage: BingOfficesStage.Plan);
        }
        foreach (var segment in rowPlan.Segments.Where(item =>
                     item.Kind == ExcelEntityRowPlan.SegmentKind.PageHeader))
        {
            if ((long)segment.Row + header.RowCount - 1 > workbook.SpreadsheetVersion.LastRowIndex
                || region.End.HasValue && (long)segment.Row + header.RowCount - 1 > region.End.Value.Row)
                throw new BingOfficesConfigurationException("重复页眉超出列表或格式行数。",
                    stage: BingOfficesStage.Plan);
            foreach (var binding in fixedCells)
            {
                var target = binding.Reference;
                var targetSheet = binding.SheetName;
                if (binding.AnchorName != null)
                {
                    var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook,
                        binding.SheetName, binding.AnchorName);
                    target = named.Reference;
                    targetSheet = named.Sheet.SheetName;
                }
                if (string.Equals(targetSheet, sheet.SheetName, StringComparison.OrdinalIgnoreCase)
                    && target.Row >= segment.Row && target.Row < segment.Row + header.RowCount
                    && target.Column >= source.First.Column && target.Column <= source.Last.Column)
                    throw new BingOfficesConfigurationException("重复页眉目标区域与固定单元格冲突。",
                        stage: BingOfficesStage.Plan);
            }
            for (var row = segment.Row; row < segment.Row + header.RowCount; row++)
            {
                var existing = sheet.GetRow(row);
                if (existing == null)
                    continue;
                for (var column = source.First.Column; column <= source.Last.Column; column++)
                {
                    var cell = existing.GetCell(column);
                    if (cell != null && cell.CellType != CellType.Blank)
                        throw new BingOfficesConfigurationException("重复页眉目标区域已有内容。",
                            stage: BingOfficesStage.Plan);
                }
            }
            if (sheet.MergedRegions.Any(merge => merge.FirstRow <= segment.Row + header.RowCount - 1
                && merge.LastRow >= segment.Row && merge.FirstColumn <= source.Last.Column
                && merge.LastColumn >= source.First.Column))
                throw new BingOfficesConfigurationException("重复页眉目标区域与模板合并冲突。",
                    stage: BingOfficesStage.Plan);
        }
    }

    /// <summary>拒绝源页眉中无法随页复制的单元格和区域元数据。</summary>
    private static void ValidatePageHeaderSourceMetadata(ISheet sheet,
        ExcelEntityPageHeaderDefinition header)
    {
        var source = header.SourceRange;
        var sourceRange = new CellRangeAddress(source.First.Row, source.Last.Row,
            source.First.Column, source.Last.Column);
        for (var row = source.First.Row; row <= source.Last.Row; row++)
        {
            var record = sheet.GetRow(row);
            for (var column = source.First.Column; column <= source.Last.Column; column++)
            {
                var cell = record?.GetCell(column);
                if (cell == null)
                    continue;
                if (cell.Hyperlink == null && cell.CellComment == null
                    && (cell.CellType != CellType.String
                        || cell.RichStringCellValue.NumFormattingRuns == 0))
                    continue;
                throw new BingOfficesConfigurationException("重复页眉源单元格包含不支持复制的元数据。",
                    BingOfficesStage.Plan, sheet.SheetName, row + 1, column + 1);
            }
        }
        foreach (var validation in sheet.GetDataValidations())
            foreach (var range in validation.Regions.CellRangeAddresses)
                if (range.Intersects(sourceRange))
                    throw new BingOfficesConfigurationException("重复页眉源区域包含数据校验。",
                        BingOfficesStage.Plan, sheet.SheetName,
                        Math.Max(range.FirstRow, source.First.Row) + 1,
                        Math.Max(range.FirstColumn, source.First.Column) + 1);
        var formats = sheet.SheetConditionalFormatting;
        for (var index = 0; index < formats.NumConditionalFormattings; index++)
            foreach (var range in formats.GetConditionalFormattingAt(index).GetFormattingRanges())
                if (range.Intersects(sourceRange))
                    throw new BingOfficesConfigurationException("重复页眉源区域包含条件格式。",
                        BingOfficesStage.Plan, sheet.SheetName,
                        Math.Max(range.FirstRow, source.First.Row) + 1,
                        Math.Max(range.FirstColumn, source.First.Column) + 1);
    }

    /// <summary>拒绝覆盖页眉源区域的配置式数据校验。</summary>
    private static void ValidateConfiguredPageHeaderValidations<TEntity>(
        ExcelEntityListRegion<TEntity> region,
        IReadOnlyList<ExcelDataValidationDefinition> validations)
        where TEntity : class, new()
    {
        var source = region.PageHeader.SourceRange;
        foreach (var validation in validations)
        {
            var range = validation.Range;
            if (range.EndRow < source.First.Row || range.StartRow > source.Last.Row
                || range.EndColumn < source.First.Column || range.StartColumn > source.Last.Column)
                continue;
            throw new BingOfficesConfigurationException("重复页眉源区域包含配置式数据校验。",
                BingOfficesStage.Plan, region.SheetName,
                Math.Max(range.StartRow, source.First.Row) + 1,
                Math.Max(range.StartColumn, source.First.Column) + 1);
        }
    }

    /// <summary>校验明细中的保留标记。</summary>
    private static void ValidateWrittenDetailMarker<TEntity>(ISheet sheet,
        ExcelEntityListRegion<TEntity> region, int rowIndex) where TEntity : class, new()
    {
        if (region.PageFooter != null && string.Equals(NpoiExcelImporter.NormalizeText(
                NpoiExcelImporter.GetRawStringValue(sheet.GetRow(rowIndex)?.GetCell(region.Start.Column)),
                ExcelWhitespacePolicy.Trim), region.PageFooter.MarkerText, StringComparison.Ordinal))
            throw new BingOfficesConfigurationException("明细值与每页尾部标记文本冲突。",
                BingOfficesStage.Plan, region.SheetName, rowIndex + 1, region.Start.Column + 1);
        var header = region.PageHeader;
        if (header == null)
            return;
        var value = NpoiExcelImporter.NormalizeText(
            NpoiExcelImporter.GetRawStringValue(sheet.GetRow(rowIndex)?.GetCell(header.Marker.Column)),
            ExcelWhitespacePolicy.Trim);
        if (string.Equals(value, header.MarkerText, StringComparison.Ordinal))
            throw new BingOfficesConfigurationException("明细值与重复页眉标记文本冲突。",
                BingOfficesStage.Plan, region.SheetName, rowIndex + 1,
                header.Marker.Column + 1);
    }

    /// <summary>校验写出的每页列表表头未被覆盖。</summary>
    private static void ValidateWrittenPageHeader<TEntity>(ISheet sheet,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<ExcelColumnPlan> columns,
        int startRow) where TEntity : class, new()
    {
        if (!region.IncludeHeader)
            return;
        var row = startRow + region.Start.Row - region.PageHeader.SourceRange.First.Row;
        for (var index = 0; index < columns.Count; index++)
        {
            var column = region.Start.Column + index;
            var actual = NpoiExcelImporter.NormalizeText(
                NpoiExcelImporter.GetRawStringValue(sheet.GetRow(row)?.GetCell(column)),
                ExcelWhitespacePolicy.Trim);
            if (!string.Equals(actual, columns[index].Title, StringComparison.Ordinal))
                throw new BingOfficesConfigurationException("重复页眉的列表表头与映射标题不一致。",
                    BingOfficesStage.Plan, region.SheetName, row + 1, column + 1);
        }
    }

    /// <summary>复制并覆盖一个逻辑页的页眉。</summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="header">页眉定义。</param>
    /// <param name="targetRow">目标起始行号。</param>
    /// <param name="pageNumber">当前逻辑页码。</param>
    /// <param name="pageCount">逻辑页总数。</param>
    /// <param name="items">当前页明细快照。</param>
    /// <param name="copySource">是否复制首页来源区域。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="afterSourceCellCopied">来源单元格复制后的观察回调。</param>
    internal static void WritePageHeader(IWorkbook workbook, ISheet sheet,
        ExcelEntityPageHeaderDefinition header, int targetRow, int pageNumber, int pageCount,
        IReadOnlyList<object> items, bool copySource, CancellationToken cancellationToken,
        Action<int, int> afterSourceCellCopied = null)
    {
        var source = header.SourceRange;
        if (copySource)
        {
            for (var offset = 0; offset < header.RowCount; offset++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var from = sheet.GetRow(source.First.Row + offset);
                var to = sheet.GetRow(targetRow + offset) ?? sheet.CreateRow(targetRow + offset);
                if (from == null)
                    continue;
                to.Height = from.Height;
                to.ZeroHeight = from.ZeroHeight;
                for (var column = source.First.Column; column <= source.Last.Column; column++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var original = from.GetCell(column);
                    if (original == null)
                        continue;
                    var cell = to.GetCell(column) ?? to.CreateCell(column);
                    cell.CellStyle = original.CellStyle;
                    switch (original.CellType)
                    {
                        case CellType.String: cell.SetCellValue(original.RichStringCellValue); break;
                        case CellType.Numeric: cell.SetCellValue(original.NumericCellValue); break;
                        case CellType.Boolean: cell.SetCellValue(original.BooleanCellValue); break;
                        case CellType.Formula: cell.SetCellFormula(original.CellFormula); break;
                        case CellType.Error: cell.SetCellErrorValue(original.ErrorCellValue); break;
                        default: cell.SetCellType(CellType.Blank); break;
                    }
                    afterSourceCellCopied?.Invoke(source.First.Row + offset, column);
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var merged in sheet.MergedRegions.Where(merge =>
                         merge.FirstRow >= source.First.Row && merge.LastRow <= source.Last.Row
                         && merge.FirstColumn >= source.First.Column
                         && merge.LastColumn <= source.Last.Column).ToArray())
            {
                cancellationToken.ThrowIfCancellationRequested();
                sheet.AddMergedRegion(new CellRangeAddress(targetRow + merged.FirstRow - source.First.Row,
                    targetRow + merged.LastRow - source.First.Row,
                    merged.FirstColumn, merged.LastColumn));
            }
        }
        foreach (var merge in header.Merges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var absolute = new CellRangeAddress(targetRow + merge.First.Row,
                targetRow + merge.Last.Row, source.First.Column + merge.First.Column,
                source.First.Column + merge.Last.Column);
            if (!sheet.MergedRegions.Any(existing => existing.Equals(absolute)))
                sheet.AddMergedRegion(absolute);
        }
        foreach (var definition in header.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = targetRow + definition.Reference.Row;
            var column = source.First.Column + definition.Reference.Column;
            var cell = sheet.GetRow(row) ?? sheet.CreateRow(row);
            var target = cell.GetCell(column) ?? cell.CreateCell(column);
            object value;
            try
            {
                value = definition.GetValue(pageNumber, pageCount, items);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesExportException("重复页眉值工厂执行失败。", exception,
                    "NPOI", BingOfficesStage.Validate, sheet.SheetName, row + 1, column + 1,
                    null, BingOfficesErrorCode.UserExtensionFailed);
            }
            target.SetCellValue(value, definition.NumberFormat ?? definition.Style?.NumberFormat);
            if (definition.Style != null)
                target.CellStyle = NpoiStyleCache.Compose(workbook, target.CellStyle, definition.Style);
            if (!string.IsNullOrWhiteSpace(definition.NumberFormat))
                target.CellStyle = NpoiStyleCache.Compose(workbook, target.CellStyle,
                    new ExcelCellStyle { NumberFormat = definition.NumberFormat });
        }
        cancellationToken.ThrowIfCancellationRequested();
        var markerRow = targetRow + header.Marker.Row - source.First.Row;
        var markerRecord = sheet.GetRow(markerRow) ?? sheet.CreateRow(markerRow);
        var markerCell = markerRecord.GetCell(header.Marker.Column)
            ?? markerRecord.CreateCell(header.Marker.Column);
        markerCell.SetCellValue(header.MarkerText);
    }

    /// <summary>
    /// 预检并登记列表区域的运行时占用范围。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="columnCount">映射计划的列数。</param>
    /// <param name="itemCount">实际明细行数。</param>
    /// <param name="footprints">已登记的列表区域范围。</param>
    /// <param name="rowPlan">已预检的分组分页物理行规划。</param>
    private static void RegisterRuntimeFootprint<TEntity>(ExcelEntityListRegion<TEntity> region,
        int columnCount, int itemCount, ICollection<EntityListFootprint> footprints,
        ExcelEntityRowPlan rowPlan = null)
        where TEntity : class, new()
    {
        var dataRows = (long)itemCount + (region.IncludeHeader ? 1 : 0);
        var hasData = dataRows > 0;
        var firstRow = (long)region.Start.Row;
        var lastRow = hasData ? firstRow + dataRows - 1 : firstRow - 1;
        var firstColumn = (long)region.Start.Column;
        var lastColumn = hasData ? firstColumn + Math.Max(0, columnCount - 1) : firstColumn - 1;
        if (region.GroupSubtotal != null || region.PageHeader != null || region.PageFooter != null)
        {
            lastRow = Math.Max(lastRow, (long)rowPlan.FinalRow - 1);
            if (region.PageHeader != null)
            {
                firstRow = Math.Min(firstRow, region.PageHeader.SourceRange.First.Row);
                lastColumn = Math.Max(lastColumn, region.PageHeader.SourceRange.Last.Column);
            }
            foreach (var cell in region.GroupSubtotal?.Footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
            foreach (var merge in region.GroupSubtotal?.Footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
                lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            foreach (var cell in region.PageFooter?.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
            foreach (var merge in region.PageFooter?.Merges ?? Array.Empty<ExcelEntityCellRange>())
                lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            if (region.PageSubtotal != null)
            {
                foreach (var cell in region.PageSubtotal.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
                    lastColumn = Math.Max(lastColumn, firstColumn + cell.Reference.Column);
                foreach (var merge in region.PageSubtotal.Merges ?? Array.Empty<ExcelEntityCellRange>())
                    lastColumn = Math.Max(lastColumn, firstColumn + merge.Last.Column);
            }
        }
        if (region.GroupSubtotal == null && region.PageSubtotal != null
            && region.PageBreakRows is int rowsPerPage)
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
                || region.PageHeader != null || region.PageFooter != null
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
        if (!hasData && region.Footer == null && region.PageHeader == null && region.PageFooter == null)
            return;
        var footprintFirstColumn = region.PageHeader == null ? firstColumn
            : Math.Min(firstColumn, region.PageHeader.SourceRange.First.Column);
        var current = new EntityListFootprint(region.SheetName, firstRow, lastRow,
            footprintFirstColumn, lastColumn);
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
    /// 将工作簿名称定位到尾部标记单元格。
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
                $"尾部命名锚点不唯一: {sheet.SheetName}!{anchorName}",
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
    /// 预检分页小计的顺序名称和模板占用情况。
    /// </summary>
    /// <param name="workbook">待导出的工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="prefix">分页小计名称前缀。</param>
    /// <param name="count">本次中间页小计数。</param>
    /// <returns>按分页顺序排列的名称。</returns>
    private static IReadOnlyList<string> PreparePageAnchorNames(IWorkbook workbook, ISheet sheet,
        string prefix, int count)
    {
        if (prefix == null)
            return Array.Empty<string>();
        var names = Enumerable.Range(1, count).Select(index =>
            ExcelEntityPageAnchorNames.Create(prefix, index)).ToArray();
        var expected = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
        var existing = workbook.GetAllNames().Where(name =>
            ExcelEntityPageAnchorNames.IsSeriesName(prefix, name.NameName)).ToArray();
        if (existing.Any(name => !expected.Contains(name.NameName)))
            throw new BingOfficesConfigurationException(
                $"模板包含本次未使用的分页小计命名锚点: {sheet.SheetName}!{prefix}",
                stage: BingOfficesStage.Plan);
        var sheetIndex = workbook.GetSheetIndex(sheet);
        foreach (var name in names)
        {
            var matches = existing.Where(item => string.Equals(item.NameName, name,
                StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length > 1 || matches.Any(item => item.SheetIndex >= 0
                && item.SheetIndex != sheetIndex))
                throw new BingOfficesConfigurationException(
                    $"分页小计命名锚点不唯一或属于其他 Sheet: {sheet.SheetName}!{name}",
                    stage: BingOfficesStage.Plan);
            if (matches.Length == 1)
                NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, sheet.SheetName, name);
        }
        return names;
    }

    /// <summary>
    /// 预检独立逐页名称及允许清理的旧名称。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="prefix">逐页名称前缀。</param>
    /// <param name="count">逻辑页数。</param>
    /// <param name="staleNames">旧名称处理策略。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>预期名称及通过预检的过期名称。</returns>
    private static (IReadOnlyList<string> Names, IReadOnlyList<IName> Obsolete)
        PrepareIndependentPageAnchorNames(IWorkbook workbook, ISheet sheet, string prefix, int count,
            ExcelEntityPageAnchorStaleNamePolicy staleNames, CancellationToken cancellationToken)
    {
        if (prefix == null)
            return (Array.Empty<string>(), Array.Empty<IName>());
        var names = Enumerable.Range(1, count).Select(index =>
            ExcelEntityPageAnchorNames.Create(prefix, index)).ToArray();
        var expected = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
        var existing = workbook.GetAllNames().Where(name =>
            ExcelEntityPageAnchorNames.IsSeriesName(prefix, name.NameName)).ToArray();
        var obsolete = new List<IName>();
        var sheetIndex = workbook.GetSheetIndex(sheet);
        foreach (var group in existing.GroupBy(name => name.NameName, StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var matches = group.ToArray();
            if (matches.Length != 1 || matches[0].SheetIndex >= 0 && matches[0].SheetIndex != sheetIndex)
                throw new BingOfficesConfigurationException(
                    $"逐页命名锚点不唯一或属于其他 Sheet: {sheet.SheetName}!{group.Key}",
                    stage: BingOfficesStage.Plan);
            NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, sheet.SheetName, group.Key);
            if (expected.Contains(group.Key))
                continue;
            if (staleNames != ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete)
                throw new BingOfficesConfigurationException(
                    $"模板包含本次未使用的逐页命名锚点: {sheet.SheetName}!{group.Key}",
                    stage: BingOfficesStage.Plan);
            obsolete.Add(matches[0]);
        }
        return (names, obsolete);
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
    /// 创建多个隐式动态字典的合并读取器。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="properties">动态字典属性。</param>
    /// <returns>将多个字典合并到同一动态值字典的表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicGetter<TItem>(
        IReadOnlyList<PropertyInfo> properties) where TItem : class, new()
    {
        if (properties == null || properties.Count == 0)
            throw new ArgumentException("动态字典属性不能为空。", nameof(properties));
        var parameter = Expression.Parameter(typeof(TItem), "item");
        var method = typeof(NpoiEntityExportExecutor).GetMethod(nameof(MergeImplicitDynamicValues),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(typeof(TItem));
        var body = Expression.Call(method, parameter, Expression.Constant(properties));
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
    /// 合并一个列表项的多个隐式动态字典。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">列表项。</param>
    /// <param name="properties">动态字典属性。</param>
    /// <returns>按动态列键合并的字典。</returns>
    private static IDictionary<string, object> MergeImplicitDynamicValues<TItem>(TItem item,
        IReadOnlyList<PropertyInfo> properties) where TItem : class, new() =>
        NpoiEntityLayoutSupport.MergeImplicitDynamicValues(item, properties);

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
