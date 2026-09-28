using System.Collections;
using System.Globalization;
using System.Reflection;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Entities;

/// <summary>
/// 执行 ClosedXML 的固定单元格、列表区域和合并布局。
/// </summary>
internal sealed class ClosedXmlEntityLayoutExecutor
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
    /// ClosedXML Provider 名称。
    /// </summary>
    private const string Provider = "ClosedXML";

    /// <summary>
    /// 用于创建实体列映射计划的构建器。
    /// </summary>
    private readonly ClosedXmlMappingPlanBuilder _planBuilder;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlEntityLayoutExecutor" /> 类型的实例。
    /// </summary>
    /// <param name="mappingPlanFactory">公共映射计划工厂。</param>
    public ClosedXmlEntityLayoutExecutor(IExcelMappingPlanFactory mappingPlanFactory)
    {
        _planBuilder = new ClosedXmlMappingPlanBuilder(mappingPlanFactory ??
            throw new ArgumentNullException(nameof(mappingPlanFactory)));
    }

    /// <summary>
    /// 校验模板是否包含实体布局声明的工作表和合并区域。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">待校验的工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public void ValidateTemplate<TEntity>(XLWorkbook workbook, ExcelEntityLayout<TEntity> layout,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        foreach (var merge in layout.Merges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, merge.SheetName, template: true);
            EnsureMerge(sheet, merge.Range, addMissing: false, requireExisting: true);
        }
        foreach (var cell in layout.Cells)
        {
            var sheet = ResolveSheet(workbook, cell.SheetName, template: true);
            if (cell.AnchorName != null)
                ResolveNamedAnchor(sheet, cell.AnchorName);
        }
        foreach (var region in layout.ListRegions)
        {
            var sheet = ResolveSheet(workbook, region.SheetName, template: true);
            if (region.AnchorName != null)
                ResolveNamedAnchor(sheet, region.AnchorName);
        }
    }

    /// <summary>
    /// 将实体固定单元格、列表区域和合并声明写入工作簿。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="template">是否按模板模式写入。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public void Write<TEntity>(XLWorkbook workbook, TEntity entity, ExcelEntityLayout<TEntity> layout,
        bool template, CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        if (entity == null)
            throw new ArgumentNullException(nameof(entity));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));

        foreach (var merge in layout.Merges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, merge.SheetName, template);
            EnsureMerge(sheet, merge.Range, addMissing: !template, requireExisting: template);
        }
        foreach (var binding in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, binding.SheetName, template);
            var resolvedBinding = binding;
            if (binding.AnchorName != null)
            {
                var named = ResolveNamedAnchor(sheet, binding.AnchorName);
                sheet = named.Sheet;
                resolvedBinding = binding.WithResolvedReference(named.Reference);
            }
            var column = CreateColumn(typeof(TEntity), resolvedBinding.PropertyName, resolvedBinding.ConverterName,
                resolvedBinding.MappingDocument, resolvedBinding.MappingConfiguration, MappingDirection.Export);
            var cell = sheet.Cell(resolvedBinding.Reference.Row + 1, resolvedBinding.Reference.Column + 1);
            var raw = resolvedBinding.Getter(entity);
            var converted = ClosedXmlValueAdapter.ConvertTo(raw, column.Column, column.Property,
                sheet.Name, resolvedBinding.Reference.Row + 1, resolvedBinding.Reference.Column + 1,
                CultureInfo.InvariantCulture);
            ValidateExport(column, raw, converted, sheet.Name, resolvedBinding.Reference);
            WriteValue(cell, converted);
        }
        var footprints = new List<EntityListFootprint>();
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, region.SheetName, template);
            var resolvedRegion = region;
            if (region.AnchorName != null)
                resolvedRegion = region.WithResolvedStart(ResolveNamedAnchor(sheet,
                    region.AnchorName).Reference);
            ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            WriteRegion(workbook, entity, resolvedRegion, template, footprints, cancellationToken);
        }
    }

    /// <summary>
    /// 从工作簿读取实体固定单元格和列表区域，并绑定关系。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">来源工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="requireTemplateMerges">是否要求布局声明的合并区域已存在。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>实体、列表区域结果和结构化导入错误。</returns>
    public ExcelEntityImportResult<TEntity> Read<TEntity>(XLWorkbook workbook, ExcelEntityLayout<TEntity> layout,
        bool requireTemplateMerges, bool isDate1904, ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));

        foreach (var merge in layout.Merges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, merge.SheetName, template: true);
            EnsureMerge(sheet, merge.Range, addMissing: false, requireExisting: requireTemplateMerges);
        }

        limits?.Validate();
        var entity = new TEntity();
        var errorCollector = new LimitedErrorCollection(limits?.MaxErrors);
        ICollection<ExcelImportError> errors = errorCollector;
        foreach (var binding in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, binding.SheetName, template: true);
            var resolvedBinding = binding;
            if (binding.AnchorName != null)
                resolvedBinding = binding.WithResolvedReference(ResolveNamedAnchor(sheet,
                    binding.AnchorName).Reference);
            var column = CreateColumn(typeof(TEntity), resolvedBinding.PropertyName, resolvedBinding.ConverterName,
                resolvedBinding.MappingDocument, resolvedBinding.MappingConfiguration, MappingDirection.Import);
            if (resolvedBinding.Setter == null)
                throw new BingOfficesConfigurationException($"实体属性不可写入: {resolvedBinding.PropertyName}",
                    stage: BingOfficesStage.Plan);
            var cell = sheet.Cell(resolvedBinding.Reference.Row + 1, resolvedBinding.Reference.Column + 1);
            var raw = ReadRaw(cell, resolvedBinding.PropertyType);
            var text = ClosedXmlValueAdapter.ToText(raw, CultureInfo.InvariantCulture);
            var cellValue = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
            if (!ValidateImport(column, raw, null, text, cellValue, sheet.Name,
                    resolvedBinding.Reference, errors, rawOnly: true))
                continue;
            try
            {
                var converted = ClosedXmlValueAdapter.ConvertFrom(raw, column.Column, column.Property,
                    sheet.Name, resolvedBinding.Reference.Row + 1, resolvedBinding.Reference.Column + 1,
                    CultureInfo.InvariantCulture, isDate1904, cellValue);
                if (!ValidateImport(column, raw, converted, text, cellValue, sheet.Name,
                        resolvedBinding.Reference, errors, rawOnly: false))
                    continue;
                resolvedBinding.Setter?.Invoke(entity, converted);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message,
                    binding.SheetName, binding.Reference.Row + 1, binding.Reference.Column + 1,
                    binding.PropertyName, rawValue: raw));
            }
        }

        var sheetResults = new List<ExcelSheetImportResult>();
        var totalRows = 0;
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = ResolveSheet(workbook, region.SheetName, template: true);
            var resolvedRegion = region;
            if (region.AnchorName != null)
                resolvedRegion = region.WithResolvedStart(ResolveNamedAnchor(sheet,
                    region.AnchorName).Reference);
            ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            sheetResults.Add(ReadRegion(workbook, entity, resolvedRegion, isDate1904, errors, limits,
                validationFailureMode, ref totalRows, cancellationToken));
        }
        ClosedXmlRelationCoordinator.Bind(entity, layout.Relations, errors, limits?.MaxErrors,
            cancellationToken);
        return new ExcelEntityImportResult<TEntity>(entity, errorCollector.Items, sheetResults,
            errorCollector.IsTruncated, limits?.MaxErrors);
    }

    /// <summary>
    /// 将一个实体列表区域写入工作表。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="entity">根实体。</param>
    /// <param name="region">列表区域声明。</param>
    /// <param name="template">是否按模板模式写入。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void WriteRegion<TEntity>(XLWorkbook workbook, TEntity entity,
        ExcelEntityListRegion<TEntity> region, bool template,
        ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var sheet = ResolveSheet(workbook, region.SheetName, template);
        var values = (region.Getter(entity) as IEnumerable)?.Cast<object>().ToArray()
            ?? Array.Empty<object>();
        var plan = CreateRegionPlan(region, MappingDirection.Export);
        if (region.GroupSubtotal != null)
        {
            WriteGroupedRegion(workbook, sheet, region, plan, values, template,
                footprints, cancellationToken);
            return;
        }
        ValidateBounds(region, plan.Width, values.Length);
        var dataEndRow = region.PageSubtotal != null
            ? ValidatePagedFooterBounds(region, plan.Width, values.Length)
            : region.Start.Row + (region.IncludeHeader ? 1 : 0) + values.Length;
        ValidateFooterBounds(region, region.Footer, dataEndRow);
        RegisterRuntimeFootprint(region, plan.Width, values.Length, footprints);
        cancellationToken.ThrowIfCancellationRequested();
        if (region.PageSubtotal != null)
        {
            WritePagedRegion(sheet, region, plan, values, template, cancellationToken);
            return;
        }
        PrepareFooterMerges(sheet, region.Footer, dataEndRow, region.Start.Column);
        var row = region.Start.Row;
        if (region.IncludeHeader)
        {
            foreach (var column in plan.Columns)
            {
                var header = sheet.Cell(row + 1, region.Start.Column + column.PhysicalColumnIndex + 1);
                header.Value = column.Title;
                ApplyCalculatedStyle(header, column, header: true);
            }
            row++;
        }
        var calculatedItems = values.ToArray();
        var calculatedEvaluators = plan.Columns.Where(column => column.Calculated != null)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, region.SheetName),
                StringComparer.OrdinalIgnoreCase);
        for (var itemIndex = 0; itemIndex < values.Length; itemIndex++)
        {
            var item = values[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            if (item == null)
                throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                    stage: BingOfficesStage.Plan);
            var dynamicValues = GetDynamicValues(item, plan, create: false);
            ValidateUnknownDynamicValues(region, plan, item, dynamicValues);
            foreach (var column in plan.Columns)
            {
                var columnIndex = region.Start.Column + column.PhysicalColumnIndex;
                object raw = null;
                object converted;
                if (column.Calculated != null)
                {
                    try
                    {
                        converted = calculatedEvaluators[column.Key](item, itemIndex, row + 1,
                            columnIndex + 1);
                    }
                    catch (BingOfficesException)
                    {
                        throw;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception) when (exception is not OutOfMemoryException
                        && exception is not StackOverflowException)
                    {
                        throw new BingOfficesExportException("Excel 实体计算列执行失败。", exception,
                            Provider, BingOfficesStage.Validate, region.SheetName, row + 1,
                            columnIndex + 1, column.Key, BingOfficesErrorCode.UserExtensionFailed);
                    }
                    WriteValue(sheet.Cell(row + 1, columnIndex + 1), converted);
                    ApplyCalculatedStyle(sheet.Cell(row + 1, columnIndex + 1), column, header: false);
                    continue;
                }
                if (column.Dynamic != null)
                {
                    dynamicValues?.TryGetValue(column.Key, out raw);
                    converted = ClosedXmlValueAdapter.ConvertDynamicTo(raw, column.Dynamic,
                        region.SheetName, row + 1, columnIndex + 1, CultureInfo.InvariantCulture);
                    ValidateDynamicExport(column.Dynamic, raw, converted, region.SheetName,
                        ExcelEntityCellReference.Parse(ToAddress(row, columnIndex)));
                }
                else
                {
                    raw = column.Property.GetValue(item);
                    converted = ClosedXmlValueAdapter.ConvertTo(raw, column.Fixed, column.Property,
                        region.SheetName, row + 1, columnIndex + 1, CultureInfo.InvariantCulture);
                    ValidateExport(new EntityColumn(column.Fixed, column.Property), raw, converted,
                        region.SheetName, ExcelEntityCellReference.Parse(ToAddress(row, columnIndex)));
                }
                WriteValue(sheet.Cell(row + 1, columnIndex + 1), converted);
                ApplyCalculatedStyle(sheet.Cell(row + 1, columnIndex + 1), column, header: false);
            }
            row++;
        }
        WriteFooter(sheet, region.Footer, values, row, region.Start.Column, template,
            cancellationToken, region.FooterAnchorName);
        ApplyPageBreaks(sheet, region, values.Length);
    }

    /// <summary>
    /// 将带分页小计的列表项写入 ClosedXML 工作表。
    /// </summary>
    private void WritePagedRegion<TEntity>(IXLWorksheet sheet, ExcelEntityListRegion<TEntity> region,
        EntityRegionPlan plan, IReadOnlyList<object> values, bool template,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (region.PageSubtotal == null || region.PageBreakRows is not int rowsPerPage)
            throw new BingOfficesConfigurationException("分页小计缺少分页配置。",
                stage: BingOfficesStage.Plan);
        var plannedRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var plannedIndex = 0;
        while (plannedIndex < values.Count)
        {
            var pageCount = Math.Min(rowsPerPage, values.Count - plannedIndex);
            plannedRow += pageCount;
            plannedIndex += pageCount;
            if (plannedIndex < values.Count)
            {
                PrepareFooterMerges(sheet, region.PageSubtotal, plannedRow, region.Start.Column);
                plannedRow += region.PageSubtotal.GapRows + GetFooterHeight(region.PageSubtotal);
            }
        }
        PrepareFooterMerges(sheet, region.Footer, plannedRow, region.Start.Column);
        var row = region.Start.Row;
        if (region.IncludeHeader)
        {
            foreach (var column in plan.Columns)
            {
                var header = sheet.Cell(row + 1, region.Start.Column + column.PhysicalColumnIndex + 1);
                header.Value = column.Title;
                ApplyCalculatedStyle(header, column, header: true);
            }
            row++;
        }
        var calculatedItems = values.ToArray();
        var calculatedEvaluators = plan.Columns.Where(column => column.Calculated != null)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, region.SheetName),
                StringComparer.OrdinalIgnoreCase);
        var itemIndex = 0;
        var detailRows = new List<(int FirstRow, int LastRow)>();
        while (itemIndex < values.Count)
        {
            var firstDetailRow = row + 1;
            var pageStart = itemIndex;
            var pageCount = Math.Min(rowsPerPage, values.Count - itemIndex);
            for (var pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                var item = values[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                if (item == null)
                    throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                        stage: BingOfficesStage.Plan);
                var dynamicValues = GetDynamicValues(item, plan, create: false);
                ValidateUnknownDynamicValues(region, plan, item, dynamicValues);
                foreach (var column in plan.Columns)
                {
                    var columnIndex = region.Start.Column + column.PhysicalColumnIndex;
                    object raw = null;
                    object converted;
                    if (column.Calculated != null)
                    {
                        try
                        {
                            converted = calculatedEvaluators[column.Key](item, itemIndex, row + 1,
                                columnIndex + 1);
                        }
                        catch (BingOfficesException)
                        {
                            throw;
                        }
                        catch (OperationCanceledException)
                        {
                            throw;
                        }
                        catch (Exception exception) when (exception is not OutOfMemoryException
                            && exception is not StackOverflowException)
                        {
                            throw new BingOfficesExportException("Excel 实体计算列执行失败。", exception,
                                Provider, BingOfficesStage.Validate, region.SheetName, row + 1,
                                columnIndex + 1, column.Key, BingOfficesErrorCode.UserExtensionFailed);
                        }
                        WriteValue(sheet.Cell(row + 1, columnIndex + 1), converted);
                        ApplyCalculatedStyle(sheet.Cell(row + 1, columnIndex + 1), column, header: false);
                        continue;
                    }
                    if (column.Dynamic != null)
                    {
                        dynamicValues?.TryGetValue(column.Key, out raw);
                        converted = ClosedXmlValueAdapter.ConvertDynamicTo(raw, column.Dynamic,
                            region.SheetName, row + 1, columnIndex + 1, CultureInfo.InvariantCulture);
                        ValidateDynamicExport(column.Dynamic, raw, converted, region.SheetName,
                            ExcelEntityCellReference.Parse(ToAddress(row, columnIndex)));
                    }
                    else
                    {
                        raw = column.Property.GetValue(item);
                        converted = ClosedXmlValueAdapter.ConvertTo(raw, column.Fixed, column.Property,
                            region.SheetName, row + 1, columnIndex + 1, CultureInfo.InvariantCulture);
                        ValidateExport(new EntityColumn(column.Fixed, column.Property), raw, converted,
                            region.SheetName, ExcelEntityCellReference.Parse(ToAddress(row, columnIndex)));
                    }
                    WriteValue(sheet.Cell(row + 1, columnIndex + 1), converted);
                    ApplyCalculatedStyle(sheet.Cell(row + 1, columnIndex + 1), column, header: false);
                }
                row++;
                itemIndex++;
            }
            detailRows.Add((firstDetailRow, row));
            if (itemIndex < values.Count)
            {
                var pageItems = values.Skip(pageStart).Take(pageCount).ToArray();
                WriteFooter(sheet, region.PageSubtotal, pageItems, row, region.Start.Column,
                    template, cancellationToken);
                row += region.PageSubtotal.GapRows + GetFooterHeight(region.PageSubtotal);
                ApplyPageBreak(sheet, row - 1);
            }
        }
        WriteFooter(sheet, region.Footer, values, row, region.Start.Column, template,
            cancellationToken, region.FooterAnchorName, detailRows);
    }

    /// <summary>
    /// 将带连续分组小计的列表项写入 ClosedXML 工作表。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="plan">列表区域映射计划。</param>
    /// <param name="values">列表项快照。</param>
    /// <param name="template">是否按模板模式写入。</param>
    /// <param name="footprints">已写入或预检的列表区域范围。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void WriteGroupedRegion<TEntity>(XLWorkbook workbook, IXLWorksheet sheet,
        ExcelEntityListRegion<TEntity> region, EntityRegionPlan plan, IReadOnlyList<object> values,
        bool template, ICollection<EntityListFootprint> footprints, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var subtotal = region.GroupSubtotal;
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var row = firstDataRow;
        var groupRanges = GetGroupedRanges(values, subtotal);
        var previewRow = firstDataRow;
        foreach (var range in groupRanges)
        {
            var dataEndRow = previewRow + range.Count;
            ValidateFooterBounds(region, subtotal.Footer, dataEndRow);
            previewRow = dataEndRow + subtotal.Footer.GapRows + GetFooterHeight(subtotal.Footer);
        }
        ValidateFooterBounds(region, region.Footer, previewRow);
        var lastRow = previewRow + (region.Footer == null ? -1
            : region.Footer.GapRows + GetFooterHeight(region.Footer) - 1);
        if (region.End.HasValue && (lastRow > region.End.Value.Row
            || region.Start.Column + plan.Width - 1 > region.End.Value.Column))
            throw new BingOfficesConfigurationException(
                $"列表区域分组小计超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        RegisterRuntimeFootprint(region, plan.Width, values.Count, footprints, values);
        cancellationToken.ThrowIfCancellationRequested();
        var mergeRow = firstDataRow;
        foreach (var range in groupRanges)
        {
            var groupDataEndRow = mergeRow + range.Count;
            PrepareFooterMerges(sheet, subtotal.Footer, groupDataEndRow, region.Start.Column);
            mergeRow = groupDataEndRow + subtotal.Footer.GapRows + GetFooterHeight(subtotal.Footer);
        }
        PrepareFooterMerges(sheet, region.Footer, previewRow, region.Start.Column);

        if (region.IncludeHeader)
        {
            foreach (var column in plan.Columns)
            {
                var header = sheet.Cell(region.Start.Row + 1,
                    region.Start.Column + column.PhysicalColumnIndex + 1);
                header.Value = column.Title;
                ApplyCalculatedStyle(header, column, header: true);
            }
        }

        var calculatedItems = values.ToArray();
        var calculatedEvaluators = plan.Columns.Where(column => column.Calculated != null)
            .ToDictionary(column => column.Key,
                column => column.Calculated.CreateEvaluator(calculatedItems, region.SheetName),
                StringComparer.OrdinalIgnoreCase);
        void WriteItem(object item, int itemIndex)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item == null)
                throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                    stage: BingOfficesStage.Plan);
            var dynamicValues = GetDynamicValues(item, plan, create: false);
            ValidateUnknownDynamicValues(region, plan, item, dynamicValues);
            foreach (var column in plan.Columns)
            {
                var columnIndex = region.Start.Column + column.PhysicalColumnIndex;
                object raw = null;
                object converted;
                if (column.Calculated != null)
                {
                    try
                    {
                        converted = calculatedEvaluators[column.Key](item, itemIndex, row + 1,
                            columnIndex + 1);
                    }
                    catch (BingOfficesException)
                    {
                        throw;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception) when (exception is not OutOfMemoryException
                        && exception is not StackOverflowException)
                    {
                        throw new BingOfficesExportException("Excel 实体计算列执行失败。", exception,
                            Provider, BingOfficesStage.Validate, region.SheetName, row + 1,
                            columnIndex + 1, column.Key, BingOfficesErrorCode.UserExtensionFailed);
                    }
                    WriteValue(sheet.Cell(row + 1, columnIndex + 1), converted);
                    ApplyCalculatedStyle(sheet.Cell(row + 1, columnIndex + 1), column, header: false);
                    continue;
                }
                if (column.Dynamic != null)
                {
                    dynamicValues?.TryGetValue(column.Key, out raw);
                    converted = ClosedXmlValueAdapter.ConvertDynamicTo(raw, column.Dynamic,
                        region.SheetName, row + 1, columnIndex + 1, CultureInfo.InvariantCulture);
                    ValidateDynamicExport(column.Dynamic, raw, converted, region.SheetName,
                        ExcelEntityCellReference.Parse(ToAddress(row, columnIndex)));
                }
                else
                {
                    raw = column.Property.GetValue(item);
                    converted = ClosedXmlValueAdapter.ConvertTo(raw, column.Fixed, column.Property,
                        region.SheetName, row + 1, columnIndex + 1, CultureInfo.InvariantCulture);
                    ValidateExport(new EntityColumn(column.Fixed, column.Property), raw, converted,
                        region.SheetName, ExcelEntityCellReference.Parse(ToAddress(row, columnIndex)));
                }
                WriteValue(sheet.Cell(row + 1, columnIndex + 1), converted);
                ApplyCalculatedStyle(sheet.Cell(row + 1, columnIndex + 1), column, header: false);
            }
            row++;
        }

        var itemIndexOffset = 0;
        var detailRows = new List<(int FirstRow, int LastRow)>();
        foreach (var range in groupRanges)
        {
            var firstDetailRow = row + 1;
            var groupItems = values.Skip(range.Start).Take(range.Count).ToArray();
            for (var itemIndex = 0; itemIndex < range.Count; itemIndex++)
                WriteItem(values[range.Start + itemIndex], itemIndexOffset + itemIndex);
            detailRows.Add((firstDetailRow, row));
            WriteFooter(sheet, subtotal.Footer, groupItems, row, region.Start.Column, template,
                cancellationToken);
            row += subtotal.Footer.GapRows + GetFooterHeight(subtotal.Footer);
            itemIndexOffset += range.Count;
        }
        WriteFooter(sheet, region.Footer, values, row, region.Start.Column, template,
            cancellationToken, region.FooterAnchorName, detailRows);
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
                    + GetFooterHeight(region.GroupSubtotal.Footer);
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
                row += region.PageSubtotal.GapRows + GetFooterHeight(region.PageSubtotal);
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
    /// 从一个实体列表区域读取项目并写回根实体集合。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="workbook">来源工作簿。</param>
    /// <param name="entity">根实体。</param>
    /// <param name="region">列表区域声明。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="errors">共享导入错误集合。</param>
    /// <param name="limits">实体导入使用的资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="totalRows">跨列表区域累计读取的数据行数。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>列表区域导入结果。</returns>
    private ExcelSheetImportResult ReadRegion<TEntity>(XLWorkbook workbook, TEntity entity,
        ExcelEntityListRegion<TEntity> region, bool isDate1904, ICollection<ExcelImportError> errors,
        ExcelResourceLimits limits, ExcelValidationFailureMode validationFailureMode,
        ref int totalRows, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var sheet = ResolveSheet(workbook, region.SheetName, template: true);
        var plan = CreateRegionPlan(region, MappingDirection.Import);
        var maxItems = GetRegionItemCapacity(region, plan.Width, sheet);
        ValidateBounds(region, plan.Width, maxItems);
        var rows = new List<int>();
        var items = new List<object>();
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var uniqueTracker = new UniqueTracker(duplicateValues, limits?.MaxTrackedUniqueValues,
            CreateStringComparer(limits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase));
        var row = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var markerRow = FindFooterMarkerRow(sheet, region, row);
        var lastRow = markerRow.HasValue ? markerRow.Value - 1 - (region.Footer?.GapRows ?? 0)
            : (region.End?.Row ?? sheet.LastRowUsed()?.RowNumber() - 1 ?? row - 1);
        var skipSpans = region.GroupSubtotal == null && region.PageSubtotal == null
            ? Array.Empty<(int StartRow, int EndRow)>()
            : FindFooterSkipSpans(sheet, region, row)
                .Where(span => span.StartRow <= lastRow).ToArray();
        while (row <= lastRow)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var skip = skipSpans.FirstOrDefault(span => row >= span.StartRow && row <= span.EndRow);
            if (skipSpans.Any(span => row >= span.StartRow && row <= span.EndRow))
            {
                row = skip.EndRow + 1;
                continue;
            }
            if (IsEmptyRow(sheet, row, region.Start.Column, plan.Width))
                break;
            if (limits?.MaxRows.HasValue == true && totalRows >= limits.MaxRows.Value)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    $"Workbook 数据行数超过限制: {limits.MaxRows.Value}", region.SheetName,
                    row + 1, 0, null));
                break;
            }
            totalRows++;
            var item = Activator.CreateInstance(region.ItemType);
            var dynamicValues = GetDynamicValues(item, plan, create: true);
            var valid = true;
            uniqueTracker.BeginRow();
            foreach (var column in plan.Columns)
            {
                var columnIndex = region.Start.Column + column.PhysicalColumnIndex;
                var reference = ExcelEntityCellReference.Parse(ToAddress(row, columnIndex));
                if (column.Calculated != null)
                    continue;
                var cell = sheet.Cell(row + 1, columnIndex + 1);
                var raw = ReadRaw(cell, column.Property?.PropertyType ?? typeof(object));
                var text = ClosedXmlValueAdapter.ToText(raw, CultureInfo.InvariantCulture);
                var cellValue = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
                if (column.Dynamic != null && !ValidateDynamicImport(column.Dynamic, raw, null, text,
                        cellValue, region.SheetName, reference, errors, rawOnly: true))
                {
                    valid = false;
                    if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                        break;
                    continue;
                }
                if (column.Fixed != null && !ValidateImport(new EntityColumn(column.Fixed, column.Property),
                        raw, null, text, cellValue, region.SheetName, reference, errors, rawOnly: true))
                {
                    valid = false;
                    if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                        break;
                    continue;
                }
                try
                {
                    var converted = column.Dynamic != null
                        ? ClosedXmlValueAdapter.ConvertDynamicFrom(raw, column.Dynamic, region.SheetName,
                            row + 1, columnIndex + 1, CultureInfo.InvariantCulture, isDate1904, cellValue)
                        : ClosedXmlValueAdapter.ConvertFrom(raw, column.Fixed, column.Property,
                            region.SheetName, row + 1, columnIndex + 1,
                            CultureInfo.InvariantCulture, isDate1904, cellValue);
                    if (column.Dynamic != null && !ValidateDynamicImport(column.Dynamic, raw, converted,
                            text, cellValue, region.SheetName, reference, errors, rawOnly: false))
                    {
                        valid = false;
                        if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                        continue;
                    }
                    if (column.Fixed != null && !ValidateImport(new EntityColumn(column.Fixed, column.Property),
                            raw, converted, text, cellValue, region.SheetName, reference, errors, rawOnly: false))
                    {
                        valid = false;
                        if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                        continue;
                    }
                    if (column.Dynamic != null && !ValidateUniqueImport(column.Dynamic.Key,
                            column.Dynamic.IsUnique, column.Dynamic.UniqueIgnoreEmpty, text,
                            region.SheetName, reference, errors, uniqueTracker))
                    {
                        valid = false;
                        if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                        continue;
                    }
                    if (column.Fixed != null && !ValidateUniqueImport(column.Fixed.Name,
                            column.Fixed.IsUnique, column.Fixed.UniqueIgnoreEmpty, text,
                            region.SheetName, reference, errors, uniqueTracker))
                    {
                        valid = false;
                        if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                        continue;
                    }
                    if (column.Dynamic != null)
                        SetDynamicValue(item, plan, column.Key, converted, dynamicValues);
                    else
                        column.Property.SetValue(item, converted);
                }
                catch (Exception exception) when (exception is not OperationCanceledException
                    && exception is not OutOfMemoryException && exception is not StackOverflowException)
                {
                    valid = false;
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion,
                        exception.Message, region.SheetName, row + 1, columnIndex + 1,
                        column.Dynamic?.Key ?? column.Property.Name, rawValue: raw));
                    if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                        break;
                }
            }
            if (valid)
            {
                items.Add(item);
                rows.Add(row);
                uniqueTracker.CommitRow();
            }
            else
                uniqueTracker.RollbackRow();
            row++;
        }
        if (region.Setter != null)
        {
            var typed = CreateTypedList(region.ItemType, items);
            region.Setter(entity, typed);
        }
        return new ExcelSheetImportResult(region.SheetName, region.ItemType, rows,
            errors.Where(error => string.Equals(error.SheetName, region.SheetName,
                StringComparison.OrdinalIgnoreCase)).ToArray());
    }

    /// <summary>
    /// 为列表项目类型创建固定列和动态列的统一物理布局。
    /// </summary>
    /// <param name="region">列表区域声明。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>包含最终物理索引和动态字典属性的区域计划。</returns>
    private EntityRegionPlan CreateRegionPlan<TEntity>(ExcelEntityListRegion<TEntity> region,
        MappingDirection direction) where TEntity : class, new()
    {
        var configuration = CreateEffectiveConfiguration(region);
        var plan = _planBuilder.Create(region.ItemType, region.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, configuration, direction);
        var definitions = region.DynamicColumnGroups.SelectMany(group => group.Definitions).ToArray();
        var columns = ClosedXmlExportColumnPlanner.Create(region.ItemType, plan,
                definitions, configuration, region.CalculatedColumns).ToArray();
        if (direction == MappingDirection.Import)
        {
            columns = columns.Where(column => column.Calculated != null
                || column.Dynamic != null || !IsNavigationProperty(column.Property)).ToArray();
            var readOnly = columns.FirstOrDefault(column => column.Calculated == null
                && column.Dynamic == null
                && column.Property != null && !column.Property.CanWrite
                && !IsNavigationProperty(column.Property));
            if (readOnly != null)
                throw new BingOfficesConfigurationException(
                    $"实体属性不可写入: {readOnly.Property.Name}", stage: BingOfficesStage.Plan);
        }
        if (columns.Length == 0)
            throw new BingOfficesConfigurationException($"实体列表没有可用映射列: {region.ItemType.FullName}",
                stage: BingOfficesStage.Plan);
        var dynamicPlanColumn = plan.DynamicColumns.Count == 0 ? null
            : plan.Columns.SingleOrDefault(column => column.IsDynamicColumn && !column.Ignored);
        var dynamicProperty = dynamicPlanColumn == null ? null : region.ItemType.GetProperty(dynamicPlanColumn.Name,
            BindingFlags.Instance | BindingFlags.Public);
        if (plan.DynamicColumns.Count > 0 && region.DynamicColumnGroups.Count == 0 &&
            (dynamicProperty == null || !dynamicProperty.CanRead
            || !typeof(IDictionary<string, object>).IsAssignableFrom(dynamicProperty.PropertyType)))
            throw new BingOfficesConfigurationException(
                "实体动态列属性必须是可读的 IDictionary<string, object>。", stage: BingOfficesStage.Plan);
        var width = columns.Max(column => column.PhysicalColumnIndex) + 1;
        return new EntityRegionPlan(columns, dynamicProperty, region.DynamicColumnGroups, width);
    }

    /// <summary>
    /// 判断属性是否属于关系或集合导航属性。
    /// </summary>
    /// <param name="property">待判断的实体属性。</param>
    /// <returns>集合导航属性返回 <see langword="true" />。</returns>
    private static bool IsNavigationProperty(PropertyInfo property)
    {
        var type = property?.PropertyType;
        return type != null && type != typeof(string) && type != typeof(byte[])
            && typeof(IEnumerable).IsAssignableFrom(type);
    }

    /// <summary>
    /// 合并列表区域映射与显式动态列组，创建供 Provider 使用的配置快照。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="region">列表区域声明。</param>
    /// <returns>包含动态列组定义的映射配置。</returns>
    private static ExcelMappingConfiguration CreateEffectiveConfiguration<TEntity>(
        ExcelEntityListRegion<TEntity> region) where TEntity : class, new()
    {
        var source = region.MappingConfiguration;
        var configuration = new ExcelMappingConfiguration
        {
            Profile = source?.Profile,
            ModelAlias = source?.ModelAlias,
            ClearDynamicColumns = source?.ClearDynamicColumns ?? false,
            DynamicColumnMergeMode = source?.DynamicColumnMergeMode,
            ResetStyle = source?.ResetStyle ?? false,
            ResetLayout = source?.ResetLayout ?? false,
            SourceKind = source?.SourceKind ?? MappingSourceKind.Request,
            Style = source?.Style,
            Layout = source?.Layout,
            Columns = (source?.Columns ?? new List<ExcelColumnConfiguration>())
                .Select(CloneColumnConfiguration).ToList(),
            DynamicColumns = (source?.DynamicColumns ?? new List<ExcelMappingDynamicColumnConfiguration>())
                .ToList(),
            DynamicColumnKeysToRemove = (source?.DynamicColumnKeysToRemove ?? new List<string>()).ToList()
        };
        if (region.DynamicColumnGroups.Count == 0)
            return configuration;

        var groupedProperties = new HashSet<string>(region.DynamicColumnGroups.Select(group => group.Property.Name),
            StringComparer.OrdinalIgnoreCase);
        foreach (var property in groupedProperties)
        {
            var existing = configuration.Columns.FirstOrDefault(column =>
                string.Equals(column.PropertyName, property, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
                configuration.Columns.Add(new ExcelColumnConfiguration { PropertyName = property, Ignored = true });
            else
                existing.Ignored = true;
        }
        foreach (var definition in region.DynamicColumnGroups.SelectMany(group => group.Definitions))
        {
            configuration.DynamicColumns.Add(new ExcelMappingDynamicColumnConfiguration
            {
                Key = definition.Key,
                Title = definition.Title,
                Aliases = (definition.Aliases ?? Array.Empty<string>()).ToList(),
                DataTypeName = GetDataTypeName(definition.DataType),
                Order = definition.Order,
                NumberFormat = definition.NumberFormat,
                ColumnIndex = definition.PhysicalColumnIndex,
                PlacementKey = definition.Placement == null ? null
                    : !string.IsNullOrWhiteSpace(definition.Placement.BeforeKey)
                        ? "before:" + definition.Placement.BeforeKey
                        : "after:" + definition.Placement.AfterKey,
                ConverterName = definition.ConverterName,
                ValidatorName = definition.ValidatorName,
                ValidationRuleNames = (definition.ValidationRuleNames ?? Array.Empty<string>()).ToList(),
                ValidationRules = (definition.ValidationRules ?? Array.Empty<ExcelMappingDynamicValidationConfiguration>())
                    .Select(CloneDynamicValidation).ToList(),
                ImageMultiplicity = definition.ImageMultiplicity
            });
        }
        return configuration;
    }

    /// <summary>
    /// 复制动态验证配置，避免与调用方共享可变实例。
    /// </summary>
    /// <param name="source">待复制的动态验证配置。</param>
    /// <returns>配置副本；源配置为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    private static ExcelMappingDynamicValidationConfiguration CloneDynamicValidation(
        ExcelMappingDynamicValidationConfiguration source) => source == null ? null : new ExcelMappingDynamicValidationConfiguration
        {
            Name = source.Name,
            Pattern = source.Pattern,
            Format = source.Format,
            CultureName = source.CultureName,
            Min = source.Min,
            Max = source.Max,
            MaxValue = source.MaxValue,
            MaxLength = source.MaxLength,
            IgnoreEmpty = source.IgnoreEmpty
        };

    /// <summary>
    /// 复制映射列配置，避免修改调用方的配置对象。
    /// </summary>
    /// <param name="source">待复制的映射列配置。</param>
    /// <returns>独立的映射列配置副本。</returns>
    private static ExcelColumnConfiguration CloneColumnConfiguration(ExcelColumnConfiguration source)
    {
        return new ExcelColumnConfiguration
        {
            PropertyName = source.PropertyName,
            Title = source.Title,
            ClearTitle = source.ClearTitle,
            Aliases = (source.Aliases ?? new List<string>()).ToList(),
            ClearAliases = source.ClearAliases,
            ColumnIndex = source.ColumnIndex,
            ResetColumnIndex = source.ResetColumnIndex,
            Ignored = source.Ignored,
            ResetIgnored = source.ResetIgnored,
            Formatter = source.Formatter,
            ClearFormatter = source.ClearFormatter,
            DecimalScale = source.DecimalScale,
            ResetDecimalScale = source.ResetDecimalScale,
            ConverterName = source.ConverterName,
            ClearConverterName = source.ClearConverterName,
            ImportWhitespace = source.ImportWhitespace,
            ResetImportWhitespace = source.ResetImportWhitespace,
            ValidationRuleNames = (source.ValidationRuleNames ?? new List<string>()).ToList(),
            ValidationRuleNamesToRemove = (source.ValidationRuleNamesToRemove ?? new List<string>()).ToList(),
            ClearValidationRules = source.ClearValidationRules,
            ValidationRuleMergeMode = source.ValidationRuleMergeMode,
            ValueMappings = (source.ValueMappings ?? new List<ExcelValueMappingConfiguration>()).ToList(),
            ClearValueMappings = source.ClearValueMappings,
            ValueMappingMergeMode = source.ValueMappingMergeMode,
            ImageMultiplicity = source.ImageMultiplicity,
            ResetImageMultiplicity = source.ResetImageMultiplicity
        };
    }

    /// <summary>
    /// 将 CLR 类型转换为映射配置使用的数据类型名称。
    /// </summary>
    /// <param name="type">动态列值类型。</param>
    /// <returns>规范化的数据类型名称。</returns>
    private static string GetDataTypeName(Type type)
    {
        type = Nullable.GetUnderlyingType(type ?? typeof(string)) ?? type ?? typeof(string);
        if (type == typeof(bool)) return "bool";
        if (type == typeof(byte)) return "byte";
        if (type == typeof(short)) return "int16";
        if (type == typeof(int)) return "int32";
        if (type == typeof(long)) return "int64";
        if (type == typeof(float)) return "single";
        if (type == typeof(double)) return "double";
        if (type == typeof(decimal)) return "decimal";
        if (type == typeof(DateTime)) return "datetime";
        if (type == typeof(DateTimeOffset)) return "datetimeoffset";
        if (type == typeof(Guid)) return "guid";
        if (type == typeof(byte[])) return "bytes";
        if (type == typeof(object)) return "object";
        return "string";
    }

    /// <summary>
    /// 为实体固定单元格创建指定属性的映射列。
    /// </summary>
    /// <param name="itemType">实体类型。</param>
    /// <param name="propertyName">属性名称。</param>
    /// <param name="converterName">请求指定的转换器名称。</param>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>实体列及其属性信息。</returns>
    private EntityColumn CreateColumn(Type itemType, string propertyName, string converterName,
        ExcelMappingDocument document, ExcelMappingConfiguration configuration, MappingDirection direction)
    {
        if (!string.IsNullOrWhiteSpace(converterName))
        {
            var overlay = new ExcelMappingConfiguration();
            overlay.Columns.Add(new ExcelColumnConfiguration
            {
                PropertyName = propertyName,
                ConverterName = converterName
            });
            configuration = MappingConfigurationMerger.Merge(configuration, overlay,
                MappingSourceKind.Request);
        }
        var plan = _planBuilder.Create(itemType, document ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, configuration, direction);
        var column = plan.Columns.FirstOrDefault(item => string.Equals(item.Name, propertyName,
            StringComparison.OrdinalIgnoreCase));
        var property = itemType.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
        if (column == null || column.Ignored || column.IsDynamicColumn || property == null)
            throw new BingOfficesConfigurationException($"实体固定单元格属性未生成可用映射列: {propertyName}",
                stage: BingOfficesStage.Plan);
        return new EntityColumn(column, property);
    }

    /// <summary>
    /// 按工作表名称解析工作表，不存在时按模板模式决定创建或失败。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="name">工作表名称。</param>
    /// <param name="template">是否按模板模式解析。</param>
    /// <returns>解析或创建的工作表。</returns>
    private static IXLWorksheet ResolveSheet(XLWorkbook workbook, string name, bool template)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(item => string.Equals(item.Name, name,
            StringComparison.OrdinalIgnoreCase));
        if (sheet != null)
            return sheet;
        if (template)
            throw new BingOfficesConfigurationException($"模板缺少请求的 Sheet: {name}",
                stage: BingOfficesStage.Plan);
        return workbook.Worksheets.Add(name);
    }

    /// <summary>
    /// 解析模板命名锚点并返回其所在工作表和左上角坐标。
    /// </summary>
    /// <param name="sheet">布局声明的工作表。</param>
    /// <param name="anchorName">命名范围名称。</param>
    /// <returns>命名锚点所在工作表和坐标。</returns>
    private static (IXLWorksheet Sheet, ExcelEntityCellReference Reference) ResolveNamedAnchor(
        IXLWorksheet sheet, string anchorName)
    {
        if (!ClosedXmlWorkbookValidationPipeline.TryResolveNamedRange(sheet, anchorName, out var range)
            || range == null)
            throw new BingOfficesConfigurationException(
                $"模板缺少或无法解析命名锚点: {sheet.Name}!{anchorName}", stage: BingOfficesStage.Plan);
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        if (first.RowNumber != last.RowNumber || first.ColumnNumber != last.ColumnNumber)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个单元格: {sheet.Name}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (!string.Equals(range.Worksheet.Name, sheet.Name, StringComparison.OrdinalIgnoreCase))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点引用了其他 Sheet: {sheet.Name}!{anchorName}", stage: BingOfficesStage.Plan);
        try
        {
            return (range.Worksheet, ExcelEntityCellReference.Parse(first.ToString()));
        }
        catch (ArgumentException exception)
        {
            throw new BingOfficesConfigurationException(
                $"模板命名锚点地址超出工作表边界: {sheet.Name}!{anchorName}", exception,
                BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 检查模板解析后的列表区域是否覆盖固定单元格。
    /// </summary>
    private static void ValidateCellRegionConflict<TEntity>(XLWorkbook workbook,
        ExcelEntityLayout<TEntity> layout, ExcelEntityListRegion<TEntity> region)
        where TEntity : class, new()
    {
        foreach (var cell in layout.Cells.Where(item => string.Equals(item.SheetName,
                     region.SheetName, StringComparison.OrdinalIgnoreCase)))
        {
            var reference = cell.AnchorName == null ? cell.Reference :
                ResolveNamedAnchor(ResolveSheet(workbook, cell.SheetName, template: true),
                    cell.AnchorName).Reference;
            if (reference.Row >= region.Start.Row && reference.Column >= region.Start.Column
                && (!region.End.HasValue || (reference.Row <= region.End.Value.Row
                    && reference.Column <= region.End.Value.Column)))
                throw new BingOfficesConfigurationException(
                    $"固定单元格与列表区域重叠: {cell.SheetName}!{reference.Address}",
                    stage: BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 校验或创建实体布局声明的合并区域。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="range">期望的合并区域。</param>
    /// <param name="addMissing">缺少时是否创建合并区域。</param>
    /// <param name="requireExisting">缺少时是否报告配置错误。</param>
    private static void EnsureMerge(IXLWorksheet sheet, ExcelEntityCellRange range,
        bool addMissing, bool requireExisting)
    {
        var existing = sheet.MergedRanges.FirstOrDefault(item => string.Equals(item.RangeAddress.ToString(),
            range.Address, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
            return;
        foreach (var candidate in sheet.MergedRanges)
        {
            if (RangesOverlap(candidate.RangeAddress, range) && !string.Equals(
                    candidate.RangeAddress.ToString(), range.Address, StringComparison.OrdinalIgnoreCase))
                throw new BingOfficesConfigurationException(
                    $"工作表 {sheet.Name} 的合并区域与模板已有区域冲突: {range.Address}",
                    stage: BingOfficesStage.Plan);
        }
        if (requireExisting)
            throw new BingOfficesConfigurationException($"模板缺少声明的合并区域: {sheet.Name}!{range.Address}",
                stage: BingOfficesStage.Plan);
        if (addMissing)
            sheet.Range(range.Address).Merge();
    }

    /// <summary>
    /// 判断已有合并区域与实体布局区域是否重叠。
    /// </summary>
    /// <param name="actual">工作表中的实际合并区域。</param>
    /// <param name="expected">实体布局声明区域。</param>
    /// <returns>两个区域重叠时返回 <see langword="true" />。</returns>
    private static bool RangesOverlap(IXLRangeAddress actual, ExcelEntityCellRange expected)
    {
        return actual.FirstAddress.RowNumber <= expected.Last.Row + 1
            && expected.First.Row + 1 <= actual.LastAddress.RowNumber
            && actual.FirstAddress.ColumnNumber <= expected.Last.Column + 1
            && expected.First.Column + 1 <= actual.LastAddress.ColumnNumber;
    }

    /// <summary>
    /// 校验列表区域的列数和项目数是否落在声明边界内。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="region">列表区域声明。</param>
    /// <param name="columnCount">映射列数量。</param>
    /// <param name="itemCount">项目数量。</param>
    /// <param name="allowUnknownRows">是否允许未声明的尾行。</param>
    private static void ValidateBounds<TEntity>(ExcelEntityListRegion<TEntity> region, int columnCount,
        int itemCount, bool allowUnknownRows = false) where TEntity : class, new()
    {
        if (!region.End.HasValue)
            return;
        var end = region.End.Value;
        var availableColumns = end.Column - region.Start.Column + 1;
        var lastRow = region.Start.Row + itemCount + (region.IncludeHeader ? 0 : -1);
        if (columnCount > availableColumns || (!allowUnknownRows && itemCount > 0 && lastRow > end.Row))
            throw new BingOfficesConfigurationException($"列表区域超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 按列表明细行数写入水平分页符。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="itemCount">已写入的明细行数。</param>
    private static void ApplyPageBreaks<TEntity>(IXLWorksheet sheet,
        ExcelEntityListRegion<TEntity> region, int itemCount) where TEntity : class, new()
    {
        if (region.PageBreakRows is not int rowsPerPage || itemCount <= rowsPerPage)
            return;
        var firstDataRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        for (var offset = rowsPerPage; offset < itemCount; offset += rowsPerPage)
            sheet.Row(firstDataRow + offset).AddHorizontalPageBreak();
    }

    /// <summary>
    /// 设置分页小计完成后的水平分页符。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="breakRow">分页符所在的零基行号。</param>
    private static void ApplyPageBreak(IXLWorksheet sheet, int breakRow)
    {
        if (sheet == null || breakRow < 0)
            return;
        sheet.Row(breakRow + 1).AddHorizontalPageBreak();
    }

    /// <summary>
    /// 校验分页小计和最终尾部的声明及物理边界。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="region">列表区域定义。</param>
    /// <param name="columnCount">映射列数量。</param>
    /// <param name="itemCount">明细数量。</param>
    /// <returns>最终明细结束后的下一行零基索引。</returns>
    private static int ValidatePagedFooterBounds<TEntity>(ExcelEntityListRegion<TEntity> region,
        int columnCount, int itemCount) where TEntity : class, new()
    {
        if (region.PageSubtotal == null || region.PageBreakRows is not int rowsPerPage)
            return region.Start.Row + (region.IncludeHeader ? 1 : 0) + itemCount;
        ValidateBounds(region, columnCount, itemCount);
        var row = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var written = 0;
        while (written < itemCount)
        {
            var count = Math.Min(rowsPerPage, itemCount - written);
            row += count;
            // 分页尾部和签字区也占用工作表行，末页明细必须按推移后的范围预检。
            ValidateBounds(region, columnCount, row - region.Start.Row - (region.IncludeHeader ? 1 : 0));
            if (row - 1 > 1_048_575 || (long)region.Start.Column + columnCount - 1 > 16_383)
                throw new BingOfficesConfigurationException(
                    $"分页列表超出工作表物理边界: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
            written += count;
            if (written >= itemCount)
                break;
            ValidateFooterBounds(region, region.PageSubtotal, row);
            row += region.PageSubtotal.GapRows + GetFooterHeight(region.PageSubtotal);
        }
        ValidateFooterBounds(region, region.Footer, row);
        return row;
    }

    /// <summary>
    /// 计算列表区域可容纳的项目数量。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="region">列表区域声明。</param>
    /// <param name="columnCount">映射列数量。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <returns>区域可容纳的项目数量。</returns>
    private static int GetRegionItemCapacity<TEntity>(ExcelEntityListRegion<TEntity> region, int columnCount,
        IXLWorksheet sheet) where TEntity : class, new()
    {
        if (region.End.HasValue)
            return Math.Max(0, region.End.Value.Row - region.Start.Row + 1 - (region.IncludeHeader ? 1 : 0));
        var last = sheet.LastRowUsed()?.RowNumber() - 1 ?? region.Start.Row - 1;
        return Math.Max(0, last - region.Start.Row + 1 - (region.IncludeHeader ? 1 : 0));
    }

    /// <summary>
    /// 判断指定行的区域单元格是否全部为空。
    /// </summary>
    /// <param name="sheet">工作表。</param>
    /// <param name="row">零基行号。</param>
    /// <param name="column">零基起始列号。</param>
    /// <param name="count">检查的列数。</param>
    /// <returns>全部为空时返回 <see langword="true" />。</returns>
    private static bool IsEmptyRow(IXLWorksheet sheet, int row, int column, int count)
    {
        for (var index = 0; index < count; index++)
        {
            var cell = sheet.Cell(row + 1, column + index + 1);
            if (!cell.IsEmpty())
                return false;
        }
        return true;
    }

    /// <summary>
    /// 将对象序列物化为指定元素类型的泛型列表。
    /// </summary>
    /// <param name="itemType">列表元素类型。</param>
    /// <param name="items">对象序列。</param>
    /// <returns>指定元素类型的列表实例。</returns>
    private static object CreateTypedList(Type itemType, IEnumerable<object> items)
    {
        var listType = typeof(List<>).MakeGenericType(itemType);
        var list = (IList)Activator.CreateInstance(listType);
        foreach (var item in items)
            list.Add(item);
        return list;
    }

    /// <summary>
    /// 获取实体项目的动态列字典，并在导入需要时创建可写实例。
    /// </summary>
    /// <param name="item">列表项目。</param>
    /// <param name="plan">列表区域的映射计划。</param>
    /// <param name="create">属性为空时是否创建字典。</param>
    /// <returns>动态列字典；布局没有动态列时返回 <see langword="null" />。</returns>
    private static IDictionary<string, object> GetDynamicValues(object item, EntityRegionPlan plan, bool create)
    {
        if (plan.DynamicGroups == null || plan.DynamicGroups.Count == 0)
            return GetDynamicValues(item, plan.DynamicProperty, create);
        var combined = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (var group in plan.DynamicGroups)
        {
            var values = group.Getter(item);
            if (values == null)
            {
                if (!create)
                    continue;
                if (group.Setter == null)
                    throw new BingOfficesConfigurationException(
                        $"实体动态列属性为空且不可写: {group.Property.DeclaringType?.FullName}.{group.Property.Name}",
                        stage: BingOfficesStage.Plan);
                values = CreateDynamicDictionary(item, group.Property, group.Setter);
            }
            foreach (var pair in values)
            {
                if (combined.ContainsKey(pair.Key))
                    throw new BingOfficesConfigurationException($"实体动态列组包含重复 Key: {pair.Key}",
                        stage: BingOfficesStage.Plan);
                combined.Add(pair.Key, pair.Value);
            }
        }
        return combined;
    }

    /// <summary>
    /// 读取对象上的动态列字典，并按需创建缺失的字典。
    /// </summary>
    /// <param name="item">动态列所属对象。</param>
    /// <param name="property">动态列字典属性。</param>
    /// <param name="create">属性值为空时是否创建字典并写回属性。</param>
    /// <returns>已有或新建的动态列字典；属性为空且未请求创建，或属性为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    private static IDictionary<string, object> GetDynamicValues(object item, PropertyInfo property, bool create)
    {
        if (property == null)
            return null;
        var current = property.GetValue(item);
        if (current is IDictionary<string, object> values)
            return values;
        if (current != null)
            throw new BingOfficesConfigurationException(
                $"实体动态列属性类型不受支持: {property.DeclaringType?.FullName}.{property.Name}",
                stage: BingOfficesStage.Plan);
        if (!create)
            return null;
        if (!property.CanWrite)
            throw new BingOfficesConfigurationException(
                $"实体动态列属性为空且不可写: {property.DeclaringType?.FullName}.{property.Name}",
                stage: BingOfficesStage.Plan);
        return CreateDynamicDictionary(item, property, (target, value) => property.SetValue(target, value));
    }

    /// <summary>
    /// 将导入得到的动态列值写回所属字典组。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <param name="plan">列表区域映射计划。</param>
    /// <param name="key">动态列键。</param>
    /// <param name="value">转换后的动态列值。</param>
    /// <param name="combined">当前列表项的动态值快照。</param>
    private static void SetDynamicValue(object item, EntityRegionPlan plan, string key, object value,
        IDictionary<string, object> combined)
    {
        if (plan.DynamicGroups == null || plan.DynamicGroups.Count == 0)
        {
            combined[key] = value;
            return;
        }
        var group = plan.DynamicGroups.FirstOrDefault(candidate => candidate.Definitions.Any(definition =>
            string.Equals(definition.Key, key, StringComparison.OrdinalIgnoreCase)));
        if (group == null)
            throw new BingOfficesConfigurationException($"动态列未找到所属组: {key}", stage: BingOfficesStage.Plan);
        var values = group.Getter(item);
        if (values == null)
        {
            if (group.Setter == null)
                throw new BingOfficesConfigurationException(
                    $"实体动态列属性为空且不可写: {group.Property.DeclaringType?.FullName}.{group.Property.Name}",
                    stage: BingOfficesStage.Plan);
            values = CreateDynamicDictionary(item, group.Property, group.Setter);
        }
        values[key] = value;
    }

    /// <summary>
    /// 创建并写入实体动态列字典。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <param name="property">动态字典属性。</param>
    /// <param name="setter">写入属性的委托。</param>
    /// <returns>已创建的动态字典。</returns>
    private static IDictionary<string, object> CreateDynamicDictionary(object item, PropertyInfo property,
        Action<object, object> setter)
    {
        var name = $"{property.DeclaringType?.FullName}.{property.Name}";
        if (setter == null)
            throw new BingOfficesConfigurationException($"实体动态列属性为空且不可写: {name}",
                stage: BingOfficesStage.Plan);
        try
        {
            if (!typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
                throw new BingOfficesConfigurationException($"实体动态列属性类型不受支持: {name}",
                    stage: BingOfficesStage.Plan);
            object instance;
            if (property.PropertyType.IsInterface || property.PropertyType.IsAbstract)
            {
                if (!property.PropertyType.IsAssignableFrom(typeof(Dictionary<string, object>)))
                    throw new BingOfficesConfigurationException($"实体动态列属性无法创建字典实例: {name}",
                        stage: BingOfficesStage.Plan);
                instance = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            }
            else
                instance = Activator.CreateInstance(property.PropertyType);
            if (instance is not IDictionary<string, object> created)
                throw new BingOfficesConfigurationException($"实体动态列属性无法创建字典实例: {name}",
                    stage: BingOfficesStage.Plan);
            setter(item, created);
            return created;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException($"实体动态列属性无法创建字典实例: {name}",
                exception, BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 按显式动态列组策略校验实体中未声明的动态键。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="region">列表区域声明。</param>
    /// <param name="plan">列表区域映射计划。</param>
    /// <param name="item">当前列表项。</param>
    /// <param name="combined">当前列表项的动态值快照。</param>
    private static void ValidateUnknownDynamicValues<TEntity>(ExcelEntityListRegion<TEntity> region,
        EntityRegionPlan plan, object item, IDictionary<string, object> combined)
        where TEntity : class, new()
    {
        if (region.DynamicColumnGroups.Count == 0 || region.UnknownDynamicValues != ExcelUnknownDynamicValuePolicy.Fail)
            return;
        var keys = new HashSet<string>(plan.Columns.Where(column => column.Dynamic != null)
            .Select(column => column.Key), StringComparer.OrdinalIgnoreCase);
        var unknown = combined.Keys.Where(key => !keys.Contains(key)).ToArray();
        if (unknown.Length > 0)
            throw new BingOfficesConfigurationException(
                $"实体动态列包含未定义 Key: {string.Join(", ", unknown)}", stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 获取尾部从标记前间隔到相对内容末行的占用行数。
    /// </summary>
    /// <param name="footer">尾部定义。</param>
    /// <returns>尾部占用的行数。</returns>
    private static int GetFooterHeight(IExcelEntityListFooter footer)
    {
        if (footer == null)
            return 0;
        var lastRelativeRow = 0;
        foreach (var cell in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
            lastRelativeRow = Math.Max(lastRelativeRow, cell.Reference.Row);
        foreach (var merge in footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
            lastRelativeRow = Math.Max(lastRelativeRow, merge.Last.Row);
        return lastRelativeRow + 1;
    }

    /// <summary>
    /// 按相邻分组键计算列表元素的连续区间。
    /// </summary>
    /// <param name="data">列表元素集合。</param>
    /// <param name="subtotal">分组小计定义。</param>
    /// <returns>连续分组的起始索引和元素数量。</returns>
    private static IReadOnlyList<(int Start, int Count)> GetGroupedRanges(
        IReadOnlyList<object> data, IExcelEntityGroupSubtotal subtotal)
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
    /// 将列表尾部标记、聚合值和相对合并区域写入 ClosedXML 工作表。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="footer">列表尾部定义。</param>
    /// <param name="items">已物化的列表项快照。</param>
    /// <param name="dataEndRow">明细结束后的零基行索引。</param>
    /// <param name="startColumn">列表区域起始零基列索引。</param>
    /// <param name="template">是否按模板模式写入。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="anchorName">可选的最终尾部命名锚点。</param>
    /// <param name="detailRows">最终尾部使用的一基明细行段。</param>
    private static void WriteFooter(IXLWorksheet sheet, IExcelEntityListFooter footer,
        IReadOnlyList<object> items, int dataEndRow, int startColumn, bool template,
        CancellationToken cancellationToken, string anchorName = null,
        IReadOnlyList<(int FirstRow, int LastRow)> detailRows = null)
    {
        if (footer == null)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        var footerRow = dataEndRow + footer.GapRows;
        var marker = sheet.Cell(footerRow + 1, startColumn + 1);
        WriteValue(marker, footer.MarkerText);
        ApplyFooterStyle(marker, footer.MarkerStyle, null);
        foreach (var cellDefinition in footer.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = sheet.Cell(footerRow + cellDefinition.Reference.Row + 1,
                startColumn + cellDefinition.Reference.Column + 1);
            var value = cellDefinition.Evaluate(items);
            var formulaA1 = (value as ExcelEntityFooterFormulaValue)?.FormulaA1;
            if (value is ExcelEntityFooterDetailSumValue detailSum)
            {
                var rows = detailRows ?? (items.Count == 0
                    ? Array.Empty<(int FirstRow, int LastRow)>()
                    : new[] { (dataEndRow - items.Count + 1, dataEndRow) });
                formulaA1 = detailSum.ToFormulaA1(rows);
            }
            if (formulaA1 != null)
                cell.FormulaA1 = formulaA1;
            else
                WriteValue(cell, value);
            ApplyFooterStyle(cell, cellDefinition.Style, cellDefinition.NumberFormat);
        }
        foreach (var merge in footer.Merges)
        {
            var absolute = ExcelEntityCellRange.Parse(
                ToAddress(footerRow + merge.First.Row, startColumn + merge.First.Column) + ":" +
                ToAddress(footerRow + merge.Last.Row, startColumn + merge.Last.Column));
            EnsureMerge(sheet, absolute, addMissing: true, requireExisting: false);
        }
        if (anchorName != null)
            UpdateFooterAnchor(sheet, anchorName, footerRow, startColumn);
    }

    /// <summary>
    /// 将工作簿名称定位到最终尾部标记单元格。
    /// </summary>
    /// <param name="sheet">尾部所在工作表。</param>
    /// <param name="anchorName">命名锚点名称。</param>
    /// <param name="row">标记的零基行索引。</param>
    /// <param name="column">标记的零基列索引。</param>
    private static void UpdateFooterAnchor(IXLWorksheet sheet, string anchorName, int row, int column)
    {
        var workbook = sheet.Workbook;
        var global = workbook.DefinedNames.Contains(anchorName);
        var scoped = workbook.Worksheets.Where(item => item.DefinedNames.Contains(anchorName)).ToArray();
        if ((global ? 1 : 0) + scoped.Length > 1 || scoped.Any(item => item != sheet))
            throw new BingOfficesConfigurationException(
                $"最终尾部命名锚点不唯一或属于其他 Sheet: {sheet.Name}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (global || scoped.Length == 1)
            ResolveNamedAnchor(sheet, anchorName);
        var target = ToAddress(row, column);
        var columnLength = target.TakeWhile(char.IsLetter).Count();
        var escapedSheet = sheet.Name.Replace("'", "''");
        var formula = $"'{escapedSheet}'!${target.Substring(0, columnLength)}${target.Substring(columnLength)}";
        if (global)
            workbook.DefinedNames.DefinedName(anchorName).RefersTo = formula;
        else if (scoped.Length == 1)
            sheet.DefinedNames.DefinedName(anchorName).RefersTo = formula;
        else
            workbook.DefinedNames.Add(anchorName, sheet.Range(row + 1, column + 1, row + 1, column + 1));
    }

    /// <summary>
    /// 在写入列表明细前校验尾部范围并准备合并区域。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="footer">列表尾部定义。</param>
    /// <param name="dataEndRow">明细结束后的零基行索引。</param>
    /// <param name="startColumn">列表区域起始零基列索引。</param>
    private static void PrepareFooterMerges(IXLWorksheet sheet, IExcelEntityListFooter footer,
        int dataEndRow, int startColumn)
    {
        if (footer == null)
            return;
        var footerRow = dataEndRow + footer.GapRows;
        foreach (var merge in footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
        {
            var absolute = ExcelEntityCellRange.Parse(
                ToAddress(footerRow + merge.First.Row, startColumn + merge.First.Column) + ":" +
                ToAddress(footerRow + merge.Last.Row, startColumn + merge.Last.Column));
            EnsureMerge(sheet, absolute, addMissing: true, requireExisting: false);
        }
    }

    /// <summary>
    /// 校验列表尾部单元格和合并区域不超出列表声明边界。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="region">列表区域声明。</param>
    /// <param name="footer">列表尾部定义。</param>
    /// <param name="dataEndRow">明细结束后的零基行索引。</param>
    private static void ValidateFooterBounds<TEntity>(ExcelEntityListRegion<TEntity> region,
        IExcelEntityListFooter footer, int dataEndRow) where TEntity : class, new()
    {
        if (footer == null)
            return;
        var footerRow = (long)dataEndRow + footer.GapRows;
        var lastRow = footerRow;
        var lastColumn = (long)region.Start.Column;
        foreach (var cell in footer.Cells ?? Array.Empty<IExcelEntityListFooterCell>())
        {
            lastRow = Math.Max(lastRow, footerRow + cell.Reference.Row);
            lastColumn = Math.Max(lastColumn, (long)region.Start.Column + cell.Reference.Column);
        }
        foreach (var merge in footer.Merges ?? Array.Empty<ExcelEntityCellRange>())
        {
            lastRow = Math.Max(lastRow, footerRow + merge.Last.Row);
            lastColumn = Math.Max(lastColumn, (long)region.Start.Column + merge.Last.Column);
        }
        if (region.End.HasValue && (lastRow > region.End.Value.Row || lastColumn > region.End.Value.Column))
            throw new BingOfficesConfigurationException(
                $"列表尾部超出声明边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (lastRow > 1_048_575L || lastColumn > 16_383L)
            throw new BingOfficesConfigurationException(
                $"列表尾部超出 ClosedXML 工作表物理边界: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 将尾部样式和显式数字格式应用到单元格。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="style">可选的尾部样式。</param>
    /// <param name="numberFormat">显式数字格式；为空时沿用样式格式。</param>
    private static void ApplyFooterStyle(IXLCell cell, ExcelCellStyle style, string numberFormat)
    {
        ClosedXmlStyleAdapter.Apply(cell, style);
        ApplyNumberFormat(cell, numberFormat ?? style?.NumberFormat);
    }

    /// <summary>
    /// 在单元格上应用非空数字格式。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="numberFormat">数字格式文本。</param>
    private static void ApplyNumberFormat(IXLCell cell, string numberFormat)
    {
        if (!string.IsNullOrWhiteSpace(numberFormat))
            cell.Style.NumberFormat.Format = numberFormat;
    }

    /// <summary>
    /// 查找并校验列表尾部的唯一 marker 行。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="sheet">来源工作表。</param>
    /// <param name="region">列表区域声明。</param>
    /// <param name="firstDataRow">明细首行的零基行索引。</param>
    /// <returns>marker 行的零基索引；未配置尾部时返回 <see langword="null" />。</returns>
    private static int? FindFooterMarkerRow<TEntity>(IXLWorksheet sheet,
        ExcelEntityListRegion<TEntity> region, int firstDataRow) where TEntity : class, new()
    {
        if (region.Footer == null)
            return null;
        var last = region.End?.Row ?? sheet.LastRowUsed()?.RowNumber() - 1 ?? firstDataRow;
        var matches = new List<int>();
        for (var row = firstDataRow; row <= last; row++)
        {
            var value = ReadRaw(sheet.Cell(row + 1, region.Start.Column + 1), typeof(string));
            if (string.Equals(Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim(),
                    region.Footer.MarkerText, StringComparison.Ordinal))
                matches.Add(row);
        }
        if (matches.Count == 0)
            throw new BingOfficesConfigurationException(
                $"列表区域缺少尾部标记: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (matches.Count > 1)
            throw new BingOfficesConfigurationException(
                $"列表区域尾部标记重复: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        if (matches[0] < firstDataRow + (region.Footer?.GapRows ?? 0))
            throw new BingOfficesConfigurationException(
                $"列表区域尾部标记位于明细起始位置之前: {region.SheetName}!{region.Start.Address}",
                stage: BingOfficesStage.Plan);
        return matches[0];
    }

    /// <summary>
    /// 查找实体列表最终尾部和连续分组小计需要跳过的工作表行段。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="sheet">来源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="firstDataRow">明细首行的零基行索引。</param>
    /// <returns>按起始行排序且不重叠的跳过行段。</returns>
    private static IReadOnlyList<(int StartRow, int EndRow)> FindFooterSkipSpans<TEntity>(
        IXLWorksheet sheet, ExcelEntityListRegion<TEntity> region, int firstDataRow)
        where TEntity : class, new()
    {
        var lastRow = region.End?.Row ?? sheet.LastRowUsed()?.RowNumber() - 1 ?? firstDataRow;
        var spans = new List<(int StartRow, int EndRow)>();
        var pageSubtotal = region.PageSubtotal;
        if (pageSubtotal != null && region.PageBreakRows is int rowsPerPage)
        {
            var matches = FindMarkerRows(sheet, region, pageSubtotal.MarkerText, firstDataRow, lastRow);
            var cursor = firstDataRow;
            foreach (var markerRow in matches)
            {
                var expected = cursor + rowsPerPage + pageSubtotal.GapRows;
                if (markerRow != expected)
                    throw new BingOfficesConfigurationException(
                        $"列表区域分页小计标记位置无效: {region.SheetName}!{region.Start.Address}",
                        stage: BingOfficesStage.Plan);
                spans.Add((markerRow - pageSubtotal.GapRows,
                    markerRow + GetFooterHeight(pageSubtotal) - 1));
                cursor = markerRow + GetFooterHeight(pageSubtotal);
            }
            var finalMarker = region.Footer == null
                ? -1
                : FindFooterMarkerRow(sheet, region, firstDataRow).GetValueOrDefault(-1);
            var detailEnd = region.Footer == null
                ? lastRow + 1
                : finalMarker - region.Footer.GapRows;
            if (matches.Count > 0 && detailEnd <= cursor)
                throw new BingOfficesConfigurationException(
                    $"列表区域分页小计之后缺少明细: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
            if (detailEnd - cursor > rowsPerPage)
                throw new BingOfficesConfigurationException(
                    $"列表区域缺少分页小计标记: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
        }
        var group = region.GroupSubtotal;
        if (group != null)
        {
            for (var row = firstDataRow; row <= lastRow; row++)
            {
                var value = ReadRaw(sheet.Cell(row + 1, region.Start.Column + 1), typeof(string));
                if (!string.Equals(Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim(),
                        group.Footer.MarkerText, StringComparison.Ordinal))
                    continue;
                if (row < firstDataRow + group.Footer.GapRows)
                    throw new BingOfficesConfigurationException(
                        $"列表区域分组小计标记位于明细起始位置之前: {region.SheetName}!{region.Start.Address}",
                        stage: BingOfficesStage.Plan);
                spans.Add((row - group.Footer.GapRows,
                    row + GetFooterHeight(group.Footer) - 1));
            }
        }
        if (region.Footer != null)
        {
            var marker = FindFooterMarkerRow(sheet, region, firstDataRow);
            spans.Add((marker.Value - region.Footer.GapRows,
                marker.Value + GetFooterHeight(region.Footer) - 1));
        }
        var ordered = spans.OrderBy(span => span.StartRow).ThenBy(span => span.EndRow).ToArray();
        for (var index = 1; index < ordered.Length; index++)
        {
            if (ordered[index].StartRow <= ordered[index - 1].EndRow)
                throw new BingOfficesConfigurationException(
                    $"列表区域尾部与分组小计行段重叠: {region.SheetName}!{region.Start.Address}",
                    stage: BingOfficesStage.Plan);
        }
        return ordered;
    }

    /// <summary>
    /// 查找指定范围内与尾部标记精确匹配的行。
    /// </summary>
    /// <typeparam name="TEntity">根实体类型。</typeparam>
    /// <param name="sheet">源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="markerText">已规范化的尾部标记。</param>
    /// <param name="firstRow">起始行的零基索引。</param>
    /// <param name="lastRow">结束行的零基索引。</param>
    /// <returns>按工作表顺序排列的匹配行索引。</returns>
    private static IReadOnlyList<int> FindMarkerRows<TEntity>(IXLWorksheet sheet,
        ExcelEntityListRegion<TEntity> region, string markerText, int firstRow, int lastRow)
        where TEntity : class, new()
    {
        if (string.IsNullOrWhiteSpace(markerText))
            return Array.Empty<int>();
        var matches = new List<int>();
        for (var row = firstRow; row <= lastRow; row++)
        {
            var value = ReadRaw(sheet.Cell(row + 1, region.Start.Column + 1), typeof(string));
            if (string.Equals(Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim(),
                    markerText, StringComparison.Ordinal))
                matches.Add(row);
        }
        return matches;
    }

    /// <summary>
    /// 将零基行列索引转换为 A1 地址。
    /// </summary>
    /// <param name="row">零基行号。</param>
    /// <param name="column">零基列号。</param>
    /// <returns>A1 格式地址。</returns>
    private static string ToAddress(int row, int column)
    {
        var value = column + 1;
        var letters = string.Empty;
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            value = (value - 1) / 26;
        }
        return letters + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 从 ClosedXML 单元格读取公共原始值，并优先使用公式缓存值。
    /// </summary>
    /// <param name="cell">来源单元格。</param>
    /// <param name="targetType">目标属性类型。</param>
    /// <returns>单元格原始值；空单元格返回 <see langword="null" />。</returns>
    private static object ReadRaw(IXLCell cell, Type targetType)
    {
        var hasFormula = !string.IsNullOrWhiteSpace(cell.FormulaA1);
        if (hasFormula && targetType == typeof(string))
            return "=" + cell.FormulaA1;
        var value = hasFormula ? cell.CachedValue : cell.Value;
        if (value.IsBlank)
            return null;
        if (value.IsBoolean)
            return value.GetBoolean();
        if (value.IsNumber)
            return value.GetNumber();
        if (value.IsDateTime)
            return value.GetDateTime();
        if (value.IsTimeSpan)
            return value.GetTimeSpan();
        if (value.IsError)
            return value.GetError().ToString();
        return value.GetText();
    }

    /// <summary>
    /// 将公共值或公式文本写入 ClosedXML 单元格。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="value">待写入的值。</param>
    private static void WriteValue(IXLCell cell, object value)
    {
        if (value is string formula && formula.StartsWith("=", StringComparison.Ordinal))
        {
            cell.FormulaA1 = formula.Substring(1);
            return;
        }
        switch (value)
        {
            case null: cell.Clear(XLClearOptions.Contents); break;
            case string text: cell.Value = text; break;
            case bool boolean: cell.Value = boolean; break;
            case DateTime date: cell.Value = date; break;
            case DateTimeOffset offset: cell.Value = offset.DateTime; break;
            case byte number: cell.Value = number; break;
            case sbyte number: cell.Value = number; break;
            case short number: cell.Value = number; break;
            case ushort number: cell.Value = number; break;
            case int number: cell.Value = number; break;
            case uint number: cell.Value = number; break;
            case long number: cell.Value = number; break;
            case ulong number: cell.Value = number; break;
            case float number: cell.Value = number; break;
            case double number: cell.Value = number; break;
            case decimal number: cell.Value = number; break;
            default: cell.Value = Convert.ToString(value, CultureInfo.InvariantCulture); break;
        }
    }

    /// <summary>
    /// 应用计算列声明的正文或表头样式。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="column">当前物理列。</param>
    /// <param name="header">是否为表头单元格。</param>
    private static void ApplyCalculatedStyle(IXLCell cell, ClosedXmlExportColumn column, bool header)
    {
        if (column?.Calculated == null)
            return;
        ClosedXmlStyleAdapter.Apply(cell, header ? column.HeaderStyle : column.BodyStyle);
        ApplyNumberFormat(cell, column.NumberFormat
            ?? (header ? column.HeaderStyle?.NumberFormat : column.BodyStyle?.NumberFormat));
    }

    /// <summary>
    /// 执行实体导出列的公共校验。
    /// </summary>
    /// <param name="column">实体列。</param>
    /// <param name="raw">属性原始值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="sheet">工作表名称。</param>
    /// <param name="reference">目标单元格引用。</param>
    private static void ValidateExport(EntityColumn column, object raw, object converted, string sheet,
        ExcelEntityCellReference reference)
    {
        var text = ClosedXmlValueAdapter.ToText(converted, CultureInfo.InvariantCulture);
        var context = new ExcelValidationContext(text, sheet, reference.Row + 1, reference.Column + 1,
            column.Property.Name, raw, column.Property.PropertyType,
            ClosedXmlValueAdapter.CreateCell(converted, text), CultureInfo.InvariantCulture);
        foreach (var binding in column.Column.ValidationBindings ??
            (IReadOnlyList<IExcelValidationBinding>)Array.Empty<IExcelValidationBinding>())
        {
            try
            {
                if (!binding.Validate(context))
                    throw new BingOfficesExportException(binding.ErrorMessage ?? "实体属性未通过校验。",
                        provider: Provider, stage: BingOfficesStage.Validate, sheetName: sheet,
                        rowIndex: reference.Row + 1, columnIndex: reference.Column + 1,
                        propertyName: column.Property.Name);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesExportException("实体校验器执行失败。", exception, Provider,
                    BingOfficesStage.Validate, sheet, reference.Row + 1, reference.Column + 1,
                    column.Property.Name, BingOfficesErrorCode.UserExtensionFailed);
            }
        }
    }

    /// <summary>
    /// 执行实体导入列的公共校验。
    /// </summary>
    /// <param name="column">实体列。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="cell">公共单元格值。</param>
    /// <param name="sheet">工作表名称。</param>
    /// <param name="reference">来源单元格引用。</param>
    /// <param name="errors">错误集合。</param>
    /// <param name="rawOnly">是否只执行原始值校验。</param>
    /// <returns>校验通过时返回 <see langword="true" />。</returns>
    private static bool ValidateImport(EntityColumn column, object raw, object converted, string text,
        ExcelCellValue cell, string sheet, ExcelEntityCellReference reference,
        ICollection<ExcelImportError> errors, bool rawOnly)
    {
        var bindings = column.Column.ValidationBindings ??
            (IReadOnlyList<IExcelValidationBinding>)Array.Empty<IExcelValidationBinding>();
        foreach (var binding in bindings.Where(item => (rawOnly ? item.IsRaw : !item.IsRaw)
            && item.Kind != ExcelValidationBindingKind.Unique))
        {
            var value = rawOnly ? raw : converted;
            var context = new ExcelValidationContext(text, sheet, reference.Row + 1, reference.Column + 1,
                column.Property.Name, value, column.Property.PropertyType, cell, CultureInfo.InvariantCulture);
            try
            {
                if (binding.Validate(context))
                    continue;
                errors.Add(new ExcelImportError(ClosedXmlValueAdapter.GetValidationCode(binding),
                    binding.ErrorMessage, sheet, reference.Row + 1, reference.Column + 1,
                    column.Property.Name, rawValue: raw));
                return false;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message, sheet,
                    reference.Row + 1, reference.Column + 1, column.Property.Name, rawValue: raw));
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// 执行实体列表列的跨行唯一性校验。
    /// </summary>
    /// <param name="key">唯一值所属的列键。</param>
    /// <param name="isUnique">是否启用唯一性校验。</param>
    /// <param name="ignoreEmpty">是否忽略空值。</param>
    /// <param name="text">单元格规范化文本。</param>
    /// <param name="sheet">工作表名称。</param>
    /// <param name="reference">来源单元格引用。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="unique">当前列表区域的唯一值跟踪器。</param>
    /// <returns>唯一性校验通过时返回 <see langword="true" />。</returns>
    private static bool ValidateUniqueImport(string key, bool isUnique, bool ignoreEmpty, string text,
        string sheet, ExcelEntityCellReference reference, ICollection<ExcelImportError> errors,
        UniqueTracker unique)
    {
        if (!isUnique || (ignoreEmpty && string.IsNullOrWhiteSpace(text)))
            return true;
        try
        {
            if (unique.TryReserve(key, text, false, ignoreEmpty, reference.Row + 1))
                return true;
            errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, "重复数据。", sheet,
                reference.Row + 1, reference.Column + 1, key, rawValue: text));
            return false;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit, exception.Message, sheet,
                reference.Row + 1, reference.Column + 1, key, rawValue: text));
            return false;
        }
    }

    /// <summary>
    /// 将公共字符串比较选项转换为 .NET 比较器。
    /// </summary>
    /// <param name="comparison">字符串比较选项。</param>
    /// <returns>对应的字符串比较器。</returns>
    private static StringComparer CreateStringComparer(StringComparison comparison) => comparison switch
    {
        StringComparison.Ordinal => StringComparer.Ordinal,
        StringComparison.OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase,
        StringComparison.InvariantCulture => StringComparer.InvariantCulture,
        StringComparison.InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase,
        StringComparison.CurrentCulture => StringComparer.CurrentCulture,
        StringComparison.CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase,
        _ => throw new ArgumentOutOfRangeException(nameof(comparison))
    };

    /// <summary>
    /// 执行实体动态列的导出校验。
    /// </summary>
    /// <param name="column">动态列映射。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="sheet">当前工作表名称。</param>
    /// <param name="reference">单元格位置。</param>
    private static void ValidateDynamicExport(IExcelDynamicMappingColumn column, object raw,
        object converted, string sheet, ExcelEntityCellReference reference)
    {
        var propertyType = ClosedXmlValueAdapter.ResolveDynamicType(column.DataTypeName);
        var text = ClosedXmlValueAdapter.ToText(converted, CultureInfo.InvariantCulture);
        foreach (var binding in column.ValidationBindings ??
            (IReadOnlyList<IExcelValidationBinding>)Array.Empty<IExcelValidationBinding>())
        {
            try
            {
                var value = binding.IsRaw ? raw : converted;
                var context = new ExcelValidationContext(text, sheet, reference.Row + 1,
                    reference.Column + 1, column.Key, value, propertyType,
                    ClosedXmlValueAdapter.CreateCell(converted, text),
                    CultureInfo.InvariantCulture);
                if (binding.Validate(context))
                    continue;
                throw new BingOfficesExportException(binding.ErrorMessage ?? "实体动态列未通过校验。",
                    provider: Provider, stage: BingOfficesStage.Validate, sheetName: sheet,
                    rowIndex: reference.Row + 1, columnIndex: reference.Column + 1,
                    propertyName: column.Key);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesExportException("实体动态列校验器执行失败。", exception,
                    Provider, BingOfficesStage.Validate, sheet, reference.Row + 1,
                    reference.Column + 1, column.Key, BingOfficesErrorCode.UserExtensionFailed);
            }
        }
    }

    /// <summary>
    /// 执行实体动态列的导入校验。
    /// </summary>
    /// <param name="column">动态列映射。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="cell">待处理的单元格。</param>
    /// <param name="sheet">当前工作表名称。</param>
    /// <param name="reference">单元格位置。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="rawOnly">为 true 时仅校验原始值；为 false 时仅校验转换后的值。</param>
    /// <returns>所有适用规则均通过时返回 true；校验失败时返回 false，并记录错误。</returns>
    private static bool ValidateDynamicImport(IExcelDynamicMappingColumn column, object raw,
        object converted, string text, ExcelCellValue cell, string sheet,
        ExcelEntityCellReference reference, ICollection<ExcelImportError> errors, bool rawOnly)
    {
        var propertyType = ClosedXmlValueAdapter.ResolveDynamicType(column.DataTypeName);
        var bindings = column.ValidationBindings ??
            (IReadOnlyList<IExcelValidationBinding>)Array.Empty<IExcelValidationBinding>();
        foreach (var binding in bindings.Where(item => (rawOnly ? item.IsRaw : !item.IsRaw)
            && item.Kind != ExcelValidationBindingKind.Unique))
        {
            var value = rawOnly ? raw : converted;
            var context = new ExcelValidationContext(text, sheet, reference.Row + 1,
                reference.Column + 1, column.Key, value, propertyType, cell,
                CultureInfo.InvariantCulture);
            try
            {
                if (binding.Validate(context))
                    continue;
                errors.Add(new ExcelImportError(ClosedXmlValueAdapter.GetValidationCode(binding),
                    binding.ErrorMessage, sheet, reference.Row + 1, reference.Column + 1,
                    column.Key, rawValue: raw));
                return false;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, exception.Message,
                    sheet, reference.Row + 1, reference.Column + 1, column.Key, rawValue: raw));
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// 在 ClosedXML 实体导入路径中共享并限制结构化错误集合。
    /// </summary>
    private sealed class LimitedErrorCollection : ICollection<ExcelImportError>
    {
        /// <summary>
        /// 当前导入允许保留的最大错误数；未指定时不限制。
        /// </summary>
        private readonly int? _maxErrors;
        /// <summary>
        /// 当前导入已收集的结构化错误。
        /// </summary>
        private readonly List<ExcelImportError> _items = new List<ExcelImportError>();

        /// <summary>
        /// 初始化一个 <see cref="LimitedErrorCollection"/> 类型的实例。
        /// </summary>
        /// <param name="maxErrors">允许保留的最大错误数。</param>
        internal LimitedErrorCollection(int? maxErrors)
        {
            _maxErrors = maxErrors;
        }

        /// <summary>
        /// 获取当前已收集的错误列表。
        /// </summary>
        internal IReadOnlyList<ExcelImportError> Items => _items;

        /// <summary>
        /// 获取是否因达到数量上限而截断。
        /// </summary>
        internal bool IsTruncated { get; private set; }

        /// <inheritdoc />
        public int Count => _items.Count;
        /// <inheritdoc />
        public bool IsReadOnly => false;

        /// <inheritdoc />
        /// <remarks>忽略 null；达到错误数量上限后不再追加，并标记截断状态。</remarks>
        public void Add(ExcelImportError item)
        {
            if (item == null)
                return;
            if (_maxErrors.HasValue && _items.Count >= _maxErrors.Value)
            {
                IsTruncated = true;
                return;
            }
            _items.Add(item);
            if (_maxErrors.HasValue && _items.Count >= _maxErrors.Value)
                IsTruncated = true;
        }

        /// <inheritdoc />
        public void Clear() => _items.Clear();
        /// <inheritdoc />
        public bool Contains(ExcelImportError item) => _items.Contains(item);
        /// <inheritdoc />
        public void CopyTo(ExcelImportError[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);
        /// <inheritdoc />
        public bool Remove(ExcelImportError item) => _items.Remove(item);
        /// <inheritdoc />
        public IEnumerator<ExcelImportError> GetEnumerator() => _items.GetEnumerator();
        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// 保存实体列映射与对应属性。
    /// </summary>
    /// <param name="Column">实体映射列。</param>
    /// <param name="Property">实体属性。</param>
    private sealed record EntityColumn(IExcelMappingColumn Column, PropertyInfo Property);

    /// <summary>
    /// 保存实体列表区域的最终物理列布局和动态字典属性。
    /// </summary>
    /// <param name="Columns">按布局生成的固定列和动态列。</param>
    /// <param name="DynamicProperty">单字典动态列属性。</param>
    /// <param name="DynamicGroups">显式动态列组。</param>
    /// <param name="Width">从区域起点计算的物理列宽度。</param>
    private sealed record EntityRegionPlan(IReadOnlyList<ClosedXmlExportColumn> Columns,
        PropertyInfo DynamicProperty, IReadOnlyList<IExcelEntityDynamicColumnGroup> DynamicGroups, int Width);
}
