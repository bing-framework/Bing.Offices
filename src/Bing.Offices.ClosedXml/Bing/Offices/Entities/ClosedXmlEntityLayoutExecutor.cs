using System.Collections;
using System.Globalization;
using System.Reflection;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Entities;

/// <summary>
/// 执行 ClosedXML 的固定单元格、列表区域和合并布局。
/// </summary>
internal sealed class ClosedXmlEntityLayoutExecutor
{
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
            ResolveSheet(workbook, cell.SheetName, template: true);
        foreach (var region in layout.ListRegions)
            ResolveSheet(workbook, region.SheetName, template: true);
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
            var column = CreateColumn(typeof(TEntity), binding.PropertyName, binding.ConverterName,
                binding.MappingDocument, binding.MappingConfiguration, MappingDirection.Export);
            var cell = sheet.Cell(binding.Reference.Row + 1, binding.Reference.Column + 1);
            var raw = binding.Getter(entity);
            var converted = ClosedXmlValueAdapter.ConvertTo(raw, column.Column, column.Property,
                binding.SheetName, binding.Reference.Row + 1, binding.Reference.Column + 1,
                CultureInfo.InvariantCulture);
            ValidateExport(column, raw, converted, binding.SheetName, binding.Reference);
            WriteValue(cell, converted);
        }
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            WriteRegion(workbook, entity, region, template, cancellationToken);
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
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>实体、列表区域结果和结构化导入错误。</returns>
    public ExcelEntityImportResult<TEntity> Read<TEntity>(XLWorkbook workbook, ExcelEntityLayout<TEntity> layout,
        bool requireTemplateMerges, bool isDate1904, ExcelResourceLimits limits,
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
            var column = CreateColumn(typeof(TEntity), binding.PropertyName, binding.ConverterName,
                binding.MappingDocument, binding.MappingConfiguration, MappingDirection.Import);
            var cell = sheet.Cell(binding.Reference.Row + 1, binding.Reference.Column + 1);
            var raw = ReadRaw(cell, binding.PropertyType);
            var text = ClosedXmlValueAdapter.ToText(raw, CultureInfo.InvariantCulture);
            var cellValue = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
            if (!ValidateImport(column, raw, null, text, cellValue, binding.SheetName,
                    binding.Reference, errors, rawOnly: true))
                continue;
            try
            {
                var converted = ClosedXmlValueAdapter.ConvertFrom(raw, column.Column, column.Property,
                    binding.SheetName, binding.Reference.Row + 1, binding.Reference.Column + 1,
                    CultureInfo.InvariantCulture, isDate1904, cellValue);
                if (!ValidateImport(column, raw, converted, text, cellValue, binding.SheetName,
                        binding.Reference, errors, rawOnly: false))
                    continue;
                binding.Setter?.Invoke(entity, converted);
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
            sheetResults.Add(ReadRegion(workbook, entity, region, isDate1904, errors, limits,
                ref totalRows, cancellationToken));
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
    /// <param name="cancellationToken">取消令牌。</param>
    private void WriteRegion<TEntity>(XLWorkbook workbook, TEntity entity,
        ExcelEntityListRegion<TEntity> region, bool template, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var sheet = ResolveSheet(workbook, region.SheetName, template);
        var values = (region.Getter(entity) as IEnumerable)?.Cast<object>().ToArray()
            ?? Array.Empty<object>();
        var plan = CreateRegionPlan(region.ItemType, region.MappingDocument, region.MappingConfiguration,
            MappingDirection.Export);
        ValidateBounds(region, plan.Width, values.Length);
        var row = region.Start.Row;
        if (region.IncludeHeader)
        {
            foreach (var column in plan.Columns)
                sheet.Cell(row + 1, region.Start.Column + column.PhysicalColumnIndex + 1).Value = column.Title;
            row++;
        }
        foreach (var item in values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item == null)
                throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                    stage: BingOfficesStage.Plan);
            var dynamicValues = GetDynamicValues(item, plan.DynamicProperty, create: false);
            foreach (var column in plan.Columns)
            {
                var columnIndex = region.Start.Column + column.PhysicalColumnIndex;
                object raw = null;
                object converted;
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
            }
            row++;
        }
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
    /// <param name="totalRows">跨列表区域累计读取的数据行数。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>列表区域导入结果。</returns>
    private ExcelSheetImportResult ReadRegion<TEntity>(XLWorkbook workbook, TEntity entity,
        ExcelEntityListRegion<TEntity> region, bool isDate1904, ICollection<ExcelImportError> errors,
        ExcelResourceLimits limits, ref int totalRows, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var sheet = ResolveSheet(workbook, region.SheetName, template: true);
        var plan = CreateRegionPlan(region.ItemType, region.MappingDocument, region.MappingConfiguration,
            MappingDirection.Import);
        var maxItems = GetRegionItemCapacity(region, plan.Width, sheet);
        ValidateBounds(region, plan.Width, maxItems);
        var rows = new List<int>();
        var items = new List<object>();
        var row = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var lastRow = region.End?.Row ?? sheet.LastRowUsed()?.RowNumber() - 1 ?? row - 1;
        while (row <= lastRow)
        {
            cancellationToken.ThrowIfCancellationRequested();
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
            var dynamicValues = GetDynamicValues(item, plan.DynamicProperty, create: true);
            var valid = true;
            foreach (var column in plan.Columns)
            {
                var columnIndex = region.Start.Column + column.PhysicalColumnIndex;
                var reference = ExcelEntityCellReference.Parse(ToAddress(row, columnIndex));
                var cell = sheet.Cell(row + 1, columnIndex + 1);
                var raw = ReadRaw(cell, column.Property?.PropertyType ?? typeof(object));
                var text = ClosedXmlValueAdapter.ToText(raw, CultureInfo.InvariantCulture);
                var cellValue = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
                if (column.Dynamic != null && !ValidateDynamicImport(column.Dynamic, raw, null, text,
                        cellValue, region.SheetName, reference, errors, rawOnly: true))
                {
                    valid = false;
                    continue;
                }
                if (column.Fixed != null && !ValidateImport(new EntityColumn(column.Fixed, column.Property),
                        raw, null, text, cellValue, region.SheetName, reference, errors, rawOnly: true))
                {
                    valid = false;
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
                        continue;
                    }
                    if (column.Fixed != null && !ValidateImport(new EntityColumn(column.Fixed, column.Property),
                            raw, converted, text, cellValue, region.SheetName, reference, errors, rawOnly: false))
                    {
                        valid = false;
                        continue;
                    }
                    if (column.Dynamic != null)
                        dynamicValues[column.Key] = converted;
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
                }
            }
            if (valid)
            {
                items.Add(item);
                rows.Add(row);
            }
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
    /// <param name="itemType">列表项目类型。</param>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>包含最终物理索引和动态字典属性的区域计划。</returns>
    private EntityRegionPlan CreateRegionPlan(Type itemType, ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
    {
        var plan = _planBuilder.Create(itemType, document ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, configuration, direction);
        var columns = ClosedXmlExportColumnPlanner.Create(itemType, plan,
                Array.Empty<ExcelDynamicColumnDefinition>(), configuration)
            .Where(column => column.Dynamic != null || direction != MappingDirection.Import
                || column.Property.CanWrite)
            .ToArray();
        if (columns.Length == 0)
            throw new BingOfficesConfigurationException($"实体列表没有可用映射列: {itemType.FullName}",
                stage: BingOfficesStage.Plan);
        var dynamicPlanColumn = plan.DynamicColumns.Count == 0 ? null
            : plan.Columns.SingleOrDefault(column => column.IsDynamicColumn && !column.Ignored);
        var dynamicProperty = dynamicPlanColumn == null ? null : itemType.GetProperty(dynamicPlanColumn.Name,
            BindingFlags.Instance | BindingFlags.Public);
        if (plan.DynamicColumns.Count > 0 && (dynamicProperty == null || !dynamicProperty.CanRead
            || !typeof(IDictionary<string, object>).IsAssignableFrom(dynamicProperty.PropertyType)))
            throw new BingOfficesConfigurationException(
                "实体动态列属性必须是可读的 IDictionary<string, object>。", stage: BingOfficesStage.Plan);
        var width = columns.Max(column => column.PhysicalColumnIndex) + 1;
        return new EntityRegionPlan(columns, dynamicProperty, width);
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
    /// <param name="property">标记了动态列的字典属性。</param>
    /// <param name="create">属性为空时是否创建字典。</param>
    /// <returns>动态列字典；布局没有动态列时返回 <see langword="null" />。</returns>
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
        object instance = property.PropertyType.IsInterface || property.PropertyType.IsAbstract
            ? new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            : Activator.CreateInstance(property.PropertyType);
        if (instance is not IDictionary<string, object> created)
            throw new BingOfficesConfigurationException(
                $"实体动态列属性无法创建字典实例: {property.DeclaringType?.FullName}.{property.Name}",
                stage: BingOfficesStage.Plan);
        property.SetValue(item, created);
        return created;
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
            column.Property.Name, raw, column.Property.PropertyType, null, CultureInfo.InvariantCulture);
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
        foreach (var binding in bindings.Where(item => rawOnly ? item.IsRaw : !item.IsRaw))
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
                    reference.Column + 1, column.Key, value, propertyType, null,
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
        foreach (var binding in bindings.Where(item => rawOnly ? item.IsRaw : !item.IsRaw))
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
    /// <param name="DynamicProperty">动态列字典属性。</param>
    /// <param name="Width">从区域起点计算的物理列宽度。</param>
    private sealed record EntityRegionPlan(IReadOnlyList<ClosedXmlExportColumn> Columns,
        PropertyInfo DynamicProperty, int Width);
}
