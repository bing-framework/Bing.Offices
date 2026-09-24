using System.Collections;
using System.Globalization;
using System.Reflection;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
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
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>实体、列表区域结果和结构化导入错误。</returns>
    public ExcelEntityImportResult<TEntity> Read<TEntity>(XLWorkbook workbook, ExcelEntityLayout<TEntity> layout,
        bool requireTemplateMerges, bool isDate1904, CancellationToken cancellationToken)
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

        var entity = new TEntity();
        var errors = new List<ExcelImportError>();
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
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheetResults.Add(ReadRegion(workbook, entity, region, isDate1904, errors, cancellationToken));
        }
        ClosedXmlRelationCoordinator.Bind(entity, layout.Relations, errors, null, cancellationToken);
        return new ExcelEntityImportResult<TEntity>(entity, errors, sheetResults);
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
        var columns = CreateColumns(region.ItemType, region.MappingDocument, region.MappingConfiguration,
            MappingDirection.Export);
        ValidateBounds(region, columns.Count, values.Length);
        var row = region.Start.Row;
        if (region.IncludeHeader)
        {
            for (var index = 0; index < columns.Count; index++)
                sheet.Cell(row + 1, region.Start.Column + index + 1).Value = columns[index].Column.Title
                    ?? columns[index].Property.Name;
            row++;
        }
        foreach (var item in values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item == null)
                throw new BingOfficesConfigurationException("实体列表区域包含 null 项。",
                    stage: BingOfficesStage.Plan);
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var raw = column.Property.GetValue(item);
                var converted = ClosedXmlValueAdapter.ConvertTo(raw, column.Column, column.Property,
                    region.SheetName, row + 1, region.Start.Column + index + 1,
                    CultureInfo.InvariantCulture);
                ValidateExport(column, raw, converted, region.SheetName,
                    ExcelEntityCellReference.Parse(ToAddress(row, region.Start.Column + index)));
                WriteValue(sheet.Cell(row + 1, region.Start.Column + index + 1), converted);
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
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>列表区域导入结果。</returns>
    private ExcelSheetImportResult ReadRegion<TEntity>(XLWorkbook workbook, TEntity entity,
        ExcelEntityListRegion<TEntity> region, bool isDate1904, ICollection<ExcelImportError> errors,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var sheet = ResolveSheet(workbook, region.SheetName, template: true);
        var columns = CreateColumns(region.ItemType, region.MappingDocument, region.MappingConfiguration,
            MappingDirection.Import);
        var maxItems = GetRegionItemCapacity(region, columns.Count, sheet);
        ValidateBounds(region, columns.Count, maxItems);
        var rows = new List<int>();
        var items = new List<object>();
        var row = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var lastRow = region.End?.Row ?? sheet.LastRowUsed()?.RowNumber() - 1 ?? row - 1;
        while (row <= lastRow)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsEmptyRow(sheet, row, region.Start.Column, columns.Count))
                break;
            var item = Activator.CreateInstance(region.ItemType);
            var valid = true;
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                var reference = ExcelEntityCellReference.Parse(ToAddress(row, region.Start.Column + index));
                var cell = sheet.Cell(row + 1, region.Start.Column + index + 1);
                var raw = ReadRaw(cell, column.Property.PropertyType);
                var text = ClosedXmlValueAdapter.ToText(raw, CultureInfo.InvariantCulture);
                var cellValue = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
                if (!ValidateImport(column, raw, null, text, cellValue, region.SheetName,
                        reference, errors, rawOnly: true))
                {
                    valid = false;
                    continue;
                }
                try
                {
                    var converted = ClosedXmlValueAdapter.ConvertFrom(raw, column.Column, column.Property,
                        region.SheetName, row + 1, region.Start.Column + index + 1,
                        CultureInfo.InvariantCulture, isDate1904, cellValue);
                    if (!ValidateImport(column, raw, converted, text, cellValue, region.SheetName,
                            reference, errors, rawOnly: false))
                    {
                        valid = false;
                        continue;
                    }
                    column.Property.SetValue(item, converted);
                }
                catch (Exception exception) when (exception is not OperationCanceledException
                    && exception is not OutOfMemoryException && exception is not StackOverflowException)
                {
                    valid = false;
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion,
                        exception.Message, region.SheetName, row + 1, region.Start.Column + index + 1,
                        column.Property.Name, rawValue: raw));
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
    /// 为列表项目类型创建固定列映射。
    /// </summary>
    /// <param name="itemType">列表项目类型。</param>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>可读写的固定实体列。</returns>
    private IReadOnlyList<EntityColumn> CreateColumns(Type itemType, ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
    {
        var plan = _planBuilder.Create(itemType, document ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, configuration, direction);
        if (plan.DynamicColumns.Count > 0)
            throw new BingOfficesUnsupportedFeatureException(
                "ClosedXML Entity List Region 暂不支持动态列。", provider: Provider,
                operation: direction == MappingDirection.Export
                    ? BingOfficesOperation.Export : BingOfficesOperation.Import,
                stage: BingOfficesStage.Plan);
        var columns = new List<EntityColumn>();
        foreach (var column in plan.Columns.Where(item => !item.Ignored && !item.IsDynamicColumn))
        {
            var property = itemType.GetProperty(column.Name, BindingFlags.Instance | BindingFlags.Public);
            if (property != null && property.CanRead && (direction != MappingDirection.Import || property.CanWrite))
                columns.Add(new EntityColumn(column, property));
        }
        if (columns.Count == 0)
            throw new BingOfficesConfigurationException($"实体列表没有可用映射列: {itemType.FullName}",
                stage: BingOfficesStage.Plan);
        return columns;
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
    /// 保存实体列映射与对应属性。
    /// </summary>
    /// <param name="Column">实体映射列。</param>
    /// <param name="Property">实体属性。</param>
    private sealed record EntityColumn(IExcelMappingColumn Column, PropertyInfo Property);
}
