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
/// NPOI 单个实体导入执行器。
/// </summary>
internal sealed class NpoiEntityImportExecutor
{
    /// <summary>
    /// NPOI 导入映射计划构建器。
    /// </summary>
    private readonly NpoiImportPlanBuilder _planBuilder;

    /// <summary>
    /// NPOI 工作表导入执行器。
    /// </summary>
    private readonly NpoiImportSheetExecutor _sheetExecutor = new NpoiImportSheetExecutor(new NpoiImportRowMaterializer());

    /// <summary>
    /// 单元格值转换器集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityImportExecutor" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 此重载使用空的值转换器集合。
    /// </remarks>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    internal NpoiEntityImportExecutor(IExcelMappingPlanFactory mappingPlanFactory)
    {
        _planBuilder = new NpoiImportPlanBuilder(mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory)));
        _valueConverters = Array.Empty<IExcelValueConverter>();
    }

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityImportExecutor" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 此重载使用调用方提供的值转换器集合。
    /// </remarks>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="valueConverters">单元格值转换器集合。</param>
    internal NpoiEntityImportExecutor(IExcelMappingPlanFactory mappingPlanFactory,
        IReadOnlyList<IExcelValueConverter> valueConverters)
    {
        _planBuilder = new NpoiImportPlanBuilder(mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory)));
        _valueConverters = valueConverters ?? Array.Empty<IExcelValueConverter>();
    }

    /// <summary>
    /// 按实体布局从工作簿读取固定单元格、列表区域和合并区域。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="entity">接收导入值的实体。</param>
    /// <param name="layout">实体布局定义。</param>
    /// <param name="requireTemplateMerges">是否要求声明的合并区域已存在。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>包含实体、工作表结果和结构化错误的导入结果。</returns>
    internal ExcelEntityImportResult<TEntity> Read<TEntity>(IWorkbook workbook, TEntity entity,
        ExcelEntityLayout<TEntity> layout, bool requireTemplateMerges, ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        limits?.Validate();
        var errors = new ExcelImportErrorCollector(limits?.MaxErrors);
        var sheetResults = new List<ExcelSheetImportResult>();
        var runtime = new ExcelImportRuntime(limits);
        foreach (var name in layout.Cells.Select(item => item.SheetName)
                     .Concat(layout.ListRegions.Select(item => item.SheetName))
                     .Concat(layout.Merges.Select(item => item.SheetName))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheet(name);
            if (sheet == null)
                throw new BingOfficesConfigurationException($"工作簿缺少请求的 Sheet: {name}", stage: BingOfficesStage.Plan);
            NpoiEntityLayoutSupport.PreflightMerges(sheet, layout.Merges, false, requireTemplateMerges);
        }
        foreach (var binding in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheet(binding.SheetName);
            var resolvedBinding = binding;
            if (binding.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, binding.SheetName,
                    binding.AnchorName);
                sheet = named.Sheet;
                resolvedBinding = binding.WithResolvedReference(named.Reference);
            }
            var anchor = NpoiEntityLayoutSupport.ResolveAnchor(sheet, resolvedBinding.Reference);
            if (resolvedBinding.Setter == null)
                throw new BingOfficesConfigurationException($"实体属性不可写入: {resolvedBinding.PropertyName}",
                    stage: BingOfficesStage.Plan);
            try
            {
                var column = NpoiEntityLayoutSupport.CreateCellPlan(_planBuilder.MappingPlanFactory, resolvedBinding,
                    MappingDirection.Import);
                var cellValue = NpoiExcelImporter.ReadCellValue(NpoiEntityLayoutSupport.GetCell(sheet, anchor, false),
                    workbook.IsDate1904());
                var value = NpoiExcelImporter.NormalizeText(cellValue.Text,
                    column.Property.ImportWhitespace ?? ExcelWhitespacePolicy.Trim);
                var normalizedCellValue = new ExcelCellValue(cellValue.Value, value, cellValue.Kind,
                    cellValue.CachedKind, cellValue.Formula, cellValue.ErrorCode, cellValue.FormatIndex,
                    cellValue.IsDate1904);
                if (!NpoiEntityLayoutSupport.ValidateCellImportRaw(column, normalizedCellValue, value,
                        sheet.SheetName, anchor, errors))
                    continue;
                var converted = column.ConvertFrom(value, normalizedCellValue, sheet.SheetName,
                    anchor.Row + 1, anchor.Column + 1, CultureInfo.InvariantCulture);
                if (!NpoiEntityLayoutSupport.ValidateCellImportConverted(column, normalizedCellValue, value,
                        converted, sheet.SheetName, anchor, errors))
                    continue;
                resolvedBinding.Setter(entity, converted);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message,
                    sheet.SheetName, anchor.Row + 1, anchor.Column + 1, binding.PropertyName,
                    binding.PropertyName, null));
            }
        }
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheet(region.SheetName);
            var resolvedRegion = region;
            if (region.AnchorName != null)
            {
                var named = NpoiEntityLayoutSupport.ResolveNamedAnchor(workbook, region.SheetName,
                    region.AnchorName);
                sheet = named.Sheet;
                resolvedRegion = region.WithResolvedStart(named.Reference);
            }
            NpoiEntityLayoutSupport.ValidateCellRegionConflict(workbook, layout, resolvedRegion);
            ReadRegion(workbook, sheet, entity, resolvedRegion, errors, sheetResults, runtime,
                limits, validationFailureMode, cancellationToken);
        }
        var sourceLocations = new Dictionary<object, SourceLocation>();
        foreach (var relation in layout.Relations ?? Array.Empty<ExcelRelationRequest>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            NpoiRelationBinder.Bind(entity, relation, errors, sourceLocations, cancellationToken);
        }
        return new ExcelEntityImportResult<TEntity>(entity, errors.Errors, sheetResults, errors.IsTruncated,
            errors.MaxErrors);
    }

    /// <summary>
    /// 根据列表区域的运行时元素类型分派强类型读取逻辑。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="sheet">源工作表。</param>
    /// <param name="entity">接收列表数据的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <param name="sheetResults">工作表结果集合。</param>
    /// <param name="runtime">导入运行时状态。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void ReadRegion<TEntity>(IWorkbook workbook, ISheet sheet, TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelImportErrorCollector errors,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportRuntime runtime,
        ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var method = GetType().GetMethod(nameof(ReadTypedRegion), System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        InvokeUnwrapped(method, this, new object[] { workbook, sheet, entity, region, errors, sheetResults, runtime,
            limits, validationFailureMode, cancellationToken });
    }

    /// <summary>
    /// 从布局区域读取强类型列表，并将临时工作表行号映射回源工作表。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="source">源工作表。</param>
    /// <param name="entity">接收列表数据的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="errors">导入错误收集器。</param>
    /// <param name="sheetResults">工作表结果集合。</param>
    /// <param name="runtime">导入运行时状态。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="validationFailureMode">当前行校验失败后的继续策略。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    private void ReadTypedRegion<TEntity, TItem>(IWorkbook workbook, ISheet source, TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelImportErrorCollector errors,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportRuntime runtime,
        ExcelResourceLimits limits,
        ExcelValidationFailureMode validationFailureMode,
        CancellationToken cancellationToken) where TEntity : class, new() where TItem : class, new()
    {
        var groups = region.DynamicColumnGroups ?? Array.Empty<IExcelEntityDynamicColumnGroup>();
        var groupDefinitions = groups.SelectMany(group => group.Definitions).ToArray();
        var request = ExcelImport.Workbook<RegionRoot<TItem>>(builder => builder.Sheet(region.SheetName,
            root => root.Items, configure =>
            {
                configure.HeaderRowIndex(0).DataRowStartIndex(1);
                if (region.MappingDocument != null)
                    configure.Mapping(region.MappingDocument);
                else if (region.MappingConfiguration != null)
                    configure.Mapping(region.MappingConfiguration);
                if (groupDefinitions.Length > 0)
                    configure.DynamicColumns(CreateDynamicTargetGetter<TItem>(groups), groupDefinitions);
                configure.FailOnUnknownDynamicColumns(
                    region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail);
            }));
        var resolved = new NpoiResolvedSheet(request.Sheets[0], 0, region.SheetName);
        var mapping = _planBuilder.Create(new[] { resolved })[request.Sheets[0]];
        var dynamicProperty = groups.Count == 0 ? GetDynamicProperty<TItem>(mapping) : null;
        var dynamicDefinitions = groupDefinitions.Length > 0
            ? groupDefinitions : CreateDynamicDefinitions(mapping);
        var columns = groups.Count == 0
            ? NpoiExportColumnPlanner.CreateColumns<TItem>(mapping, dynamicDefinitions,
                region.CalculatedColumns)
            : NpoiExportColumnPlanner.CreateEntityColumns<TItem>(mapping, dynamicDefinitions, groups,
                region.CalculatedColumns);
        var dynamicTargets = new List<(TItem Item, IDictionary<string, object> Values)>();
        var itemCount = region.End.HasValue
            ? region.End.Value.Row - region.Start.Row + (region.IncludeHeader ? 0 : 1)
            : 0;
        NpoiEntityLayoutSupport.ValidateListRegionBounds(region, columns.Count, itemCount, region.IncludeHeader);
        var temp = CreateRegionSheet<TEntity, TItem>(workbook, source, region, columns.Count, columns,
            workbook.IsDate1904(),
            cancellationToken);
        try
        {
            var options = new ExcelImportExecutionOptions<TItem>
            {
                HeaderRowIndex = 0,
                DataRowIndex = 1,
                MaxReadColumns = Math.Max(1, columns.Count),
                MappingPlan = mapping,
                Culture = CultureInfo.InvariantCulture,
                IsDate1904 = workbook.IsDate1904(),
                HeaderWhitespace = ExcelWhitespacePolicy.Trim,
                BodyWhitespace = ExcelWhitespacePolicy.Trim,
                ValidationFailureMode = validationFailureMode,
                RequireExpectedHeaders = true,
                ValidationMode = ExcelImportValidationMode.ConfiguredRules,
                MaxTrackedUniqueValues = limits?.MaxTrackedUniqueValues,
                UniqueComparison = limits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase,
                StopAtFirstEmptyRow = false,
                DynamicColumns = dynamicDefinitions,
                DynamicTargetGetter = groups.Count == 0 && dynamicProperty != null
                    ? item => dynamicProperty.GetValue(item)
                    : groups.Count == 0 ? null : item => CreateTemporaryDynamicValues(item, dynamicTargets),
                DynamicPropertyNames = groups.Count == 0 ? null :
                    new HashSet<string>(groups.Select(group => group.Property.Name), StringComparer.OrdinalIgnoreCase),
                DynamicColumnPropertyNames = groups.Count == 0 ? null : groupDefinitions
                    .ToDictionary(definition => definition.Key,
                        definition => groups.Single(group => group.Definitions.Any(column =>
                            string.Equals(column.Key, definition.Key, StringComparison.OrdinalIgnoreCase)))
                            .Property.Name, StringComparer.OrdinalIgnoreCase)
                , ErrorSheetName = region.SheetName
            };
            var items = new List<TItem>();
            var rows = new List<int>();
            var child = errors.CreateChild();
            _sheetExecutor.Execute(temp, options, items, child, runtime, cancellationToken, rows);
            if (groups.Count > 0)
                ApplyDynamicGroups(dynamicTargets, groups);
            AssignItems(entity, region, items);
            var sourceRows = rows.Select(row => region.Start.Row + row - (region.IncludeHeader ? 0 : 1)).ToArray();
            sheetResults.Add(new ExcelSheetImportResult(source.SheetName, typeof(TItem), sourceRows, child.Errors));
        }
        finally
        {
            workbook.RemoveSheetAt(workbook.GetSheetIndex(temp));
        }
    }

    /// <summary>
    /// 创建实体布局动态列组的导入目标表达式。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>供标准导入请求保存动态定义的目标表达式。</returns>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicTargetGetter<TItem>(
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        var parameter = Expression.Parameter(typeof(TItem), "item");
        var method = typeof(NpoiEntityImportExecutor).GetMethod(nameof(CreateDynamicTarget),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(typeof(TItem));
        var body = Expression.Call(method, parameter, Expression.Constant(groups));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 创建标准导入请求保存的临时动态目标字典。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">列表项。</param>
    /// <param name="groups">动态列组集合。</param>
    /// <returns>新的动态值字典。</returns>
    private static IDictionary<string, object> CreateDynamicTarget<TItem>(TItem item,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new() =>
        new Dictionary<string, object>(StringComparer.Ordinal);

    /// <summary>
    /// 创建单行导入使用的临时动态字典，并保留其所属实体。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="item">标准导入器传入的列表项。</param>
    /// <param name="targets">临时目标集合。</param>
    /// <returns>供标准导入器写入的临时字典。</returns>
    private static IDictionary<string, object> CreateTemporaryDynamicValues<TItem>(object item,
        ICollection<(TItem Item, IDictionary<string, object> Values)> targets) where TItem : class, new()
    {
        var values = new Dictionary<string, object>(StringComparer.Ordinal);
        targets.Add(((TItem)item, values));
        return values;
    }

    /// <summary>
    /// 将标准导入器收集的动态值分发回各实体动态列组。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="targets">实体与临时动态字典集合。</param>
    /// <param name="groups">动态列组集合。</param>
    private static void ApplyDynamicGroups<TItem>(
        IEnumerable<(TItem Item, IDictionary<string, object> Values)> targets,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups) where TItem : class, new()
    {
        foreach (var target in targets)
        {
            foreach (var group in groups)
            {
                var keys = new HashSet<string>(group.Definitions.Select(definition => definition.Key),
                    StringComparer.OrdinalIgnoreCase);
                var values = target.Values.Where(pair => keys.Contains(pair.Key)).ToArray();
                if (values.Length == 0)
                    continue;
                var destination = group.Getter(target.Item);
                if (destination == null)
                {
                    if (group.Setter == null)
                        throw new BingOfficesConfigurationException(
                            $"实体动态列属性不可写入: {group.Property.Name}", stage: BingOfficesStage.Plan);
                    try
                    {
                        var type = group.Property.PropertyType;
                        var instance = type.IsInterface || type.IsAbstract
                            ? new Dictionary<string, object>(StringComparer.Ordinal)
                            : Activator.CreateInstance(type);
                        if (instance is not IDictionary<string, object> created)
                            throw new InvalidOperationException();
                        destination = created;
                        group.Setter(target.Item, destination);
                    }
                    catch (Exception exception) when (exception is not BingOfficesException
                        && exception is not OperationCanceledException
                        && exception is not OutOfMemoryException && exception is not StackOverflowException)
                    {
                        throw new BingOfficesConfigurationException(
                            $"实体动态列属性无法创建字典实例: {group.Property.Name}", exception,
                            BingOfficesStage.Plan);
                    }
                }
                foreach (var pair in values)
                    destination[pair.Key] = pair.Value;
            }
        }
    }

    /// <summary>
    /// 获取映射计划中声明的单个动态字典属性。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>动态字典属性；未声明时返回 <see langword="null" />。</returns>
    private static PropertyInfo GetDynamicProperty<TItem>(IExcelMappingPlan mapping)
        where TItem : class, new()
    {
        var properties = (mapping?.Columns ?? Array.Empty<IExcelMappingColumn>())
            .Where(column => column.IsDynamicColumn).ToArray();
        if (properties.Length == 0)
            return null;
        if (properties.Length > 1)
            throw new BingOfficesConfigurationException(
                $"导入实体列表项 {typeof(TItem).FullName} 只能声明一个动态列字典属性。",
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
    /// 将映射计划中的动态列转换为实体布局导入定义。
    /// </summary>
    /// <param name="mapping">列表项映射计划。</param>
    /// <returns>按映射顺序排列的动态列定义。</returns>
    private static IReadOnlyList<ExcelDynamicColumnDefinition> CreateDynamicDefinitions(
        IExcelMappingPlan mapping) => (mapping?.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>())
        .Select(column => NpoiExportColumnPlanner.CreateDynamicDefinition(column, null, mapping))
        .ToArray();

    /// <summary>
    /// 将源列表区域投影到标准零基标题和数据行布局的临时工作表。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="workbook">承载临时工作表的工作簿。</param>
    /// <param name="source">源工作表。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="width">区域列数。</param>
    /// <param name="columns">导入列计划。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>供标准工作表导入执行器读取的临时工作表。</returns>
    private static ISheet CreateRegionSheet<TEntity, TItem>(IWorkbook workbook, ISheet source,
        ExcelEntityListRegion<TEntity> region, int width, IReadOnlyList<ExcelColumnPlan> columns,
        bool isDate1904, CancellationToken cancellationToken) where TEntity : class, new()
        where TItem : class, new()
    {
        var temp = workbook.CreateSheet("__BingEnt" + Guid.NewGuid().ToString("N").Substring(0, 12));
        var header = temp.CreateRow(0);
        var sourceHeader = region.IncludeHeader ? source.GetRow(region.Start.Row) : null;
        for (var index = 0; index < width; index++)
        {
            var headerCell = header.CreateCell(index);
            if (sourceHeader != null)
                CopyCell(sourceHeader.GetCell(region.Start.Column + index), headerCell, isDate1904);
            else
                headerCell.SetCellValue(columns[index].Title);
        }
        var startRow = region.Start.Row + (region.IncludeHeader ? 1 : 0);
        var footerMarkerRow = NpoiEntityLayoutSupport.FindFooterMarkerRow(source, region);
        var lastRow = footerMarkerRow.HasValue
            ? footerMarkerRow.Value - 1 - (region.Footer?.GapRows ?? 0)
            : (region.End?.Row ?? source.LastRowNum);
        var skipSpans = region.GroupSubtotal == null && region.PageSubtotal == null
            ? Array.Empty<(int StartRow, int EndRow)>()
            : NpoiEntityLayoutSupport.FindFooterSkipSpans(source, region)
                .Where(span => span.StartRow <= lastRow).ToArray();
        var outputRow = 1;
        for (var rowIndex = startRow; rowIndex <= lastRow; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (skipSpans.Any(span => rowIndex >= span.StartRow && rowIndex <= span.EndRow))
                continue;
            var sourceRow = source.GetRow(rowIndex);
            var targetRow = temp.CreateRow(outputRow);
            for (var column = 0; column < width; column++)
                CopyCell(sourceRow?.GetCell(region.Start.Column + column), targetRow.CreateCell(column), isDate1904);
            outputRow++;
        }
        return temp;
    }

    /// <summary>
    /// 将源单元格的样式和值复制到临时工作表单元格。
    /// </summary>
    /// <param name="source">源单元格。</param>
    /// <param name="target">目标单元格。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    private static void CopyCell(ICell source, ICell target, bool isDate1904)
    {
        if (source == null)
            return;
        target.CellStyle = source.CellStyle;
        var value = NpoiExcelImporter.ReadCellValue(source, isDate1904);
        if (value.Kind == ExcelCellKind.Empty)
            return;
        target.SetCellValue(value.Value ?? value.Text);
    }

    /// <summary>
    /// 将导入的列表元素赋给实体集合属性或填充现有集合实例。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    /// <param name="entity">接收列表数据的实体。</param>
    /// <param name="region">列表区域定义。</param>
    /// <param name="items">已导入的列表元素。</param>
    private static void AssignItems<TEntity, TItem>(TEntity entity, ExcelEntityListRegion<TEntity> region,
        IReadOnlyList<TItem> items) where TEntity : class, new() where TItem : class, new()
    {
        if (region.Setter != null)
        {
            region.Setter(entity, items.ToList());
            return;
        }
        var current = region.Getter(entity) as ICollection<TItem>;
        if (current == null)
            throw new BingOfficesConfigurationException($"实体集合属性不可写入: {region.PropertyName}",
                stage: BingOfficesStage.Plan);
        current.Clear();
        foreach (var item in items)
            current.Add(item);
    }

    /// <summary>
    /// 为标准工作表导入管线提供列表属性的临时根对象。
    /// </summary>
    /// <typeparam name="TItem">列表元素类型。</typeparam>
    private sealed class RegionRoot<TItem> where TItem : class, new()
    {
        /// <summary>
        /// 获取接收导入数据的列表集合。
        /// </summary>
        public ICollection<TItem> Items { get; } = new List<TItem>();
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
