using System.Reflection;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Metadata;
using Bing.Offices.Providers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// 按工作表执行已解析的 NPOI 导入计划。
/// </summary>
internal sealed class NpoiImportSheetExecutor
{
    /// <summary>
    /// 负责将工作表行转换为实体并收集字段错误的物化器。
    /// </summary>
    private readonly NpoiImportRowMaterializer _rowMaterializer;

    /// <summary>
    /// 初始化一个 <see cref="NpoiImportSheetExecutor" /> 类型的实例。
    /// </summary>
    /// <param name="rowMaterializer">负责将工作表行物化为实体的执行器。</param>
    internal NpoiImportSheetExecutor(NpoiImportRowMaterializer rowMaterializer)
    {
        _rowMaterializer = rowMaterializer ?? throw new ArgumentNullException(nameof(rowMaterializer));
    }

    /// <summary>
    /// 执行单个泛型工作表的行导入，并通过输出集合记录成功行索引。
    /// </summary>
    /// <typeparam name="T">当前工作表的实体类型。</typeparam>
    /// <param name="sheet">待导入的 NPOI 工作表。</param>
    /// <param name="options">当前实体类型的导入执行选项。</param>
    /// <param name="items">接收成功导入实体的集合。</param>
    /// <param name="errors">接收导入错误的收集器。</param>
    /// <param name="runtime">当前导入运行时状态。</param>
    /// <param name="cancellationToken">逐行导入过程中检查的取消令牌。</param>
    /// <param name="sourceRows">接收成功实体对应源行索引的集合；为空时不记录。</param>
    internal void Execute<T>(ISheet sheet, ExcelImportExecutionOptions<T> options, ICollection<T> items,
        ExcelImportErrorCollector errors, ExcelImportRuntime runtime, CancellationToken cancellationToken,
        ICollection<int> sourceRows = null) where T : class, new()
    {
        var header = sheet.GetRow(options.HeaderRowIndex)
            ?? throw new NpoiSheetStructureException("导入的模板不正确，未匹配表头。");
        if (header.LastCellNum > options.MaxReadColumns)
            throw new NpoiSheetStructureException($"导入表头超过最大列长度: {options.MaxReadColumns}");

        var columns = CreateColumns<T>(header, options);
        IReadOnlyDictionary<(int Row, int Column), IReadOnlyList<PictureInfo>> imageIndex = null;
        HashSet<int> imageRows = null;
        if (columns.Values.Any(NpoiImportRowMaterializer.IsImageColumn))
        {
            try
            {
                imageIndex = NpoiImportRowMaterializer.BuildImageIndex(sheet, runtime.ImageResources,
                    cancellationToken, out imageRows);
            }
            catch (ImageResourceLimitException exception)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit, exception.Message,
                    sheet.SheetName, options.HeaderRowIndex + 1, 0, null));
                return;
            }
        }
        var workbookValidations = options.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || options.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook
            ? sheet.GetDataValidations().ToArray()
            : Array.Empty<IDataValidation>();
        var validationIndex = ValidationRangeIndex.Create(workbookValidations, options.DataRowIndex, sheet.LastRowNum,
            0, Math.Max(0, header.LastCellNum - 1));
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var uniqueTracker = new UniqueTracker(duplicateValues, options.MaxTrackedUniqueValues,
            CreateStringComparer(options.UniqueComparison));
        var configuredValidationEnabled = IsConfiguredValidationEnabled(options.ValidationMode);
        for (var rowIndex = options.DataRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            if (!runtime.TryConsumeRow())
            {
                if (runtime.TryMarkRowLimitReported())
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                        $"Workbook 数据行数超过限制: {runtime.MaxRows}", sheet.SheetName, rowIndex + 1, 0, null));
                break;
            }
            var row = sheet.GetRow(rowIndex);
            if (NpoiImportRowMaterializer.IsEmpty(row, options.BodyWhitespace, imageRows, rowIndex))
            {
                if (options.StopAtFirstEmptyRow)
                    break;
                if (options.ReportEmptyRows)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidInput, "导入数据存在空行", sheet.SheetName,
                        rowIndex + 1, 0, null));
                    if (errors.IsLimitReached)
                    {
                        errors.MarkTruncated();
                        break;
                    }
                }
                continue;
            }
            if (configuredValidationEnabled)
                uniqueTracker.BeginRow();
            var workbookValid = NpoiWorkbookValidationPipeline.Validate(row, columns, validationIndex, sheet,
                sheet.SheetName, rowIndex, options.BodyWhitespace, options.ValidationFailureMode,
                options.UnsupportedFeaturePolicy, errors, options.IsDate1904);
            if (!workbookValid)
            {
                if (configuredValidationEnabled)
                    uniqueTracker.RollbackRow();
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                continue;
            }
            if (configuredValidationEnabled && !_rowMaterializer.ValidateRawValues(row, columns, duplicateValues,
                    sheet.SheetName, rowIndex, options.ValidationFailureMode, options.Culture, options.BodyWhitespace, errors,
                    options.IsDate1904))
            {
                uniqueTracker.RollbackRow();
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                continue;
            }
            if (_rowMaterializer.TryCreateItem(row, columns, duplicateValues, uniqueTracker, sheet.SheetName,
                    rowIndex, options.ValidationFailureMode, configuredValidationEnabled, errors, options.Culture,
                    options.BodyWhitespace, options.DynamicTargetGetter, imageIndex, options.IsDate1904,
                    out T item))
            {
                items.Add(item);
                sourceRows?.Add(rowIndex);
                if (configuredValidationEnabled)
                    uniqueTracker.CommitRow();
            }
            else if (configuredValidationEnabled)
                uniqueTracker.RollbackRow();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
        }
    }

    /// <summary>
    /// 判断导入验证模式是否启用配置校验规则。
    /// </summary>
    /// <param name="mode">当前导入验证模式。</param>
    /// <returns>模式包含配置校验时为 true，否则为 false。</returns>
    private static bool IsConfiguredValidationEnabled(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.ConfiguredRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    /// <summary>
    /// 根据表头和映射计划创建列执行计划。
    /// </summary>
    /// <typeparam name="T">导入实体类型。</typeparam>
    /// <param name="header">包含导入表头的工作表行。</param>
    /// <param name="options">当前实体类型的导入执行选项。</param>
    /// <returns>按零基列索引排列的列执行计划。</returns>
    private static IReadOnlyDictionary<int, ExcelColumnPlan> CreateColumns<T>(IRow header,
        ExcelImportExecutionOptions<T> options) where T : class, new()
    {
        var map = options.MappingPlan;
        if (map == null)
            throw new BingOfficesConfigurationException("工作表导入计划不可用。",
                stage: BingOfficesStage.Plan);
        var dynamicProperties = map.Columns.Where(property => property.IsDynamicColumn).ToList();
        var dynamicPlans = map.DynamicColumns;
        if (dynamicProperties.Count > 1)
            throw new BingOfficesConfigurationException(
                $"导入模板 {typeof(T).FullName} 只能声明一个动态列属性。", stage: BingOfficesStage.Plan);
        var fixedProperties = map.Columns.Where(property => !property.Ignored && !property.IsDynamicColumn
            && !IsNavigationOrDynamicContainer<T>(property)).ToList();
        var headerNames = new HashSet<string>(options.HeaderComparison == ExcelNameComparison.Ordinal
            ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase);
        var columns = new Dictionary<int, ExcelColumnPlan>();
        foreach (var headerCell in header.Cells)
        {
            if (options.ReadColumnRange != null && !options.ReadColumnRange.Contains(headerCell.ColumnIndex))
                continue;
            var headerName = NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(headerCell),
                options.HeaderWhitespace);
            if (string.IsNullOrWhiteSpace(headerName))
                continue;
            if (!headerNames.Add(headerName))
                throw new NpoiSheetStructureException($"导入的表格存在重复列:{headerName}");
            var property = FindProperty(fixedProperties, headerName, options.HeaderComparison);
            ExcelDynamicColumnDefinition dynamicDefinition = null;
            IExcelDynamicMappingColumn dynamicPlan = null;
            if (property == null && dynamicProperties.Count == 1)
            {
                dynamicPlan = FindDynamicDefinition(headerName, dynamicPlans, options.HeaderComparison);
                dynamicDefinition = dynamicPlan == null ? null : CreateDynamicDefinition(dynamicPlan);
                if (dynamicPlans.Count > 0 && dynamicDefinition == null)
                {
                    if (options.FailOnUnknownDynamicColumns)
                        throw new NpoiSheetStructureException($"导入包含未知动态列: {headerName}");
                    continue;
                }
                property = dynamicProperties[0];
            }
            if (property == null)
                continue;
            var isUnspecifiedDynamicColumn = property.IsDynamicColumn && dynamicDefinition == null;
            var reflectionProperty = typeof(T).GetProperty(property.Name, BindingFlags.Instance | BindingFlags.Public);
            if (reflectionProperty == null)
                throw new BingOfficesConfigurationException($"无法解析映射属性: {property.Name}",
                    stage: BingOfficesStage.Plan);
            if (!property.IsDynamicColumn && !reflectionProperty.CanWrite)
            {
                var cause = new InvalidOperationException($"属性不可写入: {property.Name}");
                throw new BingOfficesConfigurationException(cause.Message, cause, BingOfficesStage.Plan);
            }
            var valueConverters = isUnspecifiedDynamicColumn
                ? (IReadOnlyList<Conversions.IExcelValueConverter>)Array.Empty<Conversions.IExcelValueConverter>()
                : property.IsDynamicColumn ? dynamicPlan.ValueConverters : property.ValueConverters;
            var validationBindings = property.IsDynamicColumn && dynamicPlan != null
                ? dynamicPlan.ValidationBindings : property.ValidationBindings;
            columns[headerCell.ColumnIndex] = new ExcelColumnPlan(headerName, property, property.IsDynamicColumn,
                headerCell.ColumnIndex, dynamicDefinition, null, valueConverters, validationBindings,
                reflectionProperty: reflectionProperty, isUnique: dynamicPlan?.IsUnique,
                uniqueIgnoreEmpty: dynamicPlan?.UniqueIgnoreEmpty ?? true);
        }
        if (options.RequireExpectedHeaders)
        {
            var missing = fixedProperties.Where(property => !columns.Values.Any(column => column.Property == property)
                && (options.ReadColumnRange == null || !header.Cells.Any(cell =>
                    !options.ReadColumnRange.Contains(cell.ColumnIndex)
                    && (string.Equals(NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(cell),
                            options.HeaderWhitespace), property.Title, ToStringComparison(options.HeaderComparison))
                        || property.Aliases.Any(alias => string.Equals(
                            NpoiExcelImporter.NormalizeText(NpoiExcelImporter.GetRawStringValue(cell),
                                options.HeaderWhitespace), alias, ToStringComparison(options.HeaderComparison)))))))
                .ToList();
            if (missing.Any())
                throw new NpoiSheetStructureException($"导入的表格不存在列：{string.Join(",", missing.Select(property => property.Title))}");
        }
        return columns;
    }

    /// <summary>
    /// 按标题或别名查找动态列定义。
    /// </summary>
    /// <param name="headerName">待匹配的表头文本。</param>
    /// <param name="definitions">可用的动态列定义集合。</param>
    /// <param name="comparison">表头名称比较规则。</param>
    /// <returns>匹配的动态列定义；未匹配时返回 <see langword="null" />。</returns>
    private static IExcelDynamicMappingColumn FindDynamicDefinition(string headerName,
        IReadOnlyList<IExcelDynamicMappingColumn> definitions, ExcelNameComparison comparison) =>
        (definitions ?? Array.Empty<IExcelDynamicMappingColumn>()).FirstOrDefault(definition =>
            string.Equals(definition.Title, headerName, ToStringComparison(comparison))
            || (definition.Aliases ?? Array.Empty<string>()).Any(alias =>
                string.Equals(alias, headerName, ToStringComparison(comparison))));

    /// <summary>
    /// 将内部动态列映射转换为导入执行定义。
    /// </summary>
    /// <param name="column">内部动态列映射。</param>
    /// <returns>可供导入列计划使用的动态列定义。</returns>
    private static ExcelDynamicColumnDefinition CreateDynamicDefinition(IExcelDynamicMappingColumn column) => new()
    {
        Key = column.Key,
        Title = column.Title,
        Aliases = column.Aliases,
        DataType = ResolveDynamicType(column.DataTypeName),
        Order = column.Order,
        Placement = CreatePlacement(column.PlacementKey),
        PhysicalColumnIndex = column.ColumnIndex,
        NumberFormat = column.NumberFormat,
        ConverterName = column.ConverterName,
        ValidatorName = column.ValidatorName,
        ValidationRuleNames = column.ValidationRuleNames,
        ImageMultiplicity = column.ImageMultiplicity
    };

    /// <summary>
    /// 解析动态列相对于固定列的放置键。
    /// </summary>
    /// <param name="placementKey">before/after 形式的放置键；为空时不指定位置。</param>
    /// <returns>解析出的列放置定义；键为空时返回 <see langword="null" />。</returns>
    private static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var separator = placementKey.IndexOfAny(new[] { ':', '-' });
        var key = placementKey.Substring(separator + 1);
        return placementKey.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || placementKey.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key) : ExcelColumnPlacement.After(key);
    }

    /// <summary>
    /// 将动态列类型名称解析为 CLR 类型。
    /// </summary>
    /// <param name="name">允许列表中的类型名称；为空时按 string 处理。</param>
    /// <returns>解析出的 CLR 类型。</returns>
    private static Type ResolveDynamicType(string name) => (name ?? "string").ToLowerInvariant() switch
    {
        "object" => typeof(object),
        "string" => typeof(string),
        "boolean" or "bool" => typeof(bool),
        "byte" => typeof(byte),
        "int16" => typeof(short),
        "int32" or "int" => typeof(int),
        "int64" or "long" => typeof(long),
        "single" or "float" => typeof(float),
        "double" => typeof(double),
        "decimal" => typeof(decimal),
        "datetime" => typeof(DateTime),
        "datetimeoffset" => typeof(DateTimeOffset),
        "guid" => typeof(Guid),
        "bytes" => typeof(byte[]),
        _ => throw new NpoiSheetStructureException($"动态列数据类型不在允许列表中: {name}")
    };

    /// <summary>
    /// 按标题、别名或属性名查找固定列映射。
    /// </summary>
    /// <param name="properties">可用的固定列映射集合。</param>
    /// <param name="headerName">待匹配的表头文本。</param>
    /// <param name="comparison">表头名称比较规则。</param>
    /// <returns>匹配的固定列映射；未匹配时返回 <see langword="null" />。</returns>
    private static IExcelMappingColumn FindProperty(IEnumerable<IExcelMappingColumn> properties, string headerName,
        ExcelNameComparison comparison)
    {
        var stringComparison = ToStringComparison(comparison);
        return properties.FirstOrDefault(property => string.Equals(property.Title, headerName, stringComparison)
            || property.Aliases.Any(alias => string.Equals(alias, headerName, stringComparison))
            || string.Equals(property.Name, headerName, stringComparison));
    }

    /// <summary>
    /// 判断关系集合或动态字典是否应由关系/动态列阶段处理，而不是作为单元格列导入。
    /// </summary>
    /// <typeparam name="T">当前工作表行类型。</typeparam>
    /// <param name="property">待判断的映射列。</param>
    /// <returns>属性为非图片可枚举集合或动态字典时返回 true。</returns>
    private static bool IsNavigationOrDynamicContainer<T>(IExcelMappingColumn property)
        where T : class, new()
    {
        var reflectionProperty = typeof(T).GetProperty(property.Name, BindingFlags.Instance | BindingFlags.Public);
        if (reflectionProperty == null)
            return false;
        var propertyType = reflectionProperty.PropertyType;
        if (typeof(IDictionary<string, object>).IsAssignableFrom(propertyType))
            return true;
        if (propertyType == typeof(string) || propertyType == typeof(byte[])
            || !typeof(System.Collections.IEnumerable).IsAssignableFrom(propertyType))
            return false;
        return !IsImageCollection(propertyType);
    }

    /// <summary>
    /// 判断集合是否承载 NPOI 导入器支持的图片值。
    /// </summary>
    /// <param name="propertyType">待判断的属性类型。</param>
    /// <returns>集合元素为 <see cref="ExcelImageData" /> 时返回 true。</returns>
    private static bool IsImageCollection(Type propertyType)
    {
        if (propertyType.IsArray)
            return propertyType.GetElementType() == typeof(ExcelImageData);
        foreach (var interfaceType in propertyType.GetInterfaces())
        {
            if (interfaceType.IsGenericType
                && interfaceType.GetGenericTypeDefinition() == typeof(IEnumerable<>)
                && interfaceType.GetGenericArguments()[0] == typeof(ExcelImageData))
                return true;
        }
        return propertyType.IsGenericType
            && propertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>)
            && propertyType.GetGenericArguments()[0] == typeof(ExcelImageData);
    }

    /// <summary>
    /// 将导入名称比较策略转换为字符串比较枚举。
    /// </summary>
    /// <param name="comparison">导入名称比较策略。</param>
    /// <returns>对应的字符串比较方式。</returns>
    private static StringComparison ToStringComparison(ExcelNameComparison comparison) =>
        comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    /// <summary>
    /// 根据字符串比较枚举创建对应的比较器。
    /// </summary>
    /// <param name="comparison">字符串比较方式。</param>
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

}

/// <summary>
/// 表示导入工作表结构不符合请求的异常。
/// </summary>
internal sealed class NpoiSheetStructureException : InvalidOperationException
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiSheetStructureException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述工作表结构错误的消息。</param>
    internal NpoiSheetStructureException(string message) : base(message) { }
}
