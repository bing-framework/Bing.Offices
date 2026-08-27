using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Internals;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 基于 NPOI 的单遍流式 Excel 导入器。
/// </summary>
internal sealed class NpoiExcelImporter : IExcelImporter
{
    /// <summary>
    /// 当前导入器使用的校验规则。
    /// </summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;

    /// <summary>
    /// 当前导入器使用的值转换器。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 当前导入器使用的旧版仅文本单元格转换器。
    /// </summary>
    private readonly IReadOnlyList<ICellValueConverter> _legacyValueConverters;
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    private readonly NpoiImportPlanBuilder _planBuilder;

    /// <summary>
    /// 当前导入器使用的命名配置校验规则。
    /// </summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    private readonly NpoiImportRowMaterializer _rowMaterializer;

    /// <summary>
    /// 初始化一个<see cref="NpoiExcelImporter"/>类型的实例。
    /// </summary>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    /// <param name="legacyValueConverters">旧版仅文本单元格转换器集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    public NpoiExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IEnumerable<ICellValueConverter> legacyValueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null)
    {
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _legacyValueConverters = legacyValueConverters?.ToArray() ?? Array.Empty<ICellValueConverter>();
        _mappingPlanFactory = mappingPlanFactory ?? NpoiMappingPlanFactoryResolver.CreateDefault(
            _valueConverters, _validationRules, _namedValidationRules);
        _planBuilder = new NpoiImportPlanBuilder(_mappingPlanFactory);
        _rowMaterializer = new NpoiImportRowMaterializer(_legacyValueConverters);
    }

    /// <inheritdoc />
    public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();

        using var bufferedSource = new MemoryStream();
        NpoiStreamCopier.Copy(source, bufferedSource, cancellationToken, request.ResourceLimits?.MaxInputBytes);
        bufferedSource.Position = 0;
        using var workbook = WorkbookFactory.Create(bufferedSource);
        var root = new TWorkbook();
        var sheetResults = new List<ExcelSheetImportResult>();
        var errors = new ExcelImportErrorCollector(request.ResourceLimits?.MaxErrors);
        var sourceLocations = request.Relations.Count == 0
            ? null
            : new Dictionary<object, SourceLocation>(ReferenceObjectComparer.Instance);
        var runtime = new ExcelImportRuntime(request.ResourceLimits);
        var existingSheets = request.Sheets.Select(sheet => new KeyValuePair<ExcelSheetImportRequest, int>(
            sheet, ResolveSheetIndex(workbook, sheet.Selector, request.SheetNameComparison)))
            .Where(item => item.Value >= 0).ToArray();
        var plans = _planBuilder.Create(request, workbook, existingSheets);
        var resolvedSheetRequests = new Dictionary<string, ExcelSheetImportRequest>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheetRequest in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached || runtime.RowLimitReached)
                break;
            var sheetIndex = ResolveSheetIndex(workbook, sheetRequest.Selector, request.SheetNameComparison);
            if (sheetIndex < 0)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少请求的 Sheet: {GetSelectorDescription(sheetRequest.Selector)}",
                    GetSelectorDescription(sheetRequest.Selector), 0, 0, null));
                continue;
            }
            if (workbook.IsSheetHidden(sheetIndex) || workbook.IsSheetVeryHidden(sheetIndex))
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"请求的 Sheet 被隐藏: {workbook.GetSheetName(sheetIndex)}", workbook.GetSheetName(sheetIndex),
                    0, 0, null));
                continue;
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            resolvedSheetRequests[sheet.SheetName] = sheetRequest;
            try
            {
                ImportTypedSheet(sheet, sheetRequest, root, sheetResults, errors, request.ValidationMode,
                    request.ResourceLimits, request.UnsupportedFeaturePolicy, sourceLocations, runtime,
                    cancellationToken, plans[sheetRequest]);
            }
            catch (SheetStructureException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    sheet.SheetName, sheetRequest.HeaderRowIndex + 1, 0, null);
                errors.Add(error);
                sheetResults.Add(new ExcelSheetImportResult(sheet.SheetName, sheetRequest.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
        }

        foreach (var relation in request.Relations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
            NpoiRelationBinder.Bind(root, relation, errors, sourceLocations, cancellationToken);
        }
        NpoiFailureWorkbookWriter.Write(workbook, request.FailureOptions, errors.Errors, resolvedSheetRequests,
            cancellationToken);
        return new ExcelWorkbookImportResult<TWorkbook>(root, sheetResults, errors.Errors,
            errors.IsTruncated, errors.MaxErrors);
    }

    private static int ResolveSheetIndex(IWorkbook workbook, ExcelSheetSelector selector,
        ExcelNameComparison comparison)
    {
        if (selector.Kind == ExcelSheetSelectorKind.ByIndex)
            return selector.Index.Value < workbook.NumberOfSheets ? selector.Index.Value : -1;
        var comparer = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        for (var index = 0; index < workbook.NumberOfSheets; index++)
        {
            if (string.Equals(workbook.GetSheetName(index), selector.Name, comparer))
                return index;
        }
        return -1;
    }

    private static string GetSelectorDescription(ExcelSheetSelector selector) =>
        selector.Kind == ExcelSheetSelectorKind.ByIndex ? $"#{selector.Index.Value}" : selector.Name;

    /// <summary>
    /// 通过一次类型擦除调用执行单个 Sheet 导入计划。
    /// </summary>
    private void ImportTypedSheet<TWorkbook>(ISheet sheet, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportErrorCollector errors,
        ExcelImportValidationMode validationMode, ExcelResourceLimits resourceLimits,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations,
        ExcelImportRuntime runtime,
        CancellationToken cancellationToken, IExcelMappingPlan mappingPlan)
        where TWorkbook : class, new()
    {
        var method = GetType().GetMethod(nameof(ImportTypedSheetCore), BindingFlags.Instance | BindingFlags.NonPublic);
        try
        {
            method.MakeGenericMethod(typeof(TWorkbook), request.ItemType).Invoke(this,
                new object[] { sheet, request, root, sheetResults, errors, validationMode, resourceLimits,
                    unsupportedFeaturePolicy, sourceLocations, runtime, cancellationToken, mappingPlan });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>
    /// 导入一个具体实体类型，并将成功项写入 Workbook 根集合。
    /// </summary>
    private void ImportTypedSheetCore<TWorkbook, TItem>(ISheet sheet, ExcelSheetImportRequest request,
        TWorkbook root, ICollection<ExcelSheetImportResult> sheetResults, ExcelImportErrorCollector errors,
        ExcelImportValidationMode validationMode, ExcelResourceLimits resourceLimits,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations,
        ExcelImportRuntime runtime,
        CancellationToken cancellationToken, IExcelMappingPlan mappingPlan)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var options = new ExcelImportExecutionOptions<TItem>
        {
            HeaderRowIndex = request.HeaderRowIndex,
            DataRowIndex = request.DataRowStartIndex,
            MaxColumnLength = request.MaxColumnLength,
            ReadColumnRange = request.ReadColumnRange,
            HeaderComparison = request.HeaderComparison,
            HeaderWhitespace = request.HeaderWhitespace,
            BodyWhitespace = request.BodyWhitespace,
            ValidationMode = validationMode,
            UnsupportedFeaturePolicy = unsupportedFeaturePolicy,
            DynamicTargetGetter = request.DynamicTargetGetter,
            HeaderMatch = request.HeaderMatch,
            ValidateMode = request.ValidateMode,
            Culture = request.Culture,
            DynamicColumns = request.DynamicColumns,
            FailOnUnknownDynamicColumns = request.FailOnUnknownDynamicColumns,
            EnabledEmptyLine = request.EnabledEmptyLine,
            IgnoreEmptyLineAfterData = request.IgnoreEmptyLineAfterData,
            MappingConfiguration = request.MappingConfiguration,
            MappingDocument = request.MappingDocument,
            MappingPlan = mappingPlan,
            MaxTrackedUniqueValues = resourceLimits?.MaxTrackedUniqueValues,
            UniqueComparison = resourceLimits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase
        };
        var items = new List<TItem>();
        var rows = new List<int>();
        var sheetErrors = errors.CreateChild();
        ImportSheet(sheet, options, items, sheetErrors, runtime, cancellationToken, rows);
        var target = (ICollection<TItem>)request.Target(root);
        if (target == null)
            throw new InvalidOperationException($"Workbook 导入目标集合不可写入: {request.Name}");
        foreach (var item in items)
            target.Add(item);
        if (sourceLocations != null)
        {
            for (var index = 0; index < items.Count; index++)
                sourceLocations[items[index]] = new SourceLocation(sheet.SheetName, rows[index] + 1);
        }
        var result = new ExcelSheetImportResult(sheet.SheetName, typeof(TItem), rows,
            sheetErrors.Errors);
        sheetResults.Add(result);
    }

    /// <summary>
    /// 单遍处理工作表中的数据行。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="sheet">NPOI 工作表。</param>
    /// <param name="options">导入选项。</param>
    /// <param name="items">成功项集合。</param>
    /// <param name="errors">错误集合。</param>
    /// <param name="runtime">当前工作簿资源限制和行计数运行时。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="sourceRows">成功实体对应的 zero-based 原始行集合。</param>
    private void ImportSheet<T>(ISheet sheet, ExcelImportExecutionOptions<T> options, ICollection<T> items,
        ExcelImportErrorCollector errors, ExcelImportRuntime runtime, CancellationToken cancellationToken,
        ICollection<int> sourceRows = null)
        where T : class, new()
    {
        var header = sheet.GetRow(options.HeaderRowIndex)
            ?? throw new SheetStructureException("导入的模板不正确，未匹配表头。");
        if (header.LastCellNum > options.MaxColumnLength)
            throw new SheetStructureException($"导入表头超过最大列长度: {options.MaxColumnLength}");

        var columns = CreateColumns<T>(header, options);
        if (runtime.RowLimitReached)
            return;
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
                if (options.IgnoreEmptyLineAfterData)
                    break;
                if (options.EnabledEmptyLine)
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
                sheet.SheetName, rowIndex,
                options.BodyWhitespace, options.ValidateMode, options.UnsupportedFeaturePolicy, errors);
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
                    sheet.SheetName, rowIndex, options.ValidateMode, options.Culture, options.BodyWhitespace, errors))
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
                    rowIndex, options.ValidateMode, configuredValidationEnabled, errors, options.Culture,
                    options.BodyWhitespace, options.DynamicTargetGetter, imageIndex, out T item))
            {
                if (workbookValid)
                {
                    items.Add(item);
                    sourceRows?.Add(rowIndex);
                    if (configuredValidationEnabled)
                        uniqueTracker.CommitRow();
                }
                else if (configuredValidationEnabled)
                    uniqueTracker.RollbackRow();
            }
            else
            {
                if (configuredValidationEnabled)
                    uniqueTracker.RollbackRow();
            }
            if (errors.IsLimitReached)
            {
                errors.MarkTruncated();
                break;
            }
        }
    }

    private static bool IsConfiguredValidationEnabled(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.ConfiguredRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    /// <summary>
    /// 建立当前工作表的列绑定。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="header">表头行。</param>
    /// <param name="options">导入选项。</param>
    /// <returns>按列索引访问的属性绑定。</returns>
    private IReadOnlyDictionary<int, ExcelColumnPlan> CreateColumns<T>(IRow header, ExcelImportExecutionOptions<T> options)
        where T : class, new()
    {
        var map = options.MappingPlan;
        if (map == null)
        {
            map = _mappingPlanFactory.CreateWorkbook<T>(options.MappingDocument ?? new ExcelMappingDocument
            {
                UseConventionFallback = true
            },
                options.MappingConfiguration, MappingDirection.Import,
                new[] { header.Sheet.SheetName }).Sheets[0].Mapping;
        }
        var dynamicProperties = map.Columns.Where(property => property.IsDynamicColumn).ToList();
        var dynamicPlans = map.DynamicColumns;
        if (dynamicProperties.Count > 1)
            throw new InvalidOperationException($"导入模板 {typeof(T).FullName} 只能声明一个动态列属性。");

        var fixedProperties = map.Columns.Where(property => !property.Ignored && !property.IsDynamicColumn).ToList();
        var headerNames = new HashSet<string>(options.HeaderComparison == ExcelNameComparison.Ordinal
            ? StringComparer.Ordinal
            : StringComparer.OrdinalIgnoreCase);
        var columns = new Dictionary<int, ExcelColumnPlan>();
        foreach (var headerCell in header.Cells)
        {
            if (options.ReadColumnRange != null && !options.ReadColumnRange.Contains(headerCell.ColumnIndex))
                continue;
            var headerName = NormalizeText(GetRawStringValue(headerCell), options.HeaderWhitespace);
            if (string.IsNullOrWhiteSpace(headerName))
                continue;
            if (!headerNames.Add(headerName))
                throw new SheetStructureException($"导入的表格存在重复列:{headerName}");
            var property = FindProperty(fixedProperties, headerName, null, options.HeaderComparison);
            ExcelDynamicColumnDefinition dynamicDefinition = null;
            IExcelDynamicMappingColumn dynamicPlan = null;
            if (property == null && dynamicProperties.Count == 1)
            {
                dynamicPlan = FindDynamicDefinition(headerName, dynamicPlans, options.HeaderComparison);
                dynamicDefinition = dynamicPlan == null ? null : CreateDynamicDefinition(dynamicPlan);
                if (dynamicPlans.Count > 0 && dynamicDefinition == null)
                {
                    if (options.FailOnUnknownDynamicColumns)
                        throw new SheetStructureException($"导入包含未知动态列: {headerName}");
                    continue;
                }
                property = dynamicProperties[0];
            }
            if (property != null)
            {
                var isUnspecifiedDynamicColumn = property.IsDynamicColumn && dynamicDefinition == null;
                var reflectionProperty = typeof(T).GetProperty(property.Name,
                    BindingFlags.Instance | BindingFlags.Public);
                if (reflectionProperty == null)
                    throw new InvalidOperationException($"无法解析映射属性: {property.Name}");
                if (!property.IsDynamicColumn && !reflectionProperty.CanWrite)
                    throw new InvalidOperationException($"属性不可写入: {property.Name}");
                var valueConverters = isUnspecifiedDynamicColumn
                    ? Array.Empty<IExcelValueConverter>()
                    : property.IsDynamicColumn
                        ? dynamicPlan.ValueConverters
                        : property.ValueConverters;
                var validationBindings = property.IsDynamicColumn && dynamicPlan != null
                    ? dynamicPlan.ValidationBindings
                    : property.ValidationBindings;
                columns[headerCell.ColumnIndex] = new ExcelColumnPlan(headerName, property, property.IsDynamicColumn,
                    headerCell.ColumnIndex, dynamicDefinition, null, valueConverters, validationBindings,
                    reflectionProperty: reflectionProperty,
                    isUnique: dynamicPlan?.IsUnique,
                    uniqueIgnoreEmpty: dynamicPlan?.UniqueIgnoreEmpty ?? true);
            }
        }

        if (options.HeaderMatch)
        {
            var missing = fixedProperties.Where(property => !columns.Values.Any(column => column.Property == property)
                && (options.ReadColumnRange == null || !header.Cells.Any(cell =>
                    !options.ReadColumnRange.Contains(cell.ColumnIndex)
                        && (string.Equals(NormalizeText(GetRawStringValue(cell), options.HeaderWhitespace), property.Title,
                            ToStringComparison(options.HeaderComparison))
                            || property.Aliases.Any(alias => string.Equals(
                                NormalizeText(GetRawStringValue(cell), options.HeaderWhitespace), alias,
                                ToStringComparison(options.HeaderComparison))))))).ToList();
            if (missing.Any())
                throw new SheetStructureException($"导入的表格不存在列：{string.Join(",", missing.Select(property => property.Title))}");
        }
        return columns;
    }

    /// <summary>
    /// 按展示标题或历史别名查找 typed 动态列定义。
    /// </summary>
    private static IExcelDynamicMappingColumn FindDynamicDefinition(string headerName,
        IReadOnlyList<IExcelDynamicMappingColumn> definitions, ExcelNameComparison comparison)
    {
        return (definitions ?? Array.Empty<IExcelDynamicMappingColumn>()).FirstOrDefault(definition =>
            string.Equals(definition.Title, headerName, ToStringComparison(comparison))
            || (definition.Aliases ?? Array.Empty<string>()).Any(alias =>
                string.Equals(alias, headerName, ToStringComparison(comparison))));
    }

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

    private static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var separator = placementKey.IndexOfAny(new[] { ':', '-' });
        var key = placementKey.Substring(separator + 1);
        return placementKey.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || placementKey.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key)
            : ExcelColumnPlacement.After(key);
    }

    private static Type ResolveDynamicType(string name)
    {
        switch ((name ?? "string").ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "string": return typeof(string);
            case "boolean": case "bool": return typeof(bool);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int32": case "int": return typeof(int);
            case "int64": case "long": return typeof(long);
            case "single": case "float": return typeof(float);
            case "double": return typeof(double);
            case "decimal": return typeof(decimal);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            default: throw new InvalidOperationException($"动态列数据类型不在允许列表中: {name}");
        }
    }

    /// <summary>
    /// 验证请求级表头映射不会覆盖动态列绑定规则。
    /// </summary>
    /// <param name="headerMappings">属性名称到表头名称的请求级映射。</param>
    /// <param name="fixedProperties">可映射的固定属性集合。</param>
    /// <param name="dynamicProperties">动态列属性集合。</param>
    private static void ValidateHeaderMappings(IReadOnlyDictionary<string, string> headerMappings,
        IReadOnlyCollection<IExcelMappingColumn> fixedProperties, IReadOnlyCollection<IExcelMappingColumn> dynamicProperties)
    {
        if (headerMappings == null)
            return;
        var propertyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var headerNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mapping in headerMappings)
        {
            if (string.IsNullOrWhiteSpace(mapping.Key) || string.IsNullOrWhiteSpace(mapping.Value))
                throw new InvalidOperationException("导入模板表头映射的属性名称和表头名称不能为空。");
            if (dynamicProperties.Any(property => string.Equals(property.Name, mapping.Key,
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"导入模板不允许为动态列属性配置表头映射: {mapping.Key}");
            if (!fixedProperties.Any(property => string.Equals(property.Name, mapping.Key,
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"导入模板表头映射引用了不存在或已忽略的属性: {mapping.Key}");
            if (!propertyNames.Add(mapping.Key))
                throw new InvalidOperationException($"导入模板表头映射包含重复的属性: {mapping.Key}");
            if (!headerNames.Add(mapping.Value))
                throw new InvalidOperationException($"导入模板表头映射包含重复的表头: {mapping.Value}");
        }
    }

    /// <summary>
    /// 根据请求映射和默认标题查找固定属性。
    /// </summary>
    /// <param name="properties">固定属性集合。</param>
    /// <param name="headerName">表头名称。</param>
    /// <param name="headerMappings">请求级映射。</param>
    /// <param name="comparison">表头名称比较规则。</param>
    private static IExcelMappingColumn FindProperty(IEnumerable<IExcelMappingColumn> properties, string headerName,
        IReadOnlyDictionary<string, string> headerMappings, ExcelNameComparison comparison)
    {
        var stringComparison = ToStringComparison(comparison);
        foreach (var property in properties)
        {
            var mappedHeader = headerMappings?.FirstOrDefault(mapping =>
                string.Equals(mapping.Key, property.Name, StringComparison.OrdinalIgnoreCase)).Value;
            if (!string.IsNullOrWhiteSpace(mappedHeader)
                && string.Equals(mappedHeader, headerName, stringComparison))
                return property;
            if (string.Equals(property.Title, headerName, stringComparison)
                || property.Aliases.Any(alias => string.Equals(alias, headerName, stringComparison))
                || string.Equals(property.Name, headerName, stringComparison))
                return property;
        }
        return null;
    }

    private static StringComparison ToStringComparison(ExcelNameComparison comparison) =>
        comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    internal static string GetRawStringValue(ICell cell)
    {
        if (cell == null)
            return string.Empty;
        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        return cellType == CellType.String ? cell.StringCellValue ?? string.Empty : cell.GetStringValue();
    }

    internal static string NormalizeText(string value, ExcelWhitespacePolicy policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.Trim => value.Trim(),
            ExcelWhitespacePolicy.RemoveAll => RemoveWhitespace(value),
            _ => throw new ArgumentOutOfRangeException(nameof(policy))
        };
    }

    private static string RemoveWhitespace(string value)
    {
        var hasWhitespace = value.Any(char.IsWhiteSpace);
        if (!hasWhitespace)
            return value;
        var buffer = new char[value.Length];
        var length = 0;
        foreach (var character in value)
        {
            if (!char.IsWhiteSpace(character))
                buffer[length++] = character;
        }
        return new string(buffer, 0, length);
    }

    /// <summary>
    /// 读取不依赖 NPOI 的单元格值描述。
    /// </summary>
    /// <param name="cell">待读取的单元格。</param>
    /// <returns>用于转换器和默认转换的单元格值描述。</returns>
    internal static ExcelCellValue ReadCellValue(ICell cell)
    {
        if (cell == null)
            return new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty);

        var isFormula = cell.CellType == CellType.Formula;
        var effectiveType = isFormula ? cell.CachedFormulaResultType : cell.CellType;
        var effectiveKind = ResolveCellKind(effectiveType, cell);
        object value = effectiveType switch
        {
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => cell.DateCellValue,
            CellType.Numeric => cell.NumericCellValue,
            CellType.Error => null,
            _ => GetRawStringValue(cell)
        };
        return new ExcelCellValue(value, GetRawStringValue(cell), isFormula ? ExcelCellKind.Formula : effectiveKind,
            isFormula ? effectiveKind : null, isFormula ? cell.CellFormula : null,
            effectiveType == CellType.Error ? cell.ErrorCellValue : null, cell.CellStyle?.DataFormat);
    }

    /// <summary>
    /// 映射 NPOI 单元格类型到提供程序无关的逻辑类型。
    /// </summary>
    private static ExcelCellKind ResolveCellKind(CellType cellType, ICell cell) => cellType switch
    {
        CellType.Blank => ExcelCellKind.Empty,
        CellType.String => ExcelCellKind.Text,
        CellType.Boolean => ExcelCellKind.Boolean,
        CellType.Error => ExcelCellKind.Error,
        CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => ExcelCellKind.DateTime,
        CellType.Numeric => ExcelCellKind.Number,
        _ => ExcelCellKind.Text
    };

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

    private sealed class SheetStructureException : InvalidOperationException
    {
        public SheetStructureException(string message) : base(message)
        {
        }
    }
}

internal sealed class ExcelImportRuntime
{
    private int _rowCount;
    private bool _rowLimitReported;

    internal ExcelImportRuntime(ExcelResourceLimits limits)
    {
        MaxRows = limits?.MaxRows;
        ImageResources = new ExcelImageResourceTracker(limits);
    }

    internal int? MaxRows { get; }

    internal bool RowLimitReached => MaxRows.HasValue && _rowCount >= MaxRows.Value;

    internal ExcelImageResourceTracker ImageResources { get; }

    internal bool TryConsumeRow()
    {
        if (MaxRows.HasValue && _rowCount >= MaxRows.Value)
            return false;
        _rowCount++;
        return true;
    }

    internal bool TryMarkRowLimitReported()
    {
        if (_rowLimitReported)
            return false;
        _rowLimitReported = true;
        return true;
    }
}

internal sealed class ExcelImageResourceTracker
{
    private readonly int? _maxPictures;
    private readonly long? _maxPictureBytes;
    private readonly long? _maxTotalPictureBytes;
    private int _count;
    private long _totalBytes;

    internal ExcelImageResourceTracker(ExcelResourceLimits limits)
    {
        _maxPictures = limits?.MaxPictures;
        _maxPictureBytes = limits?.MaxPictureBytes;
        _maxTotalPictureBytes = limits?.MaxTotalPictureBytes;
    }

    internal void Consume(long bytes)
    {
        if (_maxPictureBytes.HasValue && bytes > _maxPictureBytes.Value)
            throw new ImageResourceLimitException($"单张图片超过最大字节数: {_maxPictureBytes.Value}");
        if (_maxPictures.HasValue && _count >= _maxPictures.Value)
            throw new ImageResourceLimitException($"图片数量超过限制: {_maxPictures.Value}");
        if (_maxTotalPictureBytes.HasValue && bytes > _maxTotalPictureBytes.Value - _totalBytes)
            throw new ImageResourceLimitException($"图片总字节数超过限制: {_maxTotalPictureBytes.Value}");
        _count++;
        _totalBytes += bytes;
    }
}

internal sealed class ImageResourceLimitException : InvalidOperationException
{
    internal ImageResourceLimitException(string message) : base(message)
    {
    }
}

internal sealed class SourceLocation
{
    internal SourceLocation(string sheetName, int rowIndex)
    {
        SheetName = sheetName;
        RowIndex = rowIndex;
    }

    internal string SheetName { get; }
    internal int RowIndex { get; }
}

internal sealed class ReferenceObjectComparer : IEqualityComparer<object>
{
    internal static readonly ReferenceObjectComparer Instance = new();

    public new bool Equals(object x, object y) => ReferenceEquals(x, y);

    public int GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}

#pragma warning restore CS0618
