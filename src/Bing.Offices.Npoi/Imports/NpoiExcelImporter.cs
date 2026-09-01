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
/// 基于 NPOI 的 Excel 导入器；输入会先复制并由 NPOI 建立内存中的 Workbook DOM。
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
    /// <summary>
    /// 将请求级配置编译为提供程序无关映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>
    /// 将 Workbook 导入请求展开为按工作表执行的导入计划生成器。
    /// </summary>
    private readonly NpoiImportPlanBuilder _planBuilder;

    /// <summary>
    /// 当前导入器使用的命名配置校验规则。
    /// </summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    /// <summary>
    /// 按行校验并物化导入实体的执行器。
    /// </summary>
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
        var resolvedSheets = request.Sheets.Select(sheet => ResolveSheet(workbook, sheet,
            request.SheetNameComparison)).ToArray();
        var existingSheets = resolvedSheets.Where(sheet => sheet.Exists).ToArray();
        var duplicatePhysicalSheet = existingSheets.GroupBy(item => item.Index)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicatePhysicalSheet != null)
        {
            var physicalIndex = duplicatePhysicalSheet.Key;
            var physicalName = duplicatePhysicalSheet.First().Name;
            var selectors = string.Join(", ", duplicatePhysicalSheet.Select(item =>
                GetSelectorDescription(item.Request.Selector)));
            throw new ArgumentException(
                $"多个 Sheet selector 指向同一物理 Sheet: {selectors}; 实际 Sheet=#{physicalIndex} {physicalName}");
        }
        var plans = _planBuilder.Create(existingSheets);
        var resolvedSheetRequests = resolvedSheets.Where(sheet => sheet.Exists &&
                !workbook.IsSheetHidden(sheet.Index) && !workbook.IsSheetVeryHidden(sheet.Index))
            .ToDictionary(sheet => sheet.Name, sheet => sheet.Request, StringComparer.OrdinalIgnoreCase);
        foreach (var resolvedSheet in resolvedSheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached || runtime.RowLimitReached)
                break;
            var sheetRequest = resolvedSheet.Request;
            if (!resolvedSheet.Exists)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少请求的 Sheet: {GetSelectorDescription(sheetRequest.Selector)}",
                    GetSelectorDescription(sheetRequest.Selector), 0, 0, null));
                continue;
            }
            if (workbook.IsSheetHidden(resolvedSheet.Index) || workbook.IsSheetVeryHidden(resolvedSheet.Index))
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"请求的 Sheet 被隐藏: {resolvedSheet.Name}", resolvedSheet.Name,
                    0, 0, null));
                continue;
            }
            var sheet = workbook.GetSheetAt(resolvedSheet.Index);
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

    /// <summary>根据选择器查找工作表的零基物理索引。</summary>
    /// <param name="workbook">待查找的工作簿。</param>
    /// <param name="selector">按名称或索引定位工作表的选择器。</param>
    /// <param name="comparison">工作表名称比较规则。</param>
    /// <returns>匹配的零基工作表索引；未找到时为 -1。</returns>
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

    /// <summary>解析一个工作表请求并保存物理名称，供后续所有导入阶段复用。</summary>
    /// <param name="workbook">待查找的工作簿。</param>
    /// <param name="request">待解析的工作表请求。</param>
    /// <param name="comparison">工作表名称比较规则。</param>
    /// <returns>已解析的工作表执行描述。</returns>
    private static NpoiResolvedSheet ResolveSheet(IWorkbook workbook, ExcelSheetImportRequest request,
        ExcelNameComparison comparison)
    {
        var index = ResolveSheetIndex(workbook, request.Selector, comparison);
        return new NpoiResolvedSheet(request, index, index < 0 ? null : workbook.GetSheetName(index));
    }

    /// <summary>生成用于错误消息的工作表选择器描述。</summary>
    /// <param name="selector">待描述的工作表选择器。</param>
    /// <returns>索引选择器的井号形式或名称选择器的名称。</returns>
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
    /// 按行处理已由 NPOI 加载到内存的工作表数据。
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

    /// <summary>确定导入模式是否要求执行配置的字段校验规则。</summary>
    /// <param name="mode">当前导入校验模式。</param>
    /// <returns>应执行配置校验规则时为 true。</returns>
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

    /// <summary>将已绑定的动态列映射转换为请求级动态列定义。</summary>
    /// <param name="column">提供程序无关的动态列映射。</param>
    /// <returns>可供行物化器使用的动态列定义。</returns>
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

    /// <summary>将 before/after 形式的位置键转换为动态列定位规则。</summary>
    /// <param name="placementKey">映射中保存的相对位置键。</param>
    /// <returns>解析后的定位规则；键为空时为 null。</returns>
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

    /// <summary>解析配置允许的动态列逻辑类型名称。</summary>
    /// <param name="name">配置中的类型名称；为空时使用 string。</param>
    /// <returns>对应的 CLR 类型。</returns>
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

    /// <summary>将领域名称比较枚举转换为字符串比较规则。</summary>
    /// <param name="comparison">领域名称比较规则。</param>
    /// <returns>等效的字符串比较规则。</returns>
    private static StringComparison ToStringComparison(ExcelNameComparison comparison) =>
        comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    /// <summary>读取单元格的原始文本值，并优先读取公式的缓存结果。</summary>
    /// <param name="cell">待读取的单元格。</param>
    /// <returns>单元格文本；单元格为空时为空字符串。</returns>
    internal static string GetRawStringValue(ICell cell)
    {
        if (cell == null)
            return string.Empty;
        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        return cellType == CellType.String ? cell.StringCellValue ?? string.Empty : cell.GetStringValue();
    }

    /// <summary>按照指定空白策略规范化单元格文本。</summary>
    /// <param name="value">待规范化的文本；为 null 时按空字符串处理。</param>
    /// <param name="policy">空白字符处理策略。</param>
    /// <returns>规范化后的文本。</returns>
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

    /// <summary>移除文本中的所有 Unicode 空白字符。</summary>
    /// <param name="value">待处理的文本。</param>
    /// <returns>移除空白后的文本；原文本无空白时直接返回原引用。</returns>
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

    /// <summary>创建与指定字符串比较规则等效的字符串比较器。</summary>
    /// <param name="comparison">要支持的字符串比较规则。</param>
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

    private sealed class SheetStructureException : InvalidOperationException
    {
        /// <summary>使用工作表结构错误消息初始化异常。</summary>
        /// <param name="message">描述无法导入的表结构问题的消息。</param>
        public SheetStructureException(string message) : base(message)
        {
        }
    }
}

internal sealed class ExcelImportRuntime
{
    /// <summary>当前工作簿已尝试处理的数据行数量。</summary>
    private int _rowCount;
    /// <summary>指示行数上限错误是否已加入结果，避免重复报告。</summary>
    private bool _rowLimitReported;

    /// <summary>使用请求资源限制初始化工作簿级运行时状态。</summary>
    /// <param name="limits">导入请求配置的资源限制。</param>
    internal ExcelImportRuntime(ExcelResourceLimits limits)
    {
        MaxRows = limits?.MaxRows;
        ImageResources = new ExcelImageResourceTracker(limits);
    }

    /// <summary>获取允许处理的最大数据行数。</summary>
    internal int? MaxRows { get; }

    /// <summary>获取是否已达到数据行数上限。</summary>
    internal bool RowLimitReached => MaxRows.HasValue && _rowCount >= MaxRows.Value;

    /// <summary>获取跟踪工作簿图片数量和字节数的资源限制器。</summary>
    internal ExcelImageResourceTracker ImageResources { get; }

    /// <summary>尝试为一行数据消耗全局行数配额。</summary>
    /// <returns>成功消耗配额时为 true；已达到上限时为 false。</returns>
    internal bool TryConsumeRow()
    {
        if (MaxRows.HasValue && _rowCount >= MaxRows.Value)
            return false;
        _rowCount++;
        return true;
    }

    /// <summary>确保行数上限错误在单个工作簿中仅报告一次。</summary>
    /// <returns>本次调用首次标记上限错误时为 true。</returns>
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
    /// <summary>工作簿允许读取的最大图片数量。</summary>
    private readonly int? _maxPictures;
    /// <summary>单张图片允许占用的最大字节数。</summary>
    private readonly long? _maxPictureBytes;
    /// <summary>所有图片合计允许占用的最大字节数。</summary>
    private readonly long? _maxTotalPictureBytes;
    /// <summary>已接纳的图片数量。</summary>
    private int _count;
    /// <summary>已接纳图片的累计字节数。</summary>
    private long _totalBytes;

    /// <summary>从导入请求资源限制初始化图片配额跟踪器。</summary>
    /// <param name="limits">导入请求配置的资源限制。</param>
    internal ExcelImageResourceTracker(ExcelResourceLimits limits)
    {
        _maxPictures = limits?.MaxPictures;
        _maxPictureBytes = limits?.MaxPictureBytes;
        _maxTotalPictureBytes = limits?.MaxTotalPictureBytes;
    }

    /// <summary>验证并记录一张图片对工作簿资源配额的消耗。</summary>
    /// <param name="bytes">待接纳图片的字节数。</param>
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
    /// <summary>使用图片资源限制错误消息初始化异常。</summary>
    /// <param name="message">描述超出图片资源限制的消息。</param>
    internal ImageResourceLimitException(string message) : base(message)
    {
    }
}

internal sealed class SourceLocation
{
    /// <summary>使用导入实体的来源工作表和行号初始化位置。</summary>
    /// <param name="sheetName">实体所属的工作表名称。</param>
    /// <param name="rowIndex">实体所在的一基数据行号。</param>
    internal SourceLocation(string sheetName, int rowIndex)
    {
        SheetName = sheetName;
        RowIndex = rowIndex;
    }

    /// <summary>获取实体所属的工作表名称。</summary>
    internal string SheetName { get; }
    /// <summary>获取实体所在的一基数据行号。</summary>
    internal int RowIndex { get; }
}

/// <summary>按照对象引用而非对象值比较关联实体的比较器。</summary>
internal sealed class ReferenceObjectComparer : IEqualityComparer<object>
{
    /// <summary>获取可复用的引用比较器实例。</summary>
    internal static readonly ReferenceObjectComparer Instance = new();

    /// <inheritdoc />
    public new bool Equals(object x, object y) => ReferenceEquals(x, y);

    /// <inheritdoc />
    public int GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}

#pragma warning restore CS0618
