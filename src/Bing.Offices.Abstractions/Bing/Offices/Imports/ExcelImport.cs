using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 级强类型导入请求构建入口。
/// </summary>
public static class ExcelImport
{
    /// <summary>
    /// 创建 Workbook 导入请求。
    /// </summary>
    public static ExcelWorkbookImportRequest<TWorkbook> Workbook<TWorkbook>(
        Action<ExcelWorkbookImportBuilder<TWorkbook>> configure)
        where TWorkbook : class, new()
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelWorkbookImportBuilder<TWorkbook>();
        configure(builder);
        return builder.Build();
    }
}

/// <summary>
/// Workbook 导入构建器。
/// </summary>
public sealed class ExcelWorkbookImportBuilder<TWorkbook> where TWorkbook : class, new()
{
    private readonly List<ExcelSheetImportRequest> _sheets = new List<ExcelSheetImportRequest>();
    private readonly List<ExcelRelationRequest> _relations = new List<ExcelRelationRequest>();
    private ExcelNameComparison _sheetNameComparison = ExcelNameComparison.OrdinalIgnoreCase;
    private ExcelResourceLimits _resourceLimits;
    private ExcelImportFailureOptions _failureOptions;
    private ExcelImportValidationMode _validationMode = ExcelImportValidationMode.ConfiguredRules;
    private ExcelUnsupportedFeaturePolicy _unsupportedFeaturePolicy = ExcelUnsupportedFeaturePolicy.Fail;

    /// <summary>
    /// 设置按名称选择 Sheet 时的名称比较策略。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> SheetNameComparison(ExcelNameComparison comparison)
    {
        _sheetNameComparison = comparison;
        return this;
    }

    /// <summary>
    /// 设置导入资源上限。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> ResourceLimits(ExcelResourceLimits limits)
    {
        _resourceLimits = limits;
        return this;
    }

    /// <summary>
    /// 设置失败工作簿输出。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> FailureWorkbook(ExcelImportFailureOptions options)
    {
        _failureOptions = options;
        return this;
    }

    /// <summary>
    /// 设置工作簿原生 Data Validation 规则处理模式。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> ValidationMode(ExcelImportValidationMode mode)
    {
        _validationMode = mode;
        return this;
    }

    /// <summary>
    /// 设置 Workbook 原生校验规则不支持时的处理策略。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> UnsupportedFeaturePolicy(ExcelUnsupportedFeaturePolicy policy)
    {
        _unsupportedFeaturePolicy = policy;
        return this;
    }

    /// <summary>
    /// 添加一个强类型 Sheet 导入配置。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> Sheet<TItem>(string name,
        Expression<Func<TWorkbook, ICollection<TItem>>> target,
        Action<ExcelSheetImportBuilder<TItem>> configure = null)
        where TItem : class, new()
    {
        return Sheet(ExcelSheetSelector.ByName(name), target, configure);
    }

    /// <summary>
    /// 添加一个按名称或索引选择的强类型 Sheet 导入配置。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> Sheet<TItem>(ExcelSheetSelector selector,
        Expression<Func<TWorkbook, ICollection<TItem>>> target,
        Action<ExcelSheetImportBuilder<TItem>> configure = null)
        where TItem : class, new()
    {
        if (selector == null)
            throw new ArgumentNullException(nameof(selector));
        if (target == null)
            throw new ArgumentNullException(nameof(target));
        var builder = new ExcelSheetImportBuilder<TItem>(selector, target);
        configure?.Invoke(builder);
        _sheets.Add(builder.Build());
        return this;
    }

    /// <summary>
    /// 添加显式父子集合关系。
    /// </summary>
    public ExcelWorkbookImportBuilder<TWorkbook> HasMany<TParent, TChild, TKey>(
        Expression<Func<TWorkbook, ICollection<TParent>>> parents,
        Expression<Func<TWorkbook, ICollection<TChild>>> children,
        Func<TParent, TKey> parentKey,
        Func<TChild, TKey> childKey,
        Expression<Func<TParent, ICollection<TChild>>> navigation,
        IEqualityComparer<TKey> comparer = null)
        where TParent : class
        where TChild : class
    {
        if (parents == null || children == null || parentKey == null || childKey == null || navigation == null)
            throw new ArgumentNullException("关系配置参数不能为空。");
        _relations.Add(ExcelRelationRequest.Create(parents, children, parentKey, childKey, navigation, comparer));
        return this;
    }

    internal ExcelWorkbookImportRequest<TWorkbook> Build()
    {
        if (_sheets.Count == 0)
            throw new InvalidOperationException("Workbook 至少需要一个 Sheet。");
        if (!Enum.IsDefined(typeof(ExcelNameComparison), _sheetNameComparison))
            throw new ArgumentOutOfRangeException(nameof(_sheetNameComparison));
        _resourceLimits?.Validate();
        _failureOptions?.Validate();
        if (!Enum.IsDefined(typeof(ExcelImportValidationMode), _validationMode))
            throw new ArgumentOutOfRangeException(nameof(_validationMode));
        if (!Enum.IsDefined(typeof(ExcelUnsupportedFeaturePolicy), _unsupportedFeaturePolicy))
            throw new ArgumentOutOfRangeException(nameof(_unsupportedFeaturePolicy));
        var names = new HashSet<string>(_sheetNameComparison == ExcelNameComparison.Ordinal
            ? StringComparer.Ordinal
            : StringComparer.OrdinalIgnoreCase);
        var indexes = new HashSet<int>();
        foreach (var sheet in _sheets)
        {
            if (sheet.Selector.Kind == ExcelSheetSelectorKind.ByName && !names.Add(sheet.Selector.Name))
                throw new ArgumentException($"Workbook 包含重复 Sheet selector: {sheet.Selector.Name}");
            if (sheet.Selector.Kind == ExcelSheetSelectorKind.ByIndex
                && !indexes.Add(sheet.Selector.Index.Value))
                throw new ArgumentException($"Workbook 包含重复 Sheet selector: #{sheet.Selector.Index.Value}");
        }
        return new ExcelWorkbookImportRequest<TWorkbook>(_sheets.AsReadOnly(), _relations.AsReadOnly(),
            _sheetNameComparison, _resourceLimits, _failureOptions, _validationMode, _unsupportedFeaturePolicy);
    }
}

/// <summary>
/// 单个 Sheet 导入构建器。
/// </summary>
public sealed class ExcelSheetImportBuilder<TItem> where TItem : class, new()
{
    private readonly string _name;
    private readonly Expression _target;
    private readonly ExcelSheetSelector _selector;
    private int _headerRowIndex;
    private int _dataRowStartIndex = 1;
    private IReadOnlyList<Exports.ExcelDynamicColumnDefinition> _dynamicColumns =
        Array.Empty<Exports.ExcelDynamicColumnDefinition>();
    private bool _requireExpectedHeaders = true;
    private ValidateMode _validateMode = ValidateMode.StopOnFirstFailure;
    private System.Globalization.CultureInfo _culture = System.Globalization.CultureInfo.InvariantCulture;
    private Configurations.ExcelMappingConfiguration _requestMappingConfiguration;
    private Configurations.ExcelMappingDocument _mappingDocument;
    private Expression<Func<TItem, IDictionary<string, object>>> _dynamicTarget;
    private int _maxReadColumns = 100;
    private ExcelReadColumnRange _readColumnRange;
    private ExcelNameComparison _headerComparison = ExcelNameComparison.OrdinalIgnoreCase;
    private ExcelWhitespacePolicy _headerWhitespace = ExcelWhitespacePolicy.Trim;
    private ExcelWhitespacePolicy _bodyWhitespace = ExcelWhitespacePolicy.Trim;
    private bool _failOnUnknownDynamicColumns;
    private bool _reportEmptyRows;
    private bool _stopAtFirstEmptyRow;

    internal ExcelSheetImportBuilder(ExcelSheetSelector selector, Expression target)
    {
        _selector = selector;
        _name = selector.Name ?? $"#{selector.Index}";
        _target = target;
    }

    /// <summary>
    /// 设置表头行索引，索引从零开始。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> HeaderRowIndex(int index)
    {
        _headerRowIndex = index;
        if (_dataRowStartIndex == 1)
            _dataRowStartIndex = index + 1;
        return this;
    }

    /// <summary>
    /// 设置正文起始行索引，索引从零开始。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> DataRowStartIndex(int index)
    {
        _dataRowStartIndex = index;
        return this;
    }

    /// <summary>
    /// 配置与导出相同的动态列定义。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> DynamicColumns(
        Expression<Func<TItem, IDictionary<string, object>>> target,
        IReadOnlyList<Exports.ExcelDynamicColumnDefinition> definitions)
    {
        _dynamicTarget = target ?? throw new ArgumentNullException(nameof(target));
        _dynamicColumns = definitions ?? throw new ArgumentNullException(nameof(definitions));
        return this;
    }

    /// <summary>
    /// 设置是否要求固定列全部存在。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> RequireExpectedHeaders(bool value)
    {
        _requireExpectedHeaders = value;
        return this;
    }

    /// <summary>
    /// 设置最大表头列数安全上限。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> MaxReadColumns(int value)
    {
        _maxReadColumns = value;
        return this;
    }

    /// <summary>
    /// 设置实际参与绑定的列读取范围。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> ReadColumns(int startIndex, int count)
    {
        _readColumnRange = ExcelReadColumnRange.Create(startIndex, count);
        return this;
    }

    /// <summary>
    /// 设置表头名称比较策略。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> HeaderComparison(ExcelNameComparison comparison)
    {
        _headerComparison = comparison;
        return this;
    }

    /// <summary>
    /// 设置表头文本空白规范化策略。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> HeaderWhitespace(ExcelWhitespacePolicy policy)
    {
        _headerWhitespace = policy;
        return this;
    }

    /// <summary>
    /// 设置正文文本空白规范化策略。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> BodyWhitespace(ExcelWhitespacePolicy policy)
    {
        _bodyWhitespace = policy;
        return this;
    }

    /// <summary>
    /// 设置未知动态表头是否导致当前 Sheet 失败。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> FailOnUnknownDynamicColumns(bool value = true)
    {
        _failOnUnknownDynamicColumns = value;
        return this;
    }

    /// <summary>
    /// 设置是否报告空数据行。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> ReportEmptyRows(bool value = true)
    {
        _reportEmptyRows = value;
        return this;
    }

    /// <summary>
    /// 设置是否在首个空行后停止读取。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> StopAtFirstEmptyRow(bool value = true)
    {
        _stopAtFirstEmptyRow = value;
        return this;
    }

    /// <summary>
    /// 设置校验失败处理粒度。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> Validate(ValidateMode mode)
    {
        _validateMode = mode;
        return this;
    }

    /// <summary>
    /// 设置数字和日期转换区域性。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> Culture(System.Globalization.CultureInfo culture)
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        return this;
    }

    /// <summary>
    /// 设置请求级映射配置。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> Mapping(Configurations.ExcelMappingConfiguration configuration)
    {
        _requestMappingConfiguration = configuration == null ? null :
            Configurations.MappingConfigurationCloner.Clone(configuration, Configurations.MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置规范化映射文档的导入方向配置。
    /// </summary>
    public ExcelSheetImportBuilder<TItem> Mapping(Configurations.ExcelMappingDocument document)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        _mappingDocument = Configurations.MappingDocumentCloner.Clone(document);
        return this;
    }

    internal ExcelSheetImportRequest Build()
    {
        if (_maxReadColumns <= 0)
            throw new ArgumentOutOfRangeException(nameof(_maxReadColumns));
        if (_headerRowIndex < 0 || _dataRowStartIndex < 0 || _dataRowStartIndex <= _headerRowIndex)
            throw new ArgumentOutOfRangeException(nameof(_dataRowStartIndex));
        if (!Enum.IsDefined(typeof(ValidateMode), _validateMode))
            throw new ArgumentOutOfRangeException(nameof(_validateMode));
        if (!Enum.IsDefined(typeof(ExcelNameComparison), _headerComparison))
            throw new ArgumentOutOfRangeException(nameof(_headerComparison));
        if (!Enum.IsDefined(typeof(ExcelWhitespacePolicy), _headerWhitespace))
            throw new ArgumentOutOfRangeException(nameof(_headerWhitespace));
        if (!Enum.IsDefined(typeof(ExcelWhitespacePolicy), _bodyWhitespace))
            throw new ArgumentOutOfRangeException(nameof(_bodyWhitespace));
        if (_culture == null)
            throw new ArgumentNullException(nameof(_culture));
        var compiledTarget = ((LambdaExpression)_target).Compile();
        Func<object, object> targetGetter = value => compiledTarget.DynamicInvoke(value);
        Func<object, object> dynamicGetter = null;
        if (_dynamicTarget != null)
        {
            var compiledDynamicTarget = _dynamicTarget.Compile();
            dynamicGetter = value => compiledDynamicTarget((TItem)value);
        }
        var requestConfiguration = Exports.ExcelDynamicColumnCloner.MergeIntoConfiguration(
            _requestMappingConfiguration, _dynamicColumns);
        return new ExcelSheetImportRequest(_name, _selector, typeof(TItem), targetGetter,
        _headerRowIndex, _dataRowStartIndex, Exports.ExcelDynamicColumnCloner.Clone(_dynamicColumns), _dynamicTarget, _requireExpectedHeaders, _validateMode, _culture,
        requestConfiguration, _mappingDocument,
        dynamicGetter,
        _maxReadColumns,
        _failOnUnknownDynamicColumns, _reportEmptyRows, _stopAtFirstEmptyRow, _readColumnRange,
        _headerComparison, _headerWhitespace, _bodyWhitespace);
    }
}
