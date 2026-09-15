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
    /// <typeparam name="TWorkbook">根 Workbook 模型类型。</typeparam>
    /// <param name="configure">用于配置 Workbook 导入选项的委托。</param>
    /// <returns>已完成配置的 Workbook 导入请求。</returns>
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
/// <typeparam name="TWorkbook">根 Workbook 模型类型。</typeparam>
public sealed class ExcelWorkbookImportBuilder<TWorkbook> where TWorkbook : class, new()
{
    /// <summary>按配置顺序保存待导入的工作表请求。</summary>
    private readonly List<ExcelSheetImportRequest> _sheets = new List<ExcelSheetImportRequest>();
    /// <summary>保存 Workbook 级关系绑定请求，构建完成时转换为只读快照。</summary>
    private readonly List<ExcelRelationRequest> _relations = new List<ExcelRelationRequest>();
    /// <summary>按名称选择工作表时使用的比较策略，默认忽略大小写。</summary>
    private ExcelNameComparison _sheetNameComparison = ExcelNameComparison.OrdinalIgnoreCase;
    /// <summary>导入过程的资源限制；未设置时沿用 Provider 默认值。</summary>
    private ExcelResourceLimits _resourceLimits;
    /// <summary>失败工作簿输出选项；未设置时不生成失败输出。</summary>
    private ExcelImportFailureOptions _failureOptions;
    /// <summary>导入校验模式，默认执行已配置的校验规则。</summary>
    private ExcelImportValidationMode _validationMode = ExcelImportValidationMode.ConfiguredRules;
    /// <summary>不支持功能的处理策略，默认遇到不支持功能时失败。</summary>
    private ExcelUnsupportedFeaturePolicy _unsupportedFeaturePolicy = ExcelUnsupportedFeaturePolicy.Fail;

    /// <summary>
    /// 设置按名称选择 Sheet 时的名称比较策略。
    /// </summary>
    /// <param name="comparison">工作表名称的比较策略。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
    public ExcelWorkbookImportBuilder<TWorkbook> SheetNameComparison(ExcelNameComparison comparison)
    {
        _sheetNameComparison = comparison;
        return this;
    }

    /// <summary>
    /// 设置导入资源上限。
    /// </summary>
    /// <param name="limits">导入过程中使用的资源限制配置。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
    public ExcelWorkbookImportBuilder<TWorkbook> ResourceLimits(ExcelResourceLimits limits)
    {
        _resourceLimits = limits;
        return this;
    }

    /// <summary>
    /// 设置失败工作簿输出。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
    public ExcelWorkbookImportBuilder<TWorkbook> FailureWorkbook(ExcelImportFailureOptions options)
    {
        _failureOptions = options;
        return this;
    }

    /// <summary>
    /// 设置工作簿原生 Data Validation 规则处理模式。
    /// </summary>
    /// <param name="mode">原生校验规则的处理模式。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
    public ExcelWorkbookImportBuilder<TWorkbook> ValidationMode(ExcelImportValidationMode mode)
    {
        _validationMode = mode;
        return this;
    }

    /// <summary>
    /// 设置 Workbook 原生校验规则不支持时的处理策略。
    /// </summary>
    /// <param name="policy">遇到不支持的原生校验规则时采用的策略。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
    public ExcelWorkbookImportBuilder<TWorkbook> UnsupportedFeaturePolicy(ExcelUnsupportedFeaturePolicy policy)
    {
        _unsupportedFeaturePolicy = policy;
        return this;
    }

    /// <summary>
    /// 添加一个强类型 Sheet 导入配置。
    /// </summary>
    /// <typeparam name="TItem">Sheet 明细模型类型。</typeparam>
    /// <param name="name">工作表名称。</param>
    /// <param name="target">Workbook 中接收导入结果的明细集合属性。</param>
    /// <param name="configure">用于配置当前 Sheet 的可选委托。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
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
    /// <typeparam name="TItem">Sheet 明细模型类型。</typeparam>
    /// <param name="selector">选择目标工作表的名称或索引选择器。</param>
    /// <param name="target">Workbook 中接收导入结果的明细集合属性。</param>
    /// <param name="configure">用于配置当前 Sheet 的可选委托。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
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
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">子实体类型。</typeparam>
    /// <typeparam name="TKey">父子实体关联键的类型。</typeparam>
    /// <param name="parents">Workbook 中的父实体集合属性。</param>
    /// <param name="children">Workbook 中的子实体集合属性。</param>
    /// <param name="parentKey">从父实体获取关联键的函数。</param>
    /// <param name="childKey">从子实体获取关联键的函数。</param>
    /// <param name="navigation">父实体上的子实体导航属性。</param>
    /// <param name="comparer">比较关联键的可选比较器。</param>
    /// <returns>当前 Workbook 构建器，用于继续配置。</returns>
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

    /// <summary>构建并校验当前 Workbook 导入请求。</summary>
    /// <returns>不可变的 Workbook 导入请求。</returns>
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
/// <typeparam name="TItem">Sheet 明细模型类型。</typeparam>
public sealed class ExcelSheetImportBuilder<TItem> where TItem : class, new()
{
    /// <summary>目标工作表名称或选择器显示名称。</summary>
    private readonly string _name;
    /// <summary>用于构造导入实体的目标表达式。</summary>
    private readonly Expression _target;
    /// <summary>选择源工作表的规则。</summary>
    private readonly ExcelSheetSelector _selector;
    /// <summary>表头所在的零基行索引。</summary>
    private int _headerRowIndex;
    /// <summary>数据起始行的零基索引，默认为 1。</summary>
    private int _dataRowStartIndex = 1;
    /// <summary>当前请求定义的动态列集合。</summary>
    private IReadOnlyList<Exports.ExcelDynamicColumnDefinition> _dynamicColumns =
        Array.Empty<Exports.ExcelDynamicColumnDefinition>();
    /// <summary>是否要求源工作表包含期望的表头，默认为 true。</summary>
    private bool _requireExpectedHeaders = true;
    /// <summary>校验失败时的处理模式，默认为遇到首个失败即停止。</summary>
    private ValidateMode _validateMode = ValidateMode.StopOnFirstFailure;
    /// <summary>文本转换使用的区域性，默认为不变区域性。</summary>
    private System.Globalization.CultureInfo _culture = System.Globalization.CultureInfo.InvariantCulture;
    /// <summary>请求级映射配置，未设置时使用默认映射。</summary>
    private Configurations.ExcelMappingConfiguration _requestMappingConfiguration;
    /// <summary>请求级规范化映射文档，未设置时不应用文档配置。</summary>
    private Configurations.ExcelMappingDocument _mappingDocument;
    /// <summary>动态列值读取表达式，未设置时不读取动态值。</summary>
    private Expression<Func<TItem, IDictionary<string, object>>> _dynamicTarget;
    /// <summary>单行允许读取的最大列数，默认为 100。</summary>
    private int _maxReadColumns = 100;
    /// <summary>限制读取范围的列区间，未设置时读取默认范围。</summary>
    private ExcelReadColumnRange _readColumnRange;
    /// <summary>表头名称比较策略，默认忽略大小写。</summary>
    private ExcelNameComparison _headerComparison = ExcelNameComparison.OrdinalIgnoreCase;
    /// <summary>表头文本的空白处理策略，默认为 Trim。</summary>
    private ExcelWhitespacePolicy _headerWhitespace = ExcelWhitespacePolicy.Trim;
    /// <summary>正文文本的空白处理策略，默认为 Trim。</summary>
    private ExcelWhitespacePolicy _bodyWhitespace = ExcelWhitespacePolicy.Trim;
    /// <summary>是否遇到未知动态列时失败。</summary>
    private bool _failOnUnknownDynamicColumns;
    /// <summary>是否将空行写入导入结果。</summary>
    private bool _reportEmptyRows;
    /// <summary>是否在首个空行处停止读取。</summary>
    private bool _stopAtFirstEmptyRow;

    /// <summary>初始化一个 <see cref="ExcelSheetImportBuilder{TItem}" /> 类型的实例。</summary>
    /// <param name="selector">选择源工作表的选择器。</param>
    /// <param name="target">接收导入实体的目标集合表达式。</param>
    internal ExcelSheetImportBuilder(ExcelSheetSelector selector, Expression target)
    {
        _selector = selector;
        _name = selector.Name ?? $"#{selector.Index}";
        _target = target;
    }

    /// <summary>
    /// 设置表头行索引，索引从零开始。
    /// </summary>
    /// <param name="index">表头所在的零基行索引。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
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
    /// <param name="index">正文起始行的零基索引。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> DataRowStartIndex(int index)
    {
        _dataRowStartIndex = index;
        return this;
    }

    /// <summary>
    /// 配置与导出相同的动态列定义。
    /// </summary>
    /// <param name="target">接收动态列值的实体字典属性表达式。</param>
    /// <param name="definitions">当前请求定义的动态列集合。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
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
    /// <param name="value">为 <see langword="true"/> 时要求全部固定列存在，为 <see langword="false"/> 时允许缺少固定列。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> RequireExpectedHeaders(bool value)
    {
        _requireExpectedHeaders = value;
        return this;
    }

    /// <summary>
    /// 设置最大表头列数安全上限。
    /// </summary>
    /// <param name="value">允许读取的最大表头列数。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> MaxReadColumns(int value)
    {
        _maxReadColumns = value;
        return this;
    }

    /// <summary>
    /// 设置实际参与绑定的列读取范围。
    /// </summary>
    /// <param name="startIndex">读取范围起始列的零基索引。</param>
    /// <param name="count">读取的列数。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> ReadColumns(int startIndex, int count)
    {
        _readColumnRange = ExcelReadColumnRange.Create(startIndex, count);
        return this;
    }

    /// <summary>
    /// 设置表头名称比较策略。
    /// </summary>
    /// <param name="comparison">表头名称的比较策略。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> HeaderComparison(ExcelNameComparison comparison)
    {
        _headerComparison = comparison;
        return this;
    }

    /// <summary>
    /// 设置表头文本空白规范化策略。
    /// </summary>
    /// <param name="policy">表头文本的空白规范化策略。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> HeaderWhitespace(ExcelWhitespacePolicy policy)
    {
        _headerWhitespace = policy;
        return this;
    }

    /// <summary>
    /// 设置正文文本空白规范化策略。
    /// </summary>
    /// <param name="policy">正文文本的空白规范化策略。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> BodyWhitespace(ExcelWhitespacePolicy policy)
    {
        _bodyWhitespace = policy;
        return this;
    }

    /// <summary>
    /// 设置未知动态表头是否导致当前 Sheet 失败。
    /// </summary>
    /// <param name="value">为 <see langword="true"/> 时未知动态表头使当前 Sheet 失败，为 <see langword="false"/> 时忽略未知表头。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> FailOnUnknownDynamicColumns(bool value = true)
    {
        _failOnUnknownDynamicColumns = value;
        return this;
    }

    /// <summary>
    /// 设置是否报告空数据行。
    /// </summary>
    /// <param name="value">为 <see langword="true"/> 时将空数据行纳入结果报告。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> ReportEmptyRows(bool value = true)
    {
        _reportEmptyRows = value;
        return this;
    }

    /// <summary>
    /// 设置是否在首个空行后停止读取。
    /// </summary>
    /// <param name="value">为 <see langword="true"/> 时遇到首个空行即停止读取。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> StopAtFirstEmptyRow(bool value = true)
    {
        _stopAtFirstEmptyRow = value;
        return this;
    }

    /// <summary>
    /// 设置校验失败处理粒度。
    /// </summary>
    /// <param name="mode">校验失败时继续读取或立即停止的处理模式。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> Validate(ValidateMode mode)
    {
        _validateMode = mode;
        return this;
    }

    /// <summary>
    /// 设置数字和日期转换区域性。
    /// </summary>
    /// <param name="culture">数字和日期转换使用的区域性。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> Culture(System.Globalization.CultureInfo culture)
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        return this;
    }

    /// <summary>
    /// 设置请求级映射配置。
    /// </summary>
    /// <param name="configuration">请求级导入映射配置；为 null 时清除当前覆盖。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> Mapping(Configurations.ExcelMappingConfiguration configuration)
    {
        _requestMappingConfiguration = configuration == null ? null :
            Configurations.MappingConfigurationCloner.Clone(configuration, Configurations.MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置规范化映射文档的导入方向配置。
    /// </summary>
    /// <param name="document">包含导入方向配置的规范化映射文档。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetImportBuilder<TItem> Mapping(Configurations.ExcelMappingDocument document)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        _mappingDocument = Configurations.MappingDocumentCloner.Clone(document);
        return this;
    }

    /// <summary>验证并生成不可变 Sheet 导入请求。</summary>
    /// <returns>已完成校验的 Sheet 导入请求。</returns>
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
