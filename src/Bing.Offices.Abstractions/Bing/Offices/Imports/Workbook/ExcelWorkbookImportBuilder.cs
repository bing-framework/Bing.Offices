using System.Linq.Expressions;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 导入构建器。
/// </summary>
/// <typeparam name="TWorkbook">根 Workbook 模型类型。</typeparam>
public sealed class ExcelWorkbookImportBuilder<TWorkbook> where TWorkbook : class, new()
{
    /// <summary>
    /// 按配置顺序保存待导入的工作表请求。
    /// </summary>
    private readonly List<ExcelSheetImportRequest> _sheets = new List<ExcelSheetImportRequest>();
    /// <summary>
    /// 保存 Workbook 级关系绑定请求，构建完成时转换为只读快照。
    /// </summary>
    private readonly List<ExcelRelationRequest> _relations = new List<ExcelRelationRequest>();
    /// <summary>
    /// 按名称选择工作表时使用的比较策略，默认忽略大小写。
    /// </summary>
    private ExcelNameComparison _sheetNameComparison = ExcelNameComparison.OrdinalIgnoreCase;
    /// <summary>
    /// 导入过程的资源限制；未设置时沿用 Provider 默认值。
    /// </summary>
    private ExcelResourceLimits _resourceLimits;
    /// <summary>
    /// 失败工作簿输出选项；未设置时不生成失败输出。
    /// </summary>
    private ExcelImportFailureOptions _failureOptions;
    /// <summary>
    /// 导入校验模式，默认执行已配置的校验规则。
    /// </summary>
    private ExcelImportValidationMode _validationMode = ExcelImportValidationMode.ConfiguredRules;
    /// <summary>
    /// 不支持功能的处理策略，默认遇到不支持功能时失败。
    /// </summary>
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

    /// <summary>
    /// 构建并校验当前 Workbook 导入请求。
    /// </summary>
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
