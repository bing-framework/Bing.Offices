using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 列表区域构建器。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
public sealed class ExcelEntityListRegionBuilder<TItem> where TItem : class, new()
{
    /// <summary>
    /// 按声明顺序保存动态列组配置。
    /// </summary>
    private readonly List<ExcelEntityDynamicColumnGroup<TItem>> _dynamicColumnGroups =
        new List<ExcelEntityDynamicColumnGroup<TItem>>();
    /// <summary>
    /// 按声明顺序保存计算列配置。
    /// </summary>
    private readonly List<ExcelEntityCalculatedColumn<TItem>> _calculatedColumns =
        new List<ExcelEntityCalculatedColumn<TItem>>();
    /// <summary>
    /// 保存可选的连续分组小计配置。
    /// </summary>
    private IExcelEntityGroupSubtotal _groupSubtotal;

    /// <summary>
    /// 保存可选的每页明细小计配置。
    /// </summary>
    private ExcelEntityListFooter<TItem> _pageSubtotal;

    /// <summary>
    /// 保存可选的明细分页行数。
    /// </summary>
    private int? _pageBreakRows;

    /// <summary>
    /// 获取是否在列表区域起始行包含表头。
    /// </summary>
    internal bool IncludeHeader { get; private set; } = true;
    /// <summary>
    /// 获取列表区域可选右下角的 A1 地址。
    /// </summary>
    internal string EndAddress { get; private set; }
    /// <summary>
    /// 获取列表项的请求级映射配置。
    /// </summary>
    internal ExcelMappingConfiguration MappingConfiguration { get; private set; }
    /// <summary>
    /// 获取列表项的规范化映射文档。
    /// </summary>
    internal ExcelMappingDocument MappingDocument { get; private set; }

    /// <summary>
    /// 获取显式动态列组快照。
    /// </summary>
    internal IReadOnlyList<ExcelEntityDynamicColumnGroup<TItem>> DynamicColumnGroups => _dynamicColumnGroups;

    /// <summary>
    /// 获取计算列快照。
    /// </summary>
    internal IReadOnlyList<ExcelEntityCalculatedColumn<TItem>> CalculatedColumns => _calculatedColumns;

    /// <summary>
    /// 获取连续分组小计定义。
    /// </summary>
    internal IExcelEntityGroupSubtotal GroupSubtotalDefinition => _groupSubtotal;

    /// <summary>
    /// 获取每页明细小计定义。
    /// </summary>
    internal ExcelEntityListFooter<TItem> PageSubtotalDefinition => _pageSubtotal;

    /// <summary>
    /// 获取明细分页行数。
    /// </summary>
    internal int? PageBreakRows => _pageBreakRows;

    /// <summary>
    /// 获取未知动态值处理策略。
    /// </summary>
    internal ExcelUnknownDynamicValuePolicy UnknownDynamicValuePolicy { get; private set; }

    /// <summary>
    /// 获取可变明细尾部定义。
    /// </summary>
    internal ExcelEntityListFooter<TItem> FooterDefinition { get; private set; }

    /// <summary>
    /// 获取最终尾部的工作簿命名锚点。
    /// </summary>
    internal string FooterAnchorName { get; private set; }

    /// <summary>
    /// 设置是否在列表区域起始行写入映射表头。
    /// </summary>
    /// <param name="includeHeader">是否写入表头。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> Header(bool includeHeader = true)
    {
        IncludeHeader = includeHeader;
        return this;
    }

    /// <summary>
    /// 设置列表区域的可选右下角边界。
    /// </summary>
    /// <param name="address">A1 右下角地址。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> End(string address)
    {
        EndAddress = ExcelEntityCellReference.Parse(address).Address;
        return this;
    }

    /// <summary>
    /// 设置列表项的请求级映射配置。
    /// </summary>
    /// <param name="configuration">映射配置。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> Mapping(ExcelMappingConfiguration configuration)
    {
        MappingConfiguration = configuration == null ? null :
            MappingConfigurationCloner.Clone(configuration, MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置列表项的规范化映射文档。
    /// </summary>
    /// <param name="document">映射文档。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> Mapping(ExcelMappingDocument document)
    {
        MappingDocument = document == null ? null : MappingDocumentCloner.Clone(document);
        return this;
    }

    /// <summary>
    /// 添加一个实体属性字典动态列组。
    /// </summary>
    /// <param name="groupKey">动态列组的稳定标识。</param>
    /// <param name="values">实体属性字典表达式。</param>
    /// <param name="definitions">动态列定义。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 字典属性必须是直接属性表达式；动态组与映射配置中的隐式动态列不能同时使用。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> DynamicColumnGroup(string groupKey,
        Expression<Func<TItem, IDictionary<string, object>>> values,
        IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        if (string.IsNullOrWhiteSpace(groupKey))
            throw new ArgumentException("动态列组标识不能为空。", nameof(groupKey));
        if (values == null)
            throw new ArgumentNullException(nameof(values));
        if (definitions == null || definitions.Count == 0)
            throw new ArgumentException("动态列组至少需要一个定义。", nameof(definitions));
        var property = ExcelEntityExpression.GetProperty(values.Body);
        if (!typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
            throw new ArgumentException("动态列字典属性必须实现 IDictionary<string, object>。", nameof(values));
        var getter = values.Compile();
        var snapshots = definitions.Select(CloneDefinition).ToArray();
        _dynamicColumnGroups.Add(new ExcelEntityDynamicColumnGroup<TItem>(groupKey.Trim(), property,
            value => getter((TItem)value), ExcelEntityExpression.CreateObjectSetter<TItem>(property), snapshots));
        return this;
    }

    /// <summary>
    /// 设置未知动态值的处理策略。
    /// </summary>
    /// <param name="policy">未知动态值处理策略。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 未配置显式动态组时，该设置只作为兼容性描述保留，由既有单字典映射处理。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> UnknownDynamicValues(ExcelUnknownDynamicValuePolicy policy)
    {
        if (!Enum.IsDefined(typeof(ExcelUnknownDynamicValuePolicy), policy))
            throw new ArgumentOutOfRangeException(nameof(policy));
        UnknownDynamicValuePolicy = policy;
        return this;
    }

    /// <summary>
    /// 添加一个根据当前行上下文计算值的导出列。
    /// </summary>
    /// <typeparam name="TValue">计算值类型。</typeparam>
    /// <param name="key">计算列的稳定标识。</param>
    /// <param name="title">计算列标题。</param>
    /// <param name="valueFactory">根据行上下文计算单元格值的委托。</param>
    /// <param name="configure">计算列配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 计算列只参与导出；导入时保留其物理列位置但不绑定到实体属性。委托接收本次导出的明细快照，
    /// 每个单元格最多执行一次；取消异常会原样透传，其他异常带有当前 Sheet、行、列和 Key 上下文。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> CalculatedColumn<TValue>(string key, string title,
        Func<ExcelEntityRowContext<TItem>, TValue> valueFactory,
        Action<ExcelEntityCalculatedColumnBuilder> configure = null)
    {
        if (string.IsNullOrWhiteSpace(key))
            throw new ArgumentException("计算列 Key 不能为空。", nameof(key));
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("计算列标题不能为空。", nameof(title));
        if (valueFactory == null)
            throw new ArgumentNullException(nameof(valueFactory));
        var builder = new ExcelEntityCalculatedColumnBuilder();
        configure?.Invoke(builder);
        _calculatedColumns.Add(new ExcelEntityCalculatedColumn<TItem>(key.Trim(), title.Trim(),
            context => valueFactory(context), typeof(TValue), builder));
        return this;
    }

    /// <summary>
    /// 声明可变明细结束后的尾部内容。
    /// </summary>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <param name="configure">尾部配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 尾部相对地址以 marker 所在单元格为 A1；导入时 marker 用于截断明细边界。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> Footer(string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        var builder = new ExcelEntityListFooterBuilder<TItem>();
        configure?.Invoke(builder);
        FooterDefinition = builder.Build(markerText);
        FooterAnchorName = null;
        return this;
    }

    /// <summary>
    /// 声明带工作簿命名锚点的最终尾部。
    /// </summary>
    /// <param name="anchorName">指向最终尾部标记单元格的名称。</param>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <param name="configure">尾部配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 导出时将名称指向实际标记位置；模板中同名、同工作表的单格名称会被更新。
    /// 导入仍由标记文本确定明细边界。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> FooterNamed(string anchorName, string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        ExcelEntityLayoutValidation.ValidateAnchorName(anchorName);
        Footer(markerText, configure);
        FooterAnchorName = anchorName;
        return this;
    }

    /// <summary>
    /// 按相邻明细的分组键写入连续分组小计。
    /// </summary>
    /// <typeparam name="TKey">分组键类型。</typeparam>
    /// <param name="keySelector">读取列表项分组键的表达式。</param>
    /// <param name="markerText">分组小计结束标记文本。</param>
    /// <param name="configure">分组小计尾部配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 分组按明细原有顺序处理，仅当相邻项目的键发生变化时写入小计，不自动排序或合并非连续分组。
    /// 导入时会跳过已识别的分组小计标记行及其间隔和尾部区域。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> GroupSubtotal<TKey>(
        Expression<Func<TItem, TKey>> keySelector, string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        if (keySelector == null)
            throw new ArgumentNullException(nameof(keySelector));
        if (string.IsNullOrWhiteSpace(markerText))
            throw new ArgumentException("分组小计标记不能为空。", nameof(markerText));
        if (_groupSubtotal != null)
            throw new InvalidOperationException("每个列表区域只能配置一个连续分组小计。");
        if (_pageBreakRows.HasValue)
            throw new InvalidOperationException("分页列表不能同时配置连续分组小计。");
        if (_pageSubtotal != null)
            throw new InvalidOperationException("分页小计不能同时配置连续分组小计。");
        var footerBuilder = new ExcelEntityListFooterBuilder<TItem>();
        configure?.Invoke(footerBuilder);
        var footer = footerBuilder.Build(markerText);
        if (footer.ContainsDetailSum)
            throw new ArgumentException("跨小计明细求和公式只能用于最终尾部。", nameof(configure));
        _groupSubtotal = new ExcelEntityGroupSubtotal<TItem, TKey>(keySelector.Compile(), footer);
        return this;
    }

    /// <summary>
    /// 设置列表明细的水平分页行数。
    /// </summary>
    /// <param name="rowsPerPage">每页包含的明细行数，不包含表头。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// Provider 只在导出工作表中写入分页符，不插入空白行，也不改变导入数据边界。
    /// 每个列表区域最多配置一个分页策略；分页不能与连续分组小计组合。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> PageBreak(int rowsPerPage)
    {
        if (rowsPerPage <= 0)
            throw new ArgumentOutOfRangeException(nameof(rowsPerPage), "每页明细行数必须大于零。");
        if (_pageBreakRows.HasValue)
            throw new InvalidOperationException("每个列表区域只能配置一个分页策略。");
        if (_groupSubtotal != null)
            throw new InvalidOperationException("分页列表不能同时配置连续分组小计。");
        _pageBreakRows = rowsPerPage;
        return this;
    }

    /// <summary>
    /// 声明每个完整分页明细后的聚合小计。
    /// </summary>
    /// <param name="markerText">分页小计标记文本。</param>
    /// <param name="configure">分页小计配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 分页小计仅在存在下一页明细时写入，聚合委托接收当前页的明细快照；必须同时配置
    /// <see cref="PageBreak"/>，且不能与连续分组小计组合。导入时会跳过分页小计区域。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> PageSubtotal(string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        if (_pageSubtotal != null)
            throw new InvalidOperationException("每个列表区域只能配置一个分页小计。");
        if (_groupSubtotal != null)
            throw new InvalidOperationException("分页小计不能同时配置连续分组小计。");
        var footerBuilder = new ExcelEntityListFooterBuilder<TItem>();
        configure?.Invoke(footerBuilder);
        var footer = footerBuilder.Build(markerText);
        if (footer.ContainsDetailSum)
            throw new ArgumentException("跨小计明细求和公式只能用于最终尾部。", nameof(configure));
        _pageSubtotal = footer;
        return this;
    }

    /// <summary>
    /// 复制动态列定义及其可变配置。
    /// </summary>
    /// <param name="source">待复制的动态列定义。</param>
    /// <returns>独立的动态列定义。</returns>
    private static ExcelDynamicColumnDefinition CloneDefinition(ExcelDynamicColumnDefinition source)
    {
        if (source == null)
            throw new ArgumentException("动态列定义不能为空。", nameof(source));
        return new ExcelDynamicColumnDefinition
        {
            Key = source.Key,
            Title = source.Title,
            Aliases = source.Aliases == null ? null : source.Aliases.ToArray(),
            DataType = source.DataType,
            Order = source.Order,
            Placement = source.Placement,
            PhysicalColumnIndex = source.PhysicalColumnIndex,
            NumberFormat = source.NumberFormat,
            HeaderStyle = source.HeaderStyle,
            BodyStyle = source.BodyStyle,
            ConverterName = source.ConverterName,
            ValidatorName = source.ValidatorName,
            ValidationRuleNames = source.ValidationRuleNames == null ? null : source.ValidationRuleNames.ToArray(),
            ValidationRules = source.ValidationRules == null ? null : source.ValidationRules
                .Select(CloneValidation).ToArray(),
            ImageMultiplicity = source.ImageMultiplicity
        };
    }

    /// <summary>
    /// 复制动态列校验配置。
    /// </summary>
    /// <param name="source">待复制的校验配置。</param>
    /// <returns>独立的校验配置；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    private static ExcelMappingDynamicValidationConfiguration CloneValidation(
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
}
