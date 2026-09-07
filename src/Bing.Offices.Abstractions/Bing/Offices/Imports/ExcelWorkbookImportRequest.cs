using System.Linq.Expressions;
using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 导入请求。
/// </summary>
public sealed class ExcelWorkbookImportRequest<TWorkbook> where TWorkbook : class, new()
{
    internal ExcelWorkbookImportRequest(IReadOnlyList<ExcelSheetImportRequest> sheets,
        IReadOnlyList<ExcelRelationRequest> relations, ExcelNameComparison sheetNameComparison,
        ExcelResourceLimits resourceLimits, ExcelImportFailureOptions failureOptions,
        ExcelImportValidationMode validationMode, ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy)
    {
        Sheets = sheets;
        Relations = relations;
        SheetNameComparison = sheetNameComparison;
        ResourceLimits = resourceLimits;
        FailureOptions = failureOptions;
        ValidationMode = validationMode;
        UnsupportedFeaturePolicy = unsupportedFeaturePolicy;
    }

    /// <summary>
    /// 获取 Sheet 配置数量。
    /// </summary>
    public int SheetCount => Sheets.Count;

    /// <summary>获取不可变 Sheet 导入描述。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetImportRequest> Sheets { get; }
    /// <summary>获取父子关系描述。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelRelationRequest> Relations { get; }
    /// <summary>获取 Sheet 名称比较策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelNameComparison SheetNameComparison { get; }
    /// <summary>获取输入资源限制。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelResourceLimits ResourceLimits { get; }
    /// <summary>获取失败工作簿选项。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportFailureOptions FailureOptions { get; }
    /// <summary>获取校验模式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportValidationMode ValidationMode { get; }
    /// <summary>获取不支持特性策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelUnsupportedFeaturePolicy UnsupportedFeaturePolicy { get; }
}

/// <summary>
/// 单个 Sheet 导入请求的执行描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelSheetImportRequest
{
    internal ExcelSheetImportRequest(string name, ExcelSheetSelector selector, Type itemType, Func<object, object> target,
        int headerRowIndex,
        int dataRowStartIndex, IReadOnlyList<Exports.ExcelDynamicColumnDefinition> dynamicColumns,
        Expression dynamicTarget, bool requireExpectedHeaders, ValidateMode validateMode,
        System.Globalization.CultureInfo culture, Configurations.ExcelMappingConfiguration mappingConfiguration,
        Configurations.ExcelMappingDocument mappingDocument,
        Func<object, object> dynamicTargetGetter, int maxReadColumns,
        bool failOnUnknownDynamicColumns,
        bool reportEmptyRows, bool stopAtFirstEmptyRow, ExcelReadColumnRange readColumnRange,
        ExcelNameComparison headerComparison, ExcelWhitespacePolicy headerWhitespace,
        ExcelWhitespacePolicy bodyWhitespace)
    {
        Name = name;
        Selector = selector;
        ItemType = itemType;
        Target = target;
        HeaderRowIndex = headerRowIndex;
        DataRowStartIndex = dataRowStartIndex;
        DynamicColumns = dynamicColumns?.ToArray() ?? Array.Empty<Exports.ExcelDynamicColumnDefinition>();
        DynamicTarget = dynamicTarget;
        RequireExpectedHeaders = requireExpectedHeaders;
        ValidateMode = validateMode;
        Culture = culture;
        MappingConfiguration = mappingConfiguration == null ? null :
            MappingConfigurationCloner.Clone(mappingConfiguration, mappingConfiguration.SourceKind);
        MappingDocument = Configurations.MappingDocumentCloner.Clone(mappingDocument);
        DynamicTargetGetter = dynamicTargetGetter;
        MaxReadColumns = maxReadColumns;
        FailOnUnknownDynamicColumns = failOnUnknownDynamicColumns;
        ReportEmptyRows = reportEmptyRows;
        StopAtFirstEmptyRow = stopAtFirstEmptyRow;
        ReadColumnRange = readColumnRange;
        HeaderComparison = headerComparison;
        HeaderWhitespace = headerWhitespace;
        BodyWhitespace = bodyWhitespace;
    }

    /// <summary>
    /// 获取 Sheet 名称。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 获取工作表选择器。
    /// </summary>
    public ExcelSheetSelector Selector { get; }

    /// <summary>
    /// 获取表头行索引，索引从零开始。
    /// </summary>
    public int HeaderRowIndex { get; }

    /// <summary>
    /// 获取正文起始行索引，索引从零开始。
    /// </summary>
    public int DataRowStartIndex { get; }

    /// <summary>
    /// 获取动态列数量。
    /// </summary>
    public int DynamicColumnCount => DynamicColumns.Count;

    /// <summary>获取 Sheet 数据项类型。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ItemType { get; }
    /// <summary>获取目标集合读取器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Target { get; }
    /// <summary>获取动态列定义。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<Exports.ExcelDynamicColumnDefinition> DynamicColumns { get; }
    /// <summary>获取动态目标表达式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Expression DynamicTarget { get; }
    /// <summary>获取是否要求预期表头。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool RequireExpectedHeaders { get; }
    /// <summary>获取单元格校验模式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ValidateMode ValidateMode { get; }
    /// <summary>获取解析区域性。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Globalization.CultureInfo Culture { get; }
    /// <summary>获取映射配置。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    /// <summary>获取映射文档。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingDocument MappingDocument { get; }
    /// <summary>获取动态目标读取器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> DynamicTargetGetter { get; }
    /// <summary>获取最大读取列数。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public int MaxReadColumns { get; }
    /// <summary>获取读取列范围。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelReadColumnRange ReadColumnRange { get; }
    /// <summary>获取表头比较策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelNameComparison HeaderComparison { get; }
    /// <summary>获取表头空白处理策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWhitespacePolicy HeaderWhitespace { get; }
    /// <summary>获取正文空白处理策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWhitespacePolicy BodyWhitespace { get; }
    /// <summary>获取是否拒绝未知动态列。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool FailOnUnknownDynamicColumns { get; }
    /// <summary>获取是否报告空行。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool ReportEmptyRows { get; }
    /// <summary>获取是否遇到首个空行即停止。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool StopAtFirstEmptyRow { get; }
}

/// <summary>
/// Workbook 导入父子关系执行描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelRelationRequest
{
    private ExcelRelationRequest(Func<object, object> parents, Func<object, object> children, Delegate parentKey,
        Delegate childKey, Func<object, object> navigation, Type parentType, Type childType,
        object comparer)
    {
        Parents = parents;
        Children = children;
        ParentKey = parentKey;
        ChildKey = childKey;
        Navigation = navigation;
        ParentType = parentType;
        ChildType = childType;
        Comparer = comparer;
    }

    /// <summary>获取父集合读取器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Parents { get; }
    /// <summary>获取子集合读取器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Children { get; }
    /// <summary>获取父键委托。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Delegate ParentKey { get; }
    /// <summary>获取子键委托。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Delegate ChildKey { get; }
    /// <summary>获取子导航属性写入器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Navigation { get; }
    /// <summary>获取父实体类型。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ParentType { get; }
    /// <summary>获取子实体类型。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ChildType { get; }
    /// <summary>获取键比较器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public object Comparer { get; }

    internal static ExcelRelationRequest Create<TWorkbook, TParent, TChild, TKey>(
        Expression<Func<TWorkbook, ICollection<TParent>>> parents,
        Expression<Func<TWorkbook, ICollection<TChild>>> children,
        Func<TParent, TKey> parentKey,
        Func<TChild, TKey> childKey,
        Expression<Func<TParent, ICollection<TChild>>> navigation,
        IEqualityComparer<TKey> comparer)
        where TParent : class where TChild : class
    {
        var parentGetter = parents.Compile();
        var childGetter = children.Compile();
        var navigationGetter = navigation.Compile();
        return new ExcelRelationRequest(value => parentGetter((TWorkbook)value), value => childGetter((TWorkbook)value),
            parentKey, childKey, value => navigationGetter((TParent)value), typeof(TParent), typeof(TChild),
            comparer ?? EqualityComparer<TKey>.Default);
    }
}
