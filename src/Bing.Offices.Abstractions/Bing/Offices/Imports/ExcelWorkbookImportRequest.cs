using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.ComponentModel;
using System.Linq;
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

    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetImportRequest> Sheets { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelRelationRequest> Relations { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelNameComparison SheetNameComparison { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelResourceLimits ResourceLimits { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportFailureOptions FailureOptions { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportValidationMode ValidationMode { get; }
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
        Expression dynamicTarget, bool headerMatch, ValidateMode validateMode,
        System.Globalization.CultureInfo culture, Configurations.ExcelMappingConfiguration mappingConfiguration,
        Configurations.ExcelMappingDocument mappingDocument,
        Func<object, object> dynamicTargetGetter, int maxColumnLength,
        bool failOnUnknownDynamicColumns,
        bool enabledEmptyLine, bool ignoreEmptyLineAfterData, ExcelReadColumnRange readColumnRange,
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
        HeaderMatch = headerMatch;
        ValidateMode = validateMode;
        Culture = culture;
        MappingConfiguration = mappingConfiguration == null ? null :
            MappingConfigurationCloner.Clone(mappingConfiguration, mappingConfiguration.SourceKind);
        MappingDocument = Configurations.MappingDocumentCloner.Clone(mappingDocument);
        DynamicTargetGetter = dynamicTargetGetter;
        MaxColumnLength = maxColumnLength;
        FailOnUnknownDynamicColumns = failOnUnknownDynamicColumns;
        EnabledEmptyLine = enabledEmptyLine;
        IgnoreEmptyLineAfterData = ignoreEmptyLineAfterData;
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

    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ItemType { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Target { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<Exports.ExcelDynamicColumnDefinition> DynamicColumns { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Expression DynamicTarget { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool HeaderMatch { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ValidateMode ValidateMode { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Globalization.CultureInfo Culture { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingDocument MappingDocument { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> DynamicTargetGetter { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public int MaxColumnLength { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelReadColumnRange ReadColumnRange { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelNameComparison HeaderComparison { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWhitespacePolicy HeaderWhitespace { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWhitespacePolicy BodyWhitespace { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool FailOnUnknownDynamicColumns { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool EnabledEmptyLine { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool IgnoreEmptyLineAfterData { get; }
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

    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Parents { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Children { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Delegate ParentKey { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Delegate ChildKey { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Navigation { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ParentType { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ChildType { get; }
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
