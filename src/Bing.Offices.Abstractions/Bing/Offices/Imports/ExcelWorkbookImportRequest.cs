using System;
using System.Collections.Generic;
using System.Linq.Expressions;

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

    internal IReadOnlyList<ExcelSheetImportRequest> Sheets { get; }
    internal IReadOnlyList<ExcelRelationRequest> Relations { get; }
    internal ExcelNameComparison SheetNameComparison { get; }
    internal ExcelResourceLimits ResourceLimits { get; }
    internal ExcelImportFailureOptions FailureOptions { get; }
    internal ExcelImportValidationMode ValidationMode { get; }
    internal ExcelUnsupportedFeaturePolicy UnsupportedFeaturePolicy { get; }
}

/// <summary>
/// 单个 Sheet 导入请求的执行描述。
/// </summary>
internal sealed class ExcelSheetImportRequest
{
    internal ExcelSheetImportRequest(string name, ExcelSheetSelector selector, Type itemType, Func<object, object> target,
        int headerRowIndex,
        int dataRowStartIndex, IReadOnlyList<Exports.ExcelDynamicColumnDefinition> dynamicColumns,
        Expression dynamicTarget, bool headerMatch, ValidateMode validateMode,
        System.Globalization.CultureInfo culture, Configurations.ExcelMappingConfiguration mappingConfiguration,
        object mappingProfile, Func<object, object> dynamicTargetGetter, int maxColumnLength,
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
        DynamicColumns = dynamicColumns;
        DynamicTarget = dynamicTarget;
        HeaderMatch = headerMatch;
        ValidateMode = validateMode;
        Culture = culture;
        MappingConfiguration = mappingConfiguration;
        MappingProfile = mappingProfile;
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

    internal Type ItemType { get; }
    internal Func<object, object> Target { get; }
    internal IReadOnlyList<Exports.ExcelDynamicColumnDefinition> DynamicColumns { get; }
    internal Expression DynamicTarget { get; }
    internal bool HeaderMatch { get; }
    internal ValidateMode ValidateMode { get; }
    internal System.Globalization.CultureInfo Culture { get; }
    internal Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    internal object MappingProfile { get; }
    internal Func<object, object> DynamicTargetGetter { get; }
    internal int MaxColumnLength { get; }
    internal ExcelReadColumnRange ReadColumnRange { get; }
    internal ExcelNameComparison HeaderComparison { get; }
    internal ExcelWhitespacePolicy HeaderWhitespace { get; }
    internal ExcelWhitespacePolicy BodyWhitespace { get; }
    internal bool FailOnUnknownDynamicColumns { get; }
    internal bool EnabledEmptyLine { get; }
    internal bool IgnoreEmptyLineAfterData { get; }
}

/// <summary>
/// Workbook 导入父子关系执行描述。
/// </summary>
internal sealed class ExcelRelationRequest
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

    internal Func<object, object> Parents { get; }
    internal Func<object, object> Children { get; }
    internal Delegate ParentKey { get; }
    internal Delegate ChildKey { get; }
    internal Func<object, object> Navigation { get; }
    internal Type ParentType { get; }
    internal Type ChildType { get; }
    internal object Comparer { get; }

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
