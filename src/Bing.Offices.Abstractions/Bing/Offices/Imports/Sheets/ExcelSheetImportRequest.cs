using System.Linq.Expressions;
using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// 单个 Sheet 导入请求的执行描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelSheetImportRequest
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelSheetImportRequest" /> 类型的实例。
    /// </summary>
    /// <param name="name">工作表显示名称。</param>
    /// <param name="selector">源工作表选择器。</param>
    /// <param name="itemType">工作表数据项类型。</param>
    /// <param name="target">接收导入实体的目标集合读取器。</param>
    /// <param name="headerRowIndex">表头所在的零基行索引。</param>
    /// <param name="dataRowStartIndex">正文起始行的零基索引。</param>
    /// <param name="dynamicColumns">动态列定义集合。</param>
    /// <param name="dynamicTarget">接收动态列值的目标表达式。</param>
    /// <param name="requireExpectedHeaders">是否要求源表包含期望表头。</param>
    /// <param name="validationFailureMode">单元格校验失败处理模式。</param>
    /// <param name="culture">文本和数值解析使用的区域性。</param>
    /// <param name="mappingConfiguration">请求级映射配置。</param>
    /// <param name="mappingDocument">规范化映射文档。</param>
    /// <param name="dynamicTargetGetter">读取动态目标字典的委托。</param>
    /// <param name="maxReadColumns">单行允许读取的最大列数。</param>
    /// <param name="failOnUnknownDynamicColumns">是否遇到未知动态列时失败。</param>
    /// <param name="reportEmptyRows">是否报告空行。</param>
    /// <param name="stopAtFirstEmptyRow">是否在首个空行处停止读取。</param>
    /// <param name="readColumnRange">可选的读取列范围。</param>
    /// <param name="headerComparison">表头名称比较策略。</param>
    /// <param name="headerWhitespace">表头文本空白处理策略。</param>
    /// <param name="bodyWhitespace">正文文本空白处理策略。</param>
    internal ExcelSheetImportRequest(string name, ExcelSheetSelector selector, Type itemType, Func<object, object> target,
        int headerRowIndex,
        int dataRowStartIndex, IReadOnlyList<Exports.ExcelDynamicColumnDefinition> dynamicColumns,
        Expression dynamicTarget, bool requireExpectedHeaders, ExcelValidationFailureMode validationFailureMode,
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
        ValidationFailureMode = validationFailureMode;
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
    /// 获取Sheet 名称。
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

    /// <summary>
    /// 获取Sheet 数据项类型。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ItemType { get; }
    /// <summary>
    /// 获取目标集合读取器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Target { get; }
    /// <summary>
    /// 获取动态列定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<Exports.ExcelDynamicColumnDefinition> DynamicColumns { get; }
    /// <summary>
    /// 获取动态目标表达式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Expression DynamicTarget { get; }
    /// <summary>
    /// 获取是否要求预期表头。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool RequireExpectedHeaders { get; }
    /// <summary>
    /// 获取单元格校验失败处理模式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelValidationFailureMode ValidationFailureMode { get; }
    /// <summary>
    /// 获取解析区域性。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Globalization.CultureInfo Culture { get; }
    /// <summary>
    /// 获取映射配置。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    /// <summary>
    /// 获取映射文档。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingDocument MappingDocument { get; }
    /// <summary>
    /// 获取动态目标读取器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> DynamicTargetGetter { get; }
    /// <summary>
    /// 获取最大读取列数。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public int MaxReadColumns { get; }
    /// <summary>
    /// 获取读取列范围。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelReadColumnRange ReadColumnRange { get; }
    /// <summary>
    /// 获取表头比较策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelNameComparison HeaderComparison { get; }
    /// <summary>
    /// 获取表头空白处理策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWhitespacePolicy HeaderWhitespace { get; }
    /// <summary>
    /// 获取正文空白处理策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWhitespacePolicy BodyWhitespace { get; }
    /// <summary>
    /// 获取是否拒绝未知动态列。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool FailOnUnknownDynamicColumns { get; }
    /// <summary>
    /// 获取是否报告空行。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool ReportEmptyRows { get; }
    /// <summary>
    /// 获取是否遇到首个空行即停止。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool StopAtFirstEmptyRow { get; }
}
