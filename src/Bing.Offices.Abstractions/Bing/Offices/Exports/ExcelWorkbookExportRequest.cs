using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 级 Excel 导出请求。
/// </summary>
public sealed class ExcelWorkbookExportRequest
{
    internal ExcelWorkbookExportRequest(IReadOnlyList<ExcelSheetExportRequest> sheets, Stream template,
        bool leaveTemplateOpen, ExcelFormat format, ExcelWorkbookMetadataOptions metadata,
        bool metadataSpecified)
    {
        Sheets = sheets;
        Template = template;
        LeaveTemplateOpen = leaveTemplateOpen;
        Format = format;
        Metadata = metadata?.Clone() ?? new ExcelWorkbookMetadataOptions();
        MetadataSpecified = metadataSpecified;
    }

    /// <summary>
    /// 获取请求中的 Sheet 数量。
    /// </summary>
    public int SheetCount => Sheets.Count;

    /// <summary>获取不可变 Sheet 执行描述。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetExportRequest> Sheets { get; }

    /// <summary>获取模板输入流。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Stream Template { get; }

    /// <summary>获取模板流是否由调用方继续持有。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool LeaveTemplateOpen { get; }

    /// <summary>获取目标 Excel 格式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelFormat Format { get; }

    /// <summary>获取工作簿元数据。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWorkbookMetadataOptions Metadata { get; }

    /// <summary>获取调用方是否显式设置过元数据。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool MetadataSpecified { get; }
}

/// <summary>
/// 单个 Excel 导出 Sheet 请求的不可变执行描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelSheetExportRequest
{
    internal ExcelSheetExportRequest(string name, Type itemType, System.Collections.IEnumerable data,
        int headerRowIndex, int dataRowStartIndex, IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns,
        bool failOnUnknownDynamicValues, Func<object, IDictionary<string, object>> dynamicGetter,
        Styles.ExcelCellStyle sheetStyle, Styles.ExcelCellStyle headerStyle, Styles.ExcelCellStyle bodyStyle,
        string templateRegion, bool hidden, IReadOnlyList<ExcelChartDefinition> charts,
        IReadOnlyList<ExcelHeaderRow> headerRows, Configurations.ExcelMappingConfiguration mappingConfiguration,
        Configurations.ExcelMappingDocument mappingDocument,
        System.Globalization.CultureInfo culture, ExcelColumnWidthOptions columnWidth,
        ExcelCommentConflictPolicy commentConflictPolicy,
        ExcelTemplateCellOverwritePolicy templateCellOverwritePolicy)
    {
        Name = name;
        ItemType = itemType;
        Data = data;
        HeaderRowIndex = headerRowIndex;
        DataRowStartIndex = dataRowStartIndex;
        DynamicColumns = dynamicColumns?.Select(column => new ExcelDynamicColumnDefinition
        {
            Key = column.Key,
            Title = column.Title,
            Aliases = (column.Aliases ?? Array.Empty<string>()).ToArray(),
            DataType = column.DataType,
            Order = column.Order,
            Placement = column.Placement,
            PhysicalColumnIndex = column.PhysicalColumnIndex,
            NumberFormat = column.NumberFormat,
            HeaderStyle = column.HeaderStyle,
            BodyStyle = column.BodyStyle,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = (column.ValidationRuleNames ?? Array.Empty<string>()).ToArray(),
            ImageMultiplicity = column.ImageMultiplicity
        }).ToArray() ?? Array.Empty<ExcelDynamicColumnDefinition>();
        FailOnUnknownDynamicValues = failOnUnknownDynamicValues;
        DynamicGetter = dynamicGetter;
        SheetStyle = sheetStyle;
        HeaderStyle = headerStyle;
        BodyStyle = bodyStyle;
        TemplateRegion = templateRegion;
        Hidden = hidden;
        Charts = charts;
        HeaderRows = headerRows;
        MappingConfiguration = mappingConfiguration == null ? null :
            MappingConfigurationCloner.Clone(mappingConfiguration, mappingConfiguration.SourceKind);
        MappingDocument = Configurations.MappingDocumentCloner.Clone(mappingDocument);
        Culture = culture;
        ColumnWidth = columnWidth;
        CommentConflictPolicy = commentConflictPolicy;
        TemplateCellOverwritePolicy = templateCellOverwritePolicy;
    }

    /// <summary>
    /// 获取请求中的 Sheet 名称。
    /// </summary>
    public string Name { get; }

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
    /// 获取 Sheet 是否隐藏。
    /// </summary>
    public bool Hidden { get; }

    /// <summary>获取 Sheet 数据项类型。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ItemType { get; }
    /// <summary>获取 Sheet 数据序列。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Collections.IEnumerable Data { get; }
    /// <summary>获取动态列定义。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelDynamicColumnDefinition> DynamicColumns { get; }
    /// <summary>获取是否拒绝未知动态值。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool FailOnUnknownDynamicValues { get; }
    /// <summary>获取动态值读取器。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, IDictionary<string, object>> DynamicGetter { get; }
    /// <summary>获取 Sheet 样式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle SheetStyle { get; }
    /// <summary>获取表头样式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle HeaderStyle { get; }
    /// <summary>获取正文样式。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle BodyStyle { get; }
    /// <summary>获取模板命名区域。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string TemplateRegion { get; }
    /// <summary>获取图表定义。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelChartDefinition> Charts { get; }
    /// <summary>获取额外表头行。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelHeaderRow> HeaderRows { get; }
    /// <summary>获取映射配置。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    /// <summary>获取映射文档。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingDocument MappingDocument { get; }
    /// <summary>获取格式化区域性。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Globalization.CultureInfo Culture { get; }
    /// <summary>获取列宽策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelColumnWidthOptions ColumnWidth { get; }
    /// <summary>获取批注冲突策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelCommentConflictPolicy CommentConflictPolicy { get; }
    /// <summary>获取模板单元格覆盖策略。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelTemplateCellOverwritePolicy TemplateCellOverwritePolicy { get; }
}
