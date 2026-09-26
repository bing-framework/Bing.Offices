using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 级 Excel 导出请求。
/// </summary>
public sealed class ExcelWorkbookExportRequest
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookExportRequest" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按输出顺序排列的工作表请求。</param>
    /// <param name="template">可选的模板输入流。</param>
    /// <param name="leaveTemplateOpen">是否由调用方继续持有模板流。</param>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="metadata">工作簿元数据配置。</param>
    /// <param name="metadataSpecified">调用方是否显式设置过元数据。</param>
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

    /// <summary>
    /// 获取不可变 Sheet 执行描述。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetExportRequest> Sheets { get; }

    /// <summary>
    /// 获取模板输入流。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Stream Template { get; }

    /// <summary>
    /// 获取模板流是否由调用方继续持有。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool LeaveTemplateOpen { get; }

    /// <summary>
    /// 获取目标 Excel 格式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelFormat Format { get; }

    /// <summary>
    /// 获取工作簿元数据。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWorkbookMetadataOptions Metadata { get; }

    /// <summary>
    /// 获取调用方是否显式设置过元数据。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool MetadataSpecified { get; }
}

/// <summary>
/// 单个 Excel 导出 Sheet 请求的不可变执行描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelSheetExportRequest
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelSheetExportRequest" /> 类型的实例。
    /// </summary>
    /// <param name="name">工作表名称。</param>
    /// <param name="itemType">工作表数据项类型。</param>
    /// <param name="data">待写入的数据集合。</param>
    /// <param name="headerRowIndex">表头所在的零基行索引。</param>
    /// <param name="dataRowStartIndex">正文起始行的零基索引。</param>
    /// <param name="dynamicColumns">请求级动态列定义。</param>
    /// <param name="failOnUnknownDynamicValues">遇到未知动态值时是否失败。</param>
    /// <param name="dynamicGetter">从数据项读取动态列值的委托。</param>
    /// <param name="sheetStyle">工作表级样式配置。</param>
    /// <param name="headerStyle">表头样式配置。</param>
    /// <param name="bodyStyle">正文样式配置。</param>
    /// <param name="templateRegion">模板中用于定位工作表区域的名称。</param>
    /// <param name="hidden">导出后是否隐藏工作表。</param>
    /// <param name="charts">待创建的图表定义。</param>
    /// <param name="headerRows">多行表头定义。</param>
    /// <param name="mappingConfiguration">请求级映射配置。</param>
    /// <param name="mappingDocument">规范化映射文档。</param>
    /// <param name="culture">文本格式化使用的区域性。</param>
    /// <param name="columnWidth">列宽计算和应用选项。</param>
    /// <param name="rowHeight">行高应用选项。</param>
    /// <param name="commentConflictPolicy">批注冲突处理策略。</param>
    /// <param name="templateCellOverwritePolicy">模板单元格被覆盖时的处理策略。</param>
    /// <param name="tables">公共表格定义。</param>
    /// <param name="autoFilters">公共自动筛选定义。</param>
    /// <param name="freezePane">公共冻结窗格定义。</param>
    /// <param name="conditionalFormats">公共条件格式定义。</param>
    /// <param name="namedRanges">公共名称范围定义。</param>
    /// <param name="printLayout">公共打印布局定义。</param>
    /// <param name="images">嵌入图片定义。</param>
    /// <param name="dataValidations">原生数据校验定义。</param>
    internal ExcelSheetExportRequest(string name, Type itemType, System.Collections.IEnumerable data,
        int headerRowIndex, int dataRowStartIndex, IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns,
        bool failOnUnknownDynamicValues, Func<object, IDictionary<string, object>> dynamicGetter,
        Styles.ExcelCellStyle sheetStyle, Styles.ExcelCellStyle headerStyle, Styles.ExcelCellStyle bodyStyle,
        string templateRegion, bool hidden, IReadOnlyList<ExcelChartDefinition> charts,
        IReadOnlyList<ExcelHeaderRow> headerRows, Configurations.ExcelMappingConfiguration mappingConfiguration,
        Configurations.ExcelMappingDocument mappingDocument,
        System.Globalization.CultureInfo culture, ExcelColumnWidthOptions columnWidth,
        ExcelRowHeightOptions rowHeight,
        ExcelCommentConflictPolicy commentConflictPolicy,
        ExcelTemplateCellOverwritePolicy templateCellOverwritePolicy,
        IReadOnlyList<ExcelTableDefinition> tables,
        IReadOnlyList<ExcelAutoFilterDefinition> autoFilters,
        ExcelFreezePaneDefinition freezePane,
        IReadOnlyList<ExcelConditionalFormatDefinition> conditionalFormats,
        IReadOnlyList<ExcelNamedRangeDefinition> namedRanges,
        ExcelPrintLayoutOptions printLayout,
        IReadOnlyList<ExcelSheetImageDefinition> images,
        IReadOnlyList<ExcelDataValidationDefinition> dataValidations)
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
        RowHeight = rowHeight;
        CommentConflictPolicy = commentConflictPolicy;
        TemplateCellOverwritePolicy = templateCellOverwritePolicy;
        Tables = CloneTables(tables);
        AutoFilters = CloneAutoFilters(autoFilters);
        FreezePane = CloneFreezePane(freezePane);
        ConditionalFormats = CloneConditionalFormats(conditionalFormats);
        NamedRanges = CloneNamedRanges(namedRanges);
        PrintLayout = ClonePrintLayout(printLayout);
        Images = images.Select(image => new ExcelSheetImageDefinition
        {
            Content = image.Content?.ToArray(), Row = image.Row, Column = image.Column,
            Width = image.Width, Height = image.Height, OffsetX = image.OffsetX, OffsetY = image.OffsetY
        }).ToArray();
        DataValidations = dataValidations.Select(rule => new ExcelDataValidationDefinition
        {
            Range = CloneRange(rule.Range), Type = rule.Type, Operator = rule.Operator,
            Value1 = rule.Value1, Value2 = rule.Value2, Date1 = rule.Date1, Date2 = rule.Date2,
            Values = rule.Values?.ToArray(), IgnoreBlanks = rule.IgnoreBlanks,
            ShowErrorMessage = rule.ShowErrorMessage, InputTitle = rule.InputTitle,
            InputMessage = rule.InputMessage, ErrorTitle = rule.ErrorTitle, ErrorMessage = rule.ErrorMessage
        }).ToArray();
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
    /// 获取Sheet 是否隐藏。
    /// </summary>
    public bool Hidden { get; }

    /// <summary>
    /// 获取Sheet 数据项类型。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ItemType { get; }
    /// <summary>
    /// 获取Sheet 数据序列。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Collections.IEnumerable Data { get; }
    /// <summary>
    /// 获取动态列定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelDynamicColumnDefinition> DynamicColumns { get; }
    /// <summary>
    /// 获取是否拒绝未知动态值。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool FailOnUnknownDynamicValues { get; }
    /// <summary>
    /// 获取动态值读取器。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, IDictionary<string, object>> DynamicGetter { get; }
    /// <summary>
    /// 获取Sheet 样式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle SheetStyle { get; }
    /// <summary>
    /// 获取表头样式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle HeaderStyle { get; }
    /// <summary>
    /// 获取正文样式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle BodyStyle { get; }
    /// <summary>
    /// 获取模板命名区域。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string TemplateRegion { get; }
    /// <summary>
    /// 获取图表定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelChartDefinition> Charts { get; }
    /// <summary>
    /// 获取额外表头行。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelHeaderRow> HeaderRows { get; }
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
    /// 获取格式化区域性。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Globalization.CultureInfo Culture { get; }
    /// <summary>
    /// 获取列宽策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelColumnWidthOptions ColumnWidth { get; }
    /// <summary>
    /// 获取行高策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelRowHeightOptions RowHeight { get; }
    /// <summary>
    /// 获取批注冲突策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelCommentConflictPolicy CommentConflictPolicy { get; }
    /// <summary>
    /// 获取模板单元格覆盖策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelTemplateCellOverwritePolicy TemplateCellOverwritePolicy { get; }

    /// <summary>
    /// 获取表格定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelTableDefinition> Tables { get; }
    /// <summary>
    /// 获取自动筛选定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelAutoFilterDefinition> AutoFilters { get; }
    /// <summary>
    /// 获取冻结窗格定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelFreezePaneDefinition FreezePane { get; }
    /// <summary>
    /// 获取条件格式定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelConditionalFormatDefinition> ConditionalFormats { get; }
    /// <summary>
    /// 获取名称范围定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelNamedRangeDefinition> NamedRanges { get; }
    /// <summary>
    /// 获取打印布局定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelPrintLayoutOptions PrintLayout { get; }

    /// <summary>
    /// 获取嵌入图片定义。
    /// </summary>
    public IReadOnlyList<ExcelSheetImageDefinition> Images { get; }
    /// <summary>
    /// 获取 Excel 原生数据校验定义。
    /// </summary>
    public IReadOnlyList<ExcelDataValidationDefinition> DataValidations { get; }

    /// <summary>
    /// 复制表格定义及其区域。
    /// </summary>
    /// <param name="values">待复制的表格定义；允许为空。</param>
    /// <returns>独立的定义数组；输入为空时返回空数组。</returns>
    private static IReadOnlyList<ExcelTableDefinition> CloneTables(IReadOnlyList<ExcelTableDefinition> values) =>
        (values ?? Array.Empty<ExcelTableDefinition>()).Select(value => new ExcelTableDefinition
        {
            Name = value.Name, HasHeaders = value.HasHeaders, ShowTotals = value.ShowTotals,
            StyleName = value.StyleName, Range = CloneRange(value.Range)
        }).ToArray();

    /// <summary>
    /// 复制自动筛选定义及其区域。
    /// </summary>
    /// <param name="values">待复制的筛选定义；允许为空。</param>
    /// <returns>独立的定义数组；输入为空时返回空数组。</returns>
    private static IReadOnlyList<ExcelAutoFilterDefinition> CloneAutoFilters(IReadOnlyList<ExcelAutoFilterDefinition> values) =>
        (values ?? Array.Empty<ExcelAutoFilterDefinition>()).Select(value => new ExcelAutoFilterDefinition
        {
            Range = CloneRange(value.Range)
        }).ToArray();

    /// <summary>
    /// 复制冻结窗格定义。
    /// </summary>
    /// <param name="value">待复制的定义；允许为空。</param>
    /// <returns>独立的定义；输入为空时返回 <see langword="null"/>。</returns>
    private static ExcelFreezePaneDefinition CloneFreezePane(ExcelFreezePaneDefinition value) => value == null ? null : new ExcelFreezePaneDefinition
    {
        Rows = value.Rows, Columns = value.Columns, TopRow = value.TopRow, LeftColumn = value.LeftColumn
    };

    /// <summary>
    /// 复制条件格式定义及其区域。
    /// </summary>
    /// <param name="values">待复制的条件格式定义；允许为空。</param>
    /// <returns>独立的定义数组；输入为空时返回空数组。</returns>
    private static IReadOnlyList<ExcelConditionalFormatDefinition> CloneConditionalFormats(
        IReadOnlyList<ExcelConditionalFormatDefinition> values) =>
        (values ?? Array.Empty<ExcelConditionalFormatDefinition>()).Select(value => new ExcelConditionalFormatDefinition
        {
            Type = value.Type, Range = CloneRange(value.Range), Operator = value.Operator,
            Formula1 = value.Formula1, Formula2 = value.Formula2,
            ForegroundColor = value.ForegroundColor, BackgroundColor = value.BackgroundColor
        }).ToArray();

    /// <summary>
    /// 复制名称范围定义。
    /// </summary>
    /// <param name="values">待复制的名称范围定义；允许为空。</param>
    /// <returns>独立的定义数组；输入为空时返回空数组。</returns>
    private static IReadOnlyList<ExcelNamedRangeDefinition> CloneNamedRanges(IReadOnlyList<ExcelNamedRangeDefinition> values) =>
        (values ?? Array.Empty<ExcelNamedRangeDefinition>()).Select(value => new ExcelNamedRangeDefinition
        {
            Name = value.Name, SheetName = value.SheetName, Address = value.Address
        }).ToArray();

    /// <summary>
    /// 复制打印布局选项。
    /// </summary>
    /// <param name="value">待复制的选项；允许为空。</param>
    /// <returns>独立的选项；输入为空时返回 <see langword="null"/>。</returns>
    private static ExcelPrintLayoutOptions ClonePrintLayout(ExcelPrintLayoutOptions value) => value == null ? null : new ExcelPrintLayoutOptions
    {
        PaperSize = value.PaperSize, Orientation = value.Orientation,
        TopMargin = value.TopMargin, BottomMargin = value.BottomMargin,
        LeftMargin = value.LeftMargin, RightMargin = value.RightMargin,
        ScalePercent = value.ScalePercent, FitToWidth = value.FitToWidth, FitToHeight = value.FitToHeight,
        PrintArea = value.PrintArea, RepeatRows = value.RepeatRows, RepeatColumns = value.RepeatColumns,
        Header = value.Header, Footer = value.Footer
    };

    /// <summary>
    /// 复制工作表区域定义。
    /// </summary>
    /// <param name="value">待复制的区域；允许为空。</param>
    /// <returns>独立的区域定义；输入为空时返回 <see langword="null"/>。</returns>
    private static ExcelRangeDefinition CloneRange(ExcelRangeDefinition value) => value == null ? null : new ExcelRangeDefinition
    {
        StartRow = value.StartRow, StartColumn = value.StartColumn, EndRow = value.EndRow, EndColumn = value.EndColumn
    };
}
