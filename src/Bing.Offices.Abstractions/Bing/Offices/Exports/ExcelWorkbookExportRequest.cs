using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Linq;
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

    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetExportRequest> Sheets { get; }

    [EditorBrowsable(EditorBrowsableState.Never)]
    public Stream Template { get; }

    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool LeaveTemplateOpen { get; }

    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelFormat Format { get; }

    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWorkbookMetadataOptions Metadata { get; }

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

    [EditorBrowsable(EditorBrowsableState.Never)]
    public Type ItemType { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Collections.IEnumerable Data { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelDynamicColumnDefinition> DynamicColumns { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public bool FailOnUnknownDynamicValues { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, IDictionary<string, object>> DynamicGetter { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle SheetStyle { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle HeaderStyle { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Styles.ExcelCellStyle BodyStyle { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string TemplateRegion { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelChartDefinition> Charts { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelHeaderRow> HeaderRows { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Configurations.ExcelMappingDocument MappingDocument { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Globalization.CultureInfo Culture { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelColumnWidthOptions ColumnWidth { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelCommentConflictPolicy CommentConflictPolicy { get; }
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelTemplateCellOverwritePolicy TemplateCellOverwritePolicy { get; }
}
