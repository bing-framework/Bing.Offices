using System;
using System.Collections.Generic;
using System.IO;

namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 级 Excel 导出请求。
/// </summary>
public sealed class ExcelWorkbookExportRequest
{
    internal ExcelWorkbookExportRequest(IReadOnlyList<ExcelSheetExportRequest> sheets, Stream template,
        bool leaveTemplateOpen, ExcelFormat format)
    {
        Sheets = sheets;
        Template = template;
        LeaveTemplateOpen = leaveTemplateOpen;
        Format = format;
    }

    /// <summary>
    /// 获取请求中的 Sheet 数量。
    /// </summary>
    public int SheetCount => Sheets.Count;

    internal IReadOnlyList<ExcelSheetExportRequest> Sheets { get; }

    internal Stream Template { get; }

    internal bool LeaveTemplateOpen { get; }

    internal ExcelFormat Format { get; }
}

/// <summary>
/// 单个 Excel 导出 Sheet 请求的不可变执行描述。
/// </summary>
internal sealed class ExcelSheetExportRequest
{
    internal ExcelSheetExportRequest(string name, Type itemType, System.Collections.IEnumerable data,
        int headerRowIndex, int dataRowStartIndex, IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns,
        bool failOnUnknownDynamicValues, Func<object, IDictionary<string, object>> dynamicGetter,
        Styles.ExcelCellStyle sheetStyle, Styles.ExcelCellStyle headerStyle, Styles.ExcelCellStyle bodyStyle,
        string templateRegion, bool hidden, IReadOnlyList<ExcelChartDefinition> charts,
        IReadOnlyList<ExcelHeaderRow> headerRows, Configurations.ExcelMappingConfiguration mappingConfiguration,
        object mappingProfile, System.Globalization.CultureInfo culture, ExcelColumnWidthOptions columnWidth,
        ExcelCommentConflictPolicy commentConflictPolicy)
    {
        Name = name;
        ItemType = itemType;
        Data = data;
        HeaderRowIndex = headerRowIndex;
        DataRowStartIndex = dataRowStartIndex;
        DynamicColumns = dynamicColumns;
        FailOnUnknownDynamicValues = failOnUnknownDynamicValues;
        DynamicGetter = dynamicGetter;
        SheetStyle = sheetStyle;
        HeaderStyle = headerStyle;
        BodyStyle = bodyStyle;
        TemplateRegion = templateRegion;
        Hidden = hidden;
        Charts = charts;
        HeaderRows = headerRows;
        MappingConfiguration = mappingConfiguration;
        MappingProfile = mappingProfile;
        Culture = culture;
        ColumnWidth = columnWidth;
        CommentConflictPolicy = commentConflictPolicy;
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

    internal Type ItemType { get; }
    internal System.Collections.IEnumerable Data { get; }
    internal IReadOnlyList<ExcelDynamicColumnDefinition> DynamicColumns { get; }
    internal bool FailOnUnknownDynamicValues { get; }
    internal Func<object, IDictionary<string, object>> DynamicGetter { get; }
    internal Styles.ExcelCellStyle SheetStyle { get; }
    internal Styles.ExcelCellStyle HeaderStyle { get; }
    internal Styles.ExcelCellStyle BodyStyle { get; }
    internal string TemplateRegion { get; }
    internal IReadOnlyList<ExcelChartDefinition> Charts { get; }
    internal IReadOnlyList<ExcelHeaderRow> HeaderRows { get; }
    internal Configurations.ExcelMappingConfiguration MappingConfiguration { get; }
    internal object MappingProfile { get; }
    internal System.Globalization.CultureInfo Culture { get; }
    internal ExcelColumnWidthOptions ColumnWidth { get; }
    internal ExcelCommentConflictPolicy CommentConflictPolicy { get; }
}
