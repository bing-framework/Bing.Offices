using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 泛型 Sheet 导出构建器。
/// </summary>
public sealed class ExcelSheetExportBuilder<T> where T : class, new()
{
    private readonly string _name;
    private readonly IEnumerable<T> _data;
    private int _headerRowIndex;
    private int _dataRowStartIndex = 1;
    private IReadOnlyList<ExcelDynamicColumnDefinition> _dynamicColumns = Array.Empty<ExcelDynamicColumnDefinition>();
    private bool _failOnUnknownDynamicValues;
    private Func<object, IDictionary<string, object>> _dynamicGetter;
    private Styles.ExcelCellStyle _sheetStyle;
    private Styles.ExcelCellStyle _headerStyle;
    private Styles.ExcelCellStyle _bodyStyle;
    private string _templateRegion;
    private bool _hidden;
    private readonly List<ExcelChartDefinition> _charts = new List<ExcelChartDefinition>();
    private IReadOnlyList<ExcelHeaderRow> _headerRows = Array.Empty<ExcelHeaderRow>();
    private Configurations.ExcelMappingConfiguration _documentMappingConfiguration;
    private Configurations.ExcelMappingConfiguration _requestMappingConfiguration;
    private Configurations.ExcelMappingDocument _mappingDocument;
    private System.Globalization.CultureInfo _culture = System.Globalization.CultureInfo.InvariantCulture;
    private ExcelColumnWidthOptions _columnWidth;
    private ExcelCommentConflictPolicy _commentConflictPolicy = ExcelCommentConflictPolicy.Preserve;

    internal ExcelSheetExportBuilder(string name, IEnumerable<T> data)
    {
        _name = name;
        _data = data;
    }

    /// <summary>
    /// 设置表头行索引，索引从零开始。
    /// </summary>
    public ExcelSheetExportBuilder<T> HeaderRowIndex(int index)
    {
        _headerRowIndex = index;
        if (_dataRowStartIndex == 1)
            _dataRowStartIndex = index + 1;
        return this;
    }

    /// <summary>
    /// 设置正文起始行索引，索引从零开始。
    /// </summary>
    public ExcelSheetExportBuilder<T> DataRowStartIndex(int index)
    {
        _dataRowStartIndex = index;
        return this;
    }

    /// <summary>
    /// 配置请求级动态列，值读取使用稳定 Key。
    /// </summary>
    public ExcelSheetExportBuilder<T> DynamicColumns(
        Expression<Func<T, IDictionary<string, object>>> values,
        IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        if (values == null)
            throw new ArgumentNullException(nameof(values));
        if (definitions == null)
            throw new ArgumentNullException(nameof(definitions));
        _dynamicGetter = values.Compile().ToObjectDictionaryGetter();
        _dynamicColumns = definitions.ToArray();
        return this;
    }

    /// <summary>
    /// 配置未知动态值策略。
    /// </summary>
    public ExcelSheetExportBuilder<T> UnknownDynamicValues(ExcelUnknownDynamicValuePolicy policy)
    {
        _failOnUnknownDynamicValues = policy == ExcelUnknownDynamicValuePolicy.Fail;
        return this;
    }

    /// <summary>
    /// 设置 Sheet 默认样式。
    /// </summary>
    public ExcelSheetExportBuilder<T> SheetStyle(Styles.ExcelCellStyle style)
    {
        _sheetStyle = style;
        return this;
    }

    /// <summary>
    /// 设置表头区域样式。
    /// </summary>
    public ExcelSheetExportBuilder<T> HeaderStyle(Styles.ExcelCellStyle style)
    {
        _headerStyle = style;
        return this;
    }

    /// <summary>
    /// 设置正文区域样式。
    /// </summary>
    public ExcelSheetExportBuilder<T> BodyStyle(Styles.ExcelCellStyle style)
    {
        _bodyStyle = style;
        return this;
    }

    /// <summary>
    /// 设置自定义多级表头。
    /// </summary>
    public ExcelSheetExportBuilder<T> HeaderRows(IReadOnlyList<ExcelHeaderRow> rows)
    {
        _headerRows = rows ?? throw new ArgumentNullException(nameof(rows));
        return this;
    }

    /// <summary>
    /// 设置请求级映射配置。
    /// </summary>
    public ExcelSheetExportBuilder<T> Mapping(Configurations.ExcelMappingConfiguration configuration)
    {
        _requestMappingConfiguration = configuration == null ? null :
            Configurations.MappingConfigurationCloner.Clone(configuration, Configurations.MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置规范化映射文档的导出方向配置。
    /// </summary>
    public ExcelSheetExportBuilder<T> Mapping(Configurations.ExcelMappingDocument document)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        _documentMappingConfiguration = document.Export == null ? null :
            Configurations.MappingConfigurationCloner.Clone(document.Export, Configurations.MappingSourceKind.Document);
        _mappingDocument = Configurations.MappingDocumentCloner.Clone(document);
        return this;
    }

    /// <summary>
    /// 设置值转换使用的区域性。
    /// </summary>
    public ExcelSheetExportBuilder<T> Culture(System.Globalization.CultureInfo culture)
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        return this;
    }

    /// <summary>
    /// 设置当前 Sheet 的列宽策略。
    /// </summary>
    public ExcelSheetExportBuilder<T> ColumnWidth(ExcelColumnWidthOptions options)
    {
        _columnWidth = options ?? throw new ArgumentNullException(nameof(options));
        return this;
    }

    /// <summary>
    /// 设置表头批注与模板已有批注冲突时的处理策略。
    /// </summary>
    public ExcelSheetExportBuilder<T> CommentConflicts(ExcelCommentConflictPolicy policy)
    {
        _commentConflictPolicy = policy;
        return this;
    }

    /// <summary>
    /// 使用模板中的命名区域作为当前 Sheet 写入区域。
    /// </summary>
    public ExcelSheetExportBuilder<T> UseTemplateRegion(string name)
    {
        _templateRegion = name;
        return this;
    }

    /// <summary>
    /// 设置 Sheet 隐藏状态。
    /// </summary>
    public ExcelSheetExportBuilder<T> Hidden(bool hidden = true)
    {
        _hidden = hidden;
        return this;
    }

    /// <summary>
    /// 添加一个基于当前 Sheet 列 Key 的图表。
    /// </summary>
    public ExcelSheetExportBuilder<T> Chart(ExcelChartDefinition chart)
    {
        if (chart == null)
            throw new ArgumentNullException(nameof(chart));
        chart.Validate();
        _charts.Add(chart);
        return this;
    }

    internal ExcelSheetExportRequest Build()
    {
        if (_headerRowIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(_headerRowIndex));
        if (_dataRowStartIndex <= _headerRowIndex)
            throw new ArgumentOutOfRangeException(nameof(_dataRowStartIndex));
        _columnWidth?.Validate();
        if (!Enum.IsDefined(typeof(ExcelCommentConflictPolicy), _commentConflictPolicy))
            throw new ArgumentOutOfRangeException(nameof(_commentConflictPolicy));
        ValidateHeaderRows();
        foreach (var definition in _dynamicColumns)
        {
            if (definition == null)
                throw new ArgumentException("动态列定义不能为空。", nameof(_dynamicColumns));
            if (definition.PhysicalColumnIndex.HasValue && definition.Placement != null)
                throw new ArgumentException($"动态列 {definition.Key} 不能同时指定相对位置和物理索引。",
                    nameof(_dynamicColumns));
            if (definition.Placement?.PhysicalColumnIndex != null && definition.PhysicalColumnIndex.HasValue)
                throw new ArgumentException($"动态列 {definition.Key} 不能重复指定物理索引。",
                    nameof(_dynamicColumns));
        }
        var mappingConfiguration = Configurations.MappingConfigurationMerger.Merge(_documentMappingConfiguration,
            _requestMappingConfiguration, Configurations.MappingSourceKind.Request);
        mappingConfiguration = ExcelDynamicColumnCloner.MergeIntoConfiguration(mappingConfiguration, _dynamicColumns);
        return new ExcelSheetExportRequest(_name, typeof(T), _data, _headerRowIndex, _dataRowStartIndex,
            CloneDynamicColumns(_dynamicColumns), _failOnUnknownDynamicValues, _dynamicGetter, _sheetStyle, _headerStyle, _bodyStyle,
            _templateRegion, _hidden, _charts.AsReadOnly(), _headerRows,
            mappingConfiguration, _mappingDocument,
            _culture, _columnWidth, _commentConflictPolicy);
    }

    private static IReadOnlyList<ExcelDynamicColumnDefinition> CloneDynamicColumns(
        IReadOnlyList<ExcelDynamicColumnDefinition> columns) =>
        (columns ?? Array.Empty<ExcelDynamicColumnDefinition>()).Select(column => new ExcelDynamicColumnDefinition
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
        }).ToArray();

    private void ValidateHeaderRows()
    {
        var occupiedCells = new HashSet<(int Row, int Column)>();
        foreach (var headerRow in _headerRows ?? Array.Empty<ExcelHeaderRow>())
        {
            if (headerRow == null)
                throw new ArgumentException("自定义表头不能包含空行。", nameof(_headerRows));
            foreach (var headerCell in headerRow.Cells)
            {
                if (headerCell == null)
                    throw new ArgumentException("自定义表头不能包含空单元格。", nameof(_headerRows));
                var lastRowIndex = headerRow.RowIndex + headerCell.RowSpan - 1;
                if (lastRowIndex >= _headerRowIndex)
                    throw new ArgumentException("自定义表头不能覆盖属性表头或数据区域。", nameof(_headerRows));
                for (var rowIndex = headerRow.RowIndex; rowIndex <= lastRowIndex; rowIndex++)
                for (var columnIndex = headerCell.ColumnIndex;
                     columnIndex < headerCell.ColumnIndex + headerCell.ColumnSpan; columnIndex++)
                if (!occupiedCells.Add((rowIndex, columnIndex)))
                    throw new ArgumentException("自定义表头包含重叠单元格。", nameof(_headerRows));
            }
        }
    }
}

internal static class ExcelDynamicGetterExtensions
{
    public static Func<object, IDictionary<string, object>> ToObjectDictionaryGetter<T>(
        this Func<T, IDictionary<string, object>> getter) => value => getter((T)value);
}
