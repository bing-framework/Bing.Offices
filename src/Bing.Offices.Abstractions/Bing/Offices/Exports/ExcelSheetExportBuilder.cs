using System.Linq.Expressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 泛型 Sheet 导出构建器。
/// </summary>
/// <typeparam name="T">当前 Sheet 的数据项类型。</typeparam>
public sealed class ExcelSheetExportBuilder<T> where T : class, new()
{
    /// <summary>目标工作表名称。</summary>
    private readonly string _name;
    /// <summary>待写入工作表的数据集合。</summary>
    private readonly IEnumerable<T> _data;
    /// <summary>表头所在的零基行索引，默认为 0。</summary>
    private int _headerRowIndex;
    /// <summary>正文起始行的零基索引，默认为 1。</summary>
    private int _dataRowStartIndex = 1;
    /// <summary>当前请求定义的动态列集合。</summary>
    private IReadOnlyList<ExcelDynamicColumnDefinition> _dynamicColumns = Array.Empty<ExcelDynamicColumnDefinition>();
    /// <summary>是否遇到实体未提供的动态值时失败。</summary>
    private bool _failOnUnknownDynamicValues;
    /// <summary>从实体读取动态列值的委托；未设置时不读取动态值。</summary>
    private Func<object, IDictionary<string, object>> _dynamicGetter;
    /// <summary>工作表级样式配置。</summary>
    private Styles.ExcelCellStyle _sheetStyle;
    /// <summary>表头样式配置。</summary>
    private Styles.ExcelCellStyle _headerStyle;
    /// <summary>正文样式配置。</summary>
    private Styles.ExcelCellStyle _bodyStyle;
    /// <summary>模板中用于定位工作表区域的名称；未设置时使用默认区域。</summary>
    private string _templateRegion;
    /// <summary>导出后是否将工作表标记为隐藏。</summary>
    private bool _hidden;
    /// <summary>按配置顺序保存待创建的图表定义。</summary>
    private readonly List<ExcelChartDefinition> _charts = new List<ExcelChartDefinition>();
    /// <summary>多行表头定义；未设置时使用单行表头。</summary>
    private IReadOnlyList<ExcelHeaderRow> _headerRows = Array.Empty<ExcelHeaderRow>();
    /// <summary>当前工作表的请求级映射配置。</summary>
    private Configurations.ExcelMappingConfiguration _requestMappingConfiguration;
    /// <summary>当前工作表使用的规范化映射文档。</summary>
    private Configurations.ExcelMappingDocument _mappingDocument;
    /// <summary>文本格式化使用的区域性，默认为不变区域性。</summary>
    private System.Globalization.CultureInfo _culture = System.Globalization.CultureInfo.InvariantCulture;
    /// <summary>列宽计算和应用选项。</summary>
    private ExcelColumnWidthOptions _columnWidth;
    /// <summary>单元格批注冲突处理策略，默认为保留模板批注。</summary>
    private ExcelCommentConflictPolicy _commentConflictPolicy = ExcelCommentConflictPolicy.Preserve;
    /// <summary>模板单元格被导出值覆盖时的处理策略，默认为保留模板值。</summary>
    private ExcelTemplateCellOverwritePolicy _templateCellOverwritePolicy =
        ExcelTemplateCellOverwritePolicy.PreserveTemplate;

    /// <summary>初始化一个 <see cref="ExcelSheetExportBuilder{T}" /> 类型的实例。</summary>
    /// <param name="name">目标工作表名称。</param>
    /// <param name="data">待写入工作表的数据集合。</param>
    internal ExcelSheetExportBuilder(string name, IEnumerable<T> data)
    {
        _name = name;
        _data = data;
    }

    /// <summary>
    /// 设置表头行索引，索引从零开始。
    /// </summary>
    /// <param name="index">表头所在的零基行索引。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
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
    /// <param name="index">正文起始行的零基索引。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> DataRowStartIndex(int index)
    {
        _dataRowStartIndex = index;
        return this;
    }

    /// <summary>
    /// 配置请求级动态列。
    /// </summary>
    /// <remarks>动态列值读取使用稳定 Key。</remarks>
    /// <param name="values">从实体读取动态列值的表达式，返回以列 Key 为键的字典。</param>
    /// <param name="definitions">当前请求定义的动态列集合。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
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
    /// <param name="policy">发现未定义动态列值时采用的处理策略。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> UnknownDynamicValues(ExcelUnknownDynamicValuePolicy policy)
    {
        _failOnUnknownDynamicValues = policy == ExcelUnknownDynamicValuePolicy.Fail;
        return this;
    }

    /// <summary>
    /// 设置 Sheet 默认样式。
    /// </summary>
    /// <param name="style">应用于当前 Sheet 的默认样式；为 null 时不覆盖默认样式。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> SheetStyle(Styles.ExcelCellStyle style)
    {
        _sheetStyle = style;
        return this;
    }

    /// <summary>
    /// 设置表头区域样式。
    /// </summary>
    /// <param name="style">应用于表头区域的样式；为 null 时不单独设置。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> HeaderStyle(Styles.ExcelCellStyle style)
    {
        _headerStyle = style;
        return this;
    }

    /// <summary>
    /// 设置正文区域样式。
    /// </summary>
    /// <param name="style">应用于正文区域的样式；为 null 时不单独设置。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> BodyStyle(Styles.ExcelCellStyle style)
    {
        _bodyStyle = style;
        return this;
    }

    /// <summary>
    /// 设置自定义多级表头。
    /// </summary>
    /// <param name="rows">要写入的多级表头行定义。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> HeaderRows(IReadOnlyList<ExcelHeaderRow> rows)
    {
        _headerRows = rows ?? throw new ArgumentNullException(nameof(rows));
        return this;
    }

    /// <summary>
    /// 设置请求级映射配置。
    /// </summary>
    /// <param name="configuration">请求级导出映射配置；为 null 时清除当前覆盖。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> Mapping(Configurations.ExcelMappingConfiguration configuration)
    {
        _requestMappingConfiguration = configuration == null ? null :
            Configurations.MappingConfigurationCloner.Clone(configuration, Configurations.MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置规范化映射文档的导出方向配置。
    /// </summary>
    /// <param name="document">包含导出方向配置的规范化映射文档。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> Mapping(Configurations.ExcelMappingDocument document)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        _mappingDocument = Configurations.MappingDocumentCloner.Clone(document);
        return this;
    }

    /// <summary>
    /// 设置值转换使用的区域性。
    /// </summary>
    /// <param name="culture">值格式化和转换使用的区域性。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> Culture(System.Globalization.CultureInfo culture)
    {
        _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        return this;
    }

    /// <summary>
    /// 设置当前 Sheet 的列宽策略。
    /// </summary>
    /// <param name="options">当前 Sheet 的列宽配置。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> ColumnWidth(ExcelColumnWidthOptions options)
    {
        _columnWidth = options ?? throw new ArgumentNullException(nameof(options));
        return this;
    }

    /// <summary>
    /// 设置表头批注与模板已有批注冲突时的处理策略。
    /// </summary>
    /// <param name="policy">表头批注与模板批注冲突时采用的策略。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> CommentConflicts(ExcelCommentConflictPolicy policy)
    {
        _commentConflictPolicy = policy;
        return this;
    }

    /// <summary>
    /// 设置模板单元格写入策略。
    /// </summary>
    /// <param name="policy">写入模板单元格时采用的覆盖策略。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> TemplateCellOverwrite(ExcelTemplateCellOverwritePolicy policy)
    {
        _templateCellOverwritePolicy = policy;
        return this;
    }

    /// <summary>
    /// 使用模板中的命名区域作为当前 Sheet 写入区域。
    /// </summary>
    /// <param name="name">模板命名区域名称；为空时不指定命名区域。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> UseTemplateRegion(string name)
    {
        _templateRegion = name;
        return this;
    }

    /// <summary>
    /// 设置 Sheet 隐藏状态。
    /// </summary>
    /// <param name="hidden">为 <see langword="true"/> 时隐藏当前 Sheet，为 <see langword="false"/> 时保持可见。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> Hidden(bool hidden = true)
    {
        _hidden = hidden;
        return this;
    }

    /// <summary>
    /// 添加一个基于当前 Sheet 列 Key 的图表。
    /// </summary>
    /// <param name="chart">要添加到当前 Sheet 的图表定义。</param>
    /// <returns>当前 Sheet 构建器，用于继续配置。</returns>
    public ExcelSheetExportBuilder<T> Chart(ExcelChartDefinition chart)
    {
        if (chart == null)
            throw new ArgumentNullException(nameof(chart));
        chart.Validate();
        _charts.Add(chart);
        return this;
    }

    /// <summary>验证并生成不可变 Sheet 导出请求。</summary>
    /// <returns>已完成校验的 Sheet 导出请求。</returns>
    internal ExcelSheetExportRequest Build()
    {
        if (_headerRowIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(_headerRowIndex));
        if (_dataRowStartIndex <= _headerRowIndex)
            throw new ArgumentOutOfRangeException(nameof(_dataRowStartIndex));
        _columnWidth?.Validate();
        if (!Enum.IsDefined(typeof(ExcelCommentConflictPolicy), _commentConflictPolicy))
            throw new ArgumentOutOfRangeException(nameof(_commentConflictPolicy));
        if (!Enum.IsDefined(typeof(ExcelTemplateCellOverwritePolicy), _templateCellOverwritePolicy))
            throw new ArgumentOutOfRangeException(nameof(_templateCellOverwritePolicy));
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
        var requestConfiguration = ExcelDynamicColumnCloner.MergeIntoConfiguration(
            _requestMappingConfiguration, _dynamicColumns);
        return new ExcelSheetExportRequest(_name, typeof(T), _data, _headerRowIndex, _dataRowStartIndex,
            CloneDynamicColumns(_dynamicColumns), _failOnUnknownDynamicValues, _dynamicGetter, _sheetStyle, _headerStyle, _bodyStyle,
            _templateRegion, _hidden, _charts.AsReadOnly(), _headerRows,
            requestConfiguration, _mappingDocument,
            _culture, _columnWidth, _commentConflictPolicy, _templateCellOverwritePolicy);
    }

    /// <summary>复制 Sheet 请求中的动态列定义。</summary>
    /// <param name="columns">待复制的动态列定义集合。</param>
    /// <returns>动态列定义的独立数组。</returns>
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

    /// <summary>验证自定义多行表头的范围和单元格重叠。</summary>
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

/// <summary>将泛型动态值读取器转换为对象字典读取器的扩展类。</summary>
internal static class ExcelDynamicGetterExtensions
{
    /// <summary>将泛型动态值读取器适配为对象读取器。</summary>
    /// <typeparam name="T">动态值所属的实体类型。</typeparam>
    /// <param name="getter">读取实体动态值的泛型委托。</param>
    /// <returns>接受对象并调用泛型读取器的委托。</returns>
    public static Func<object, IDictionary<string, object>> ToObjectDictionaryGetter<T>(
        this Func<T, IDictionary<string, object>> getter) => value => getter((T)value);
}
