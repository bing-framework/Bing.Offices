using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 列表区域绑定描述。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityListRegion<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 保存与列表项对应的请求级映射配置快照。
    /// </summary>
    private readonly ExcelMappingConfiguration _mappingConfiguration;
    /// <summary>
    /// 保存与列表项对应的规范化映射文档快照。
    /// </summary>
    private readonly ExcelMappingDocument _mappingDocument;

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityListRegion{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="start">列表区域起始坐标。</param>
    /// <param name="propertyName">集合属性名称。</param>
    /// <param name="itemType">列表项类型。</param>
    /// <param name="getter">从聚合对象读取列表项的委托。</param>
    /// <param name="setter">向聚合对象替换列表项的委托；只读属性时为 <see langword="null" />。</param>
    /// <param name="includeHeader">列表区域起始行是否包含表头。</param>
    /// <param name="end">可选的列表区域右下角坐标。</param>
    /// <param name="mappingConfiguration">列表项的请求级映射配置。</param>
    /// <param name="mappingDocument">列表项的规范化映射文档。</param>
    /// <param name="dynamicColumnGroups">显式动态列组。</param>
    /// <param name="unknownDynamicValues">未知动态值处理策略。</param>
    /// <param name="footer">可选的明细尾部定义。</param>
    /// <param name="calculatedColumns">导出计算列定义。</param>
    /// <param name="groupSubtotal">可选的连续分组小计定义。</param>
    /// <param name="pageBreakRows">可选的每页明细行数。</param>
    /// <param name="pageSubtotal">可选的每页明细小计定义。</param>
    /// <param name="anchorName">可选的模板命名锚点名称。</param>
    /// <param name="footerAnchorName">可选的最终尾部命名锚点名称。</param>
    internal ExcelEntityListRegion(string sheetName, ExcelEntityCellReference start, string propertyName,
        Type itemType, Func<object, IEnumerable> getter, Action<object, object> setter, bool includeHeader,
        ExcelEntityCellReference? end, ExcelMappingConfiguration mappingConfiguration,
        ExcelMappingDocument mappingDocument,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> dynamicColumnGroups = null,
        ExcelUnknownDynamicValuePolicy unknownDynamicValues = ExcelUnknownDynamicValuePolicy.Ignore,
        IExcelEntityListFooter footer = null,
        IReadOnlyList<IExcelEntityCalculatedColumn> calculatedColumns = null,
        IExcelEntityGroupSubtotal groupSubtotal = null,
        int? pageBreakRows = null,
        IExcelEntityListFooter pageSubtotal = null,
        string anchorName = null,
        string footerAnchorName = null)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Start = start;
        PropertyName = propertyName ?? throw new ArgumentNullException(nameof(propertyName));
        ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
        Getter = getter ?? throw new ArgumentNullException(nameof(getter));
        Setter = setter;
        IncludeHeader = includeHeader;
        End = end;
        AnchorName = anchorName;
        FooterAnchorName = footerAnchorName;
        DynamicColumnGroups = Array.AsReadOnly((dynamicColumnGroups ??
            Array.Empty<IExcelEntityDynamicColumnGroup>()).ToArray());
        UnknownDynamicValues = unknownDynamicValues;
        Footer = footer;
        CalculatedColumns = Array.AsReadOnly((calculatedColumns ??
            Array.Empty<IExcelEntityCalculatedColumn>()).ToArray());
        GroupSubtotal = groupSubtotal;
        PageBreakRows = pageBreakRows;
        PageSubtotal = pageSubtotal;
        _mappingConfiguration = mappingConfiguration == null ? null :
            MappingConfigurationCloner.Clone(mappingConfiguration, mappingConfiguration.SourceKind);
        _mappingDocument = mappingDocument == null ? null : MappingDocumentCloner.Clone(mappingDocument);
    }

    /// <summary>
    /// 获取目标工作表名称。
    /// </summary>
    public string SheetName { get; }
    /// <summary>
    /// 获取列表区域的零基起始坐标。
    /// </summary>
    public ExcelEntityCellReference Start { get; }
    /// <summary>
    /// 获取可选的模板命名锚点名称。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string AnchorName { get; }

    /// <summary>
    /// 获取最终尾部标记单元格的工作簿命名锚点。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string FooterAnchorName { get; }
    /// <summary>
    /// 获取列表区域可选的零基右下角坐标。
    /// </summary>
    public ExcelEntityCellReference? End { get; }
    /// <summary>
    /// 获取绑定的集合属性名称。
    /// </summary>
    public string PropertyName { get; }
    /// <summary>
    /// 获取列表区域中的项类型。
    /// </summary>
    public Type ItemType { get; }
    /// <summary>
    /// 获取列表区域起始行是否包含映射表头。
    /// </summary>
    public bool IncludeHeader { get; }
    /// <summary>
    /// 获取从聚合对象读取列表项的已编译委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, IEnumerable> Getter { get; }
    /// <summary>
    /// 获取向聚合对象替换列表项的委托；只读集合属性时为 <see langword="null" />。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Action<object, object> Setter { get; }
    /// <summary>
    /// 获取显式动态列组。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<IExcelEntityDynamicColumnGroup> DynamicColumnGroups { get; }
    /// <summary>
    /// 获取未知动态值处理策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelUnknownDynamicValuePolicy UnknownDynamicValues { get; }
    /// <summary>
    /// 获取可变明细尾部定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IExcelEntityListFooter Footer { get; }
    /// <summary>
    /// 获取导出计算列定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<IExcelEntityCalculatedColumn> CalculatedColumns { get; }
    /// <summary>
    /// 获取连续分组小计定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IExcelEntityGroupSubtotal GroupSubtotal { get; }

    /// <summary>
    /// 获取每页明细小计定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IExcelEntityListFooter PageSubtotal { get; }

    /// <summary>
    /// 获取每页明细行数；未配置分页时为 <see langword="null" />。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public int? PageBreakRows { get; }
    /// <summary>
    /// 获取列表项的请求级映射配置副本。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelMappingConfiguration MappingConfiguration => _mappingConfiguration == null ? null :
        MappingConfigurationCloner.Clone(_mappingConfiguration, _mappingConfiguration.SourceKind);
    /// <summary>
    /// 获取列表项的规范化映射文档副本。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelMappingDocument MappingDocument => _mappingDocument == null ? null :
        MappingDocumentCloner.Clone(_mappingDocument);

    /// <summary>
    /// 使用运行时解析出的起始坐标创建列表区域副本。
    /// </summary>
    /// <param name="start">解析后的列表起始坐标。</param>
    /// <returns>不再依赖命名锚点的区域副本。</returns>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelEntityListRegion<TEntity> WithResolvedStart(ExcelEntityCellReference start) =>
        new ExcelEntityListRegion<TEntity>(SheetName, start, PropertyName, ItemType, Getter, Setter,
            IncludeHeader, End, _mappingConfiguration, _mappingDocument, DynamicColumnGroups,
            UnknownDynamicValues, Footer, CalculatedColumns, GroupSubtotal, PageBreakRows,
            PageSubtotal, footerAnchorName: FooterAnchorName);
}
