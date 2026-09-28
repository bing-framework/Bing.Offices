using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 固定单元格绑定描述。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityCellBinding<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCellBinding{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="reference">目标单元格坐标。</param>
    /// <param name="propertyName">实体属性名称。</param>
    /// <param name="propertyType">实体属性类型。</param>
    /// <param name="getter">从实体读取属性值的委托。</param>
    /// <param name="setter">向实体写入属性值的委托；只读属性时为 <see langword="null" />。</param>
    /// <param name="converterName">可选的命名值转换器名称。</param>
    /// <param name="mappingConfiguration">可选的请求级映射配置。</param>
    /// <param name="mappingDocument">可选的规范化映射文档。</param>
    /// <param name="anchorName">可选的模板命名锚点名称。</param>
    internal ExcelEntityCellBinding(string sheetName, ExcelEntityCellReference reference, string propertyName,
        Type propertyType, Func<object, object> getter, Action<object, object> setter, string converterName,
        ExcelMappingConfiguration mappingConfiguration = null, ExcelMappingDocument mappingDocument = null,
        string anchorName = null)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Reference = reference;
        PropertyName = propertyName ?? throw new ArgumentNullException(nameof(propertyName));
        PropertyType = propertyType ?? throw new ArgumentNullException(nameof(propertyType));
        Getter = getter ?? throw new ArgumentNullException(nameof(getter));
        Setter = setter;
        ConverterName = converterName;
        AnchorName = anchorName;
        _mappingConfiguration = mappingConfiguration == null ? null :
            MappingConfigurationCloner.Clone(mappingConfiguration, mappingConfiguration.SourceKind);
        _mappingDocument = mappingDocument == null ? null : MappingDocumentCloner.Clone(mappingDocument);
    }

    /// <summary>
    /// 保存固定单元格的请求级映射配置快照。
    /// </summary>
    private readonly ExcelMappingConfiguration _mappingConfiguration;

    /// <summary>
    /// 保存固定单元格的规范化映射文档快照。
    /// </summary>
    private readonly ExcelMappingDocument _mappingDocument;

    /// <summary>
    /// 获取目标工作表名称。
    /// </summary>
    public string SheetName { get; }
    /// <summary>
    /// 获取可选的模板命名锚点名称。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public string AnchorName { get; }
    /// <summary>
    /// 获取固定单元格的零基坐标。
    /// </summary>
    public ExcelEntityCellReference Reference { get; }
    /// <summary>
    /// 获取绑定的实体属性名称。
    /// </summary>
    public string PropertyName { get; }
    /// <summary>
    /// 获取绑定的实体属性类型。
    /// </summary>
    public Type PropertyType { get; }
    /// <summary>
    /// 获取可选的命名值转换器名称。
    /// </summary>
    public string ConverterName { get; }
    /// <summary>
    /// 获取固定单元格的请求级映射配置副本。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelMappingConfiguration MappingConfiguration => _mappingConfiguration == null ? null :
        MappingConfigurationCloner.Clone(_mappingConfiguration, _mappingConfiguration.SourceKind);
    /// <summary>
    /// 获取固定单元格的规范化映射文档副本。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelMappingDocument MappingDocument => _mappingDocument == null ? null :
        MappingDocumentCloner.Clone(_mappingDocument);
    /// <summary>
    /// 获取从实体读取属性值的已编译委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, object> Getter { get; }
    /// <summary>
    /// 获取向实体写入属性值的委托；只读属性时为 <see langword="null" />。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Action<object, object> Setter { get; }

    /// <summary>
    /// 使用运行时解析出的坐标创建固定单元格绑定副本。
    /// </summary>
    /// <param name="reference">解析后的单元格坐标。</param>
    /// <returns>不再依赖命名锚点的绑定副本。</returns>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelEntityCellBinding<TEntity> WithResolvedReference(ExcelEntityCellReference reference) =>
        new ExcelEntityCellBinding<TEntity>(SheetName, reference, PropertyName, PropertyType, Getter, Setter,
            ConverterName, _mappingConfiguration, _mappingDocument);
}
