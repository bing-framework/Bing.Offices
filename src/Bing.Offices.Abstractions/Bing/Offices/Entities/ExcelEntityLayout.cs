using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 创建单个聚合对象的 Excel 布局。
/// </summary>
public static class ExcelEntity
{
    /// <summary>
    /// 创建并校验一个不可变实体布局。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="configure">布局配置委托。</param>
    /// <returns>不可变布局。</returns>
    public static ExcelEntityLayout<TEntity> Layout<TEntity>(
        Action<ExcelEntityLayoutBuilder<TEntity>> configure)
        where TEntity : class, new()
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelEntityLayoutBuilder<TEntity>();
        configure(builder);
        return builder.Build();
    }
}

/// <summary>
/// 单个聚合对象的类型化布局构建器。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityLayoutBuilder<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 按配置顺序保存固定单元格绑定。
    /// </summary>
    private readonly List<ExcelEntityCellBinding<TEntity>> _cells =
        new List<ExcelEntityCellBinding<TEntity>>();
    /// <summary>
    /// 按配置顺序保存列表区域绑定。
    /// </summary>
    private readonly List<ExcelEntityListRegion<TEntity>> _listRegions =
        new List<ExcelEntityListRegion<TEntity>>();
    /// <summary>
    /// 按配置顺序保存固定合并区域声明。
    /// </summary>
    private readonly List<ExcelEntityMergeRegion> _merges = new List<ExcelEntityMergeRegion>();
    /// <summary>
    /// 按配置顺序保存聚合对象父子关系绑定。
    /// </summary>
    private readonly List<ExcelRelationRequest> _relations = new List<ExcelRelationRequest>();

    /// <summary>
    /// 将实体属性绑定到固定单元格。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="address">A1 单元格地址。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="converterName">可选的命名值转换器。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityLayoutBuilder<TEntity> Cell<TValue>(string sheetName, string address,
        Expression<Func<TEntity, TValue>> property, string converterName = null)
        => AddCell(sheetName, address, property, converterName, null);

    /// <summary>
    /// 将实体属性绑定到固定单元格，并配置该属性使用的映射计划覆盖。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="address">A1 单元格地址。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="configure">固定单元格映射配置委托。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityLayoutBuilder<TEntity> Cell<TValue>(string sheetName, string address,
        Expression<Func<TEntity, TValue>> property, Action<ExcelEntityCellBuilder<TValue>> configure)
        => AddCell(sheetName, address, property, null, configure);

    /// <summary>
    /// 添加固定单元格绑定并保存其映射配置快照。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="address">A1 单元格地址。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="converterName">可选的命名值转换器。</param>
    /// <param name="configure">固定单元格映射配置委托。</param>
    /// <returns>当前构建器。</returns>
    private ExcelEntityLayoutBuilder<TEntity> AddCell<TValue>(string sheetName, string address,
        Expression<Func<TEntity, TValue>> property, string converterName,
        Action<ExcelEntityCellBuilder<TValue>> configure)
    {
        if (property == null)
            throw new ArgumentNullException(nameof(property));
        var member = ExcelEntityExpression.GetProperty(property.Body);
        var getter = property.Compile();
        var setter = ExcelEntityExpression.CreateSetter<TEntity, TValue>(member);
        var cellBuilder = new ExcelEntityCellBuilder<TValue>();
        configure?.Invoke(cellBuilder);
        _cells.Add(new ExcelEntityCellBinding<TEntity>(sheetName, ExcelEntityCellReference.Parse(address),
            member.Name, typeof(TValue), value => getter((TEntity)value), setter, converterName,
            cellBuilder.MappingConfiguration, cellBuilder.MappingDocument));
        return this;
    }

    /// <summary>
    /// 声明一个固定合并区域。值只能写入左上角锚点。
    /// </summary>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="range">A1 区域地址。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityLayoutBuilder<TEntity> Merge(string sheetName, string range)
    {
        _merges.Add(new ExcelEntityMergeRegion(sheetName, ExcelEntityCellRange.Parse(range)));
        return this;
    }

    /// <summary>
    /// 将实体集合属性绑定到从起始单元格开始的列表区域。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="startAddress">列表表头或第一行的起始地址。</param>
    /// <param name="items">实体集合属性表达式。</param>
    /// <param name="configure">列表区域配置委托。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityLayoutBuilder<TEntity> ListRegion<TItem>(string sheetName, string startAddress,
        Expression<Func<TEntity, IEnumerable<TItem>>> items,
        Action<ExcelEntityListRegionBuilder<TItem>> configure = null)
        where TItem : class, new()
    {
        if (items == null)
            throw new ArgumentNullException(nameof(items));
        var member = ExcelEntityExpression.GetProperty(items.Body);
        var getter = items.Compile();
        var setter = ExcelEntityExpression.CreateObjectSetter<TEntity>(member);
        var builder = new ExcelEntityListRegionBuilder<TItem>();
        configure?.Invoke(builder);
        _listRegions.Add(new ExcelEntityListRegion<TEntity>(sheetName,
            ExcelEntityCellReference.Parse(startAddress), member.Name, typeof(TItem),
            value => (IEnumerable)(getter((TEntity)value) ?? Array.Empty<TItem>()), setter,
            builder.IncludeHeader, builder.EndAddress == null ? (ExcelEntityCellReference?)null :
            ExcelEntityCellReference.Parse(builder.EndAddress), builder.MappingConfiguration,
            builder.MappingDocument));
        return this;
    }

    /// <summary>
    /// 添加聚合对象内的父子集合关系。
    /// </summary>
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">子实体类型。</typeparam>
    /// <typeparam name="TKey">父子实体关联键类型。</typeparam>
    /// <param name="parents">聚合对象中的父实体集合属性。</param>
    /// <param name="children">聚合对象中的子实体集合属性。</param>
    /// <param name="parentKey">从父实体获取关联键的函数。</param>
    /// <param name="childKey">从子实体获取关联键的函数。</param>
    /// <param name="navigation">父实体上的子实体导航属性。</param>
    /// <param name="comparer">比较关联键的可选比较器。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityLayoutBuilder<TEntity> HasMany<TParent, TChild, TKey>(
        Expression<Func<TEntity, ICollection<TParent>>> parents,
        Expression<Func<TEntity, ICollection<TChild>>> children,
        Func<TParent, TKey> parentKey,
        Func<TChild, TKey> childKey,
        Expression<Func<TParent, ICollection<TChild>>> navigation,
        IEqualityComparer<TKey> comparer = null)
        where TParent : class
        where TChild : class
    {
        if (parents == null || children == null || parentKey == null || childKey == null || navigation == null)
            throw new ArgumentNullException("关系配置参数不能为空。");
        _relations.Add(ExcelRelationRequest.Create(parents, children, parentKey, childKey, navigation, comparer));
        return this;
    }

    /// <summary>
    /// 完成布局构建并执行冲突校验。
    /// </summary>
    /// <returns>不可变实体布局。</returns>
    public ExcelEntityLayout<TEntity> Build()
    {
        if (_cells.Count == 0 && _listRegions.Count == 0)
            throw new InvalidOperationException("实体布局至少需要一个固定单元格或列表区域。");
        foreach (var cell in _cells)
            ExcelEntityLayoutValidation.ValidateSheetName(cell.SheetName);
        foreach (var region in _listRegions)
            ExcelEntityLayoutValidation.ValidateSheetName(region.SheetName);
        foreach (var merge in _merges)
            ExcelEntityLayoutValidation.ValidateSheetName(merge.SheetName);

        foreach (var region in _listRegions)
        {
            if (region.End.HasValue && !ExcelEntityLayoutValidation.Contains(region.Start, region.End.Value))
                throw new ArgumentException($"列表区域结束地址不能位于起始地址之前: {region.SheetName}!{region.Start.Address}");
        }

        var normalizedCells = _cells.Select(cell =>
        {
            var merge = _merges.FirstOrDefault(item => string.Equals(item.SheetName, cell.SheetName,
                StringComparison.OrdinalIgnoreCase) && item.Range.Contains(cell.Reference));
            if (merge == null || merge.Range.First.Equals(cell.Reference))
                return cell;
            return new ExcelEntityCellBinding<TEntity>(cell.SheetName, merge.Range.First, cell.PropertyName,
                cell.PropertyType, cell.Getter, cell.Setter, cell.ConverterName,
                cell.MappingConfiguration, cell.MappingDocument);
        }).ToArray();

        var duplicateCell = normalizedCells.GroupBy(cell => (cell.SheetName, cell.Reference.Row, cell.Reference.Column),
                StringTupleComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateCell != null)
            throw new ArgumentException($"实体布局包含重复固定单元格: {duplicateCell.Key.SheetName}!{duplicateCell.First().Reference.Address}");

        var duplicateMerge = _merges.GroupBy(merge => (merge.SheetName, merge.Range.First.Row,
                merge.Range.First.Column, merge.Range.Last.Row, merge.Range.Last.Column),
                StringTupleComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateMerge != null)
            throw new ArgumentException($"实体布局包含重复合并区域: {duplicateMerge.First().SheetName}!{duplicateMerge.First().Range.Address}");

        foreach (var mergePair in _merges.SelectMany((left, index) => _merges.Skip(index + 1)
                     .Where(right => string.Equals(left.SheetName, right.SheetName,
                         StringComparison.OrdinalIgnoreCase)
                         && ExcelEntityLayoutValidation.Overlaps(left.Range, right.Range))))
            throw new ArgumentException($"实体布局包含重叠合并区域: {mergePair.SheetName}!{mergePair.Range.Address}");

        foreach (var cell in normalizedCells)
        {
            var merge = _merges.FirstOrDefault(item => string.Equals(item.SheetName, cell.SheetName,
                StringComparison.OrdinalIgnoreCase) && item.Range.Contains(cell.Reference));
            foreach (var region in _listRegions.Where(item => string.Equals(item.SheetName, cell.SheetName,
                         StringComparison.OrdinalIgnoreCase)))
            {
                if (ExcelEntityLayoutValidation.Contains(region.Start, region.End, cell.Reference))
                    throw new ArgumentException($"固定单元格与列表区域重叠: {cell.SheetName}!{cell.Reference.Address}");
            }
        }

        foreach (var pair in _listRegions.SelectMany((left, index) => _listRegions.Skip(index + 1)
                     .Where(right => string.Equals(left.SheetName, right.SheetName,
                         StringComparison.OrdinalIgnoreCase)
                         && ExcelEntityLayoutValidation.Overlaps(left, right))))
            throw new ArgumentException($"实体布局包含重叠列表区域: {pair.SheetName}!{pair.Start.Address}");

        foreach (var region in _listRegions)
        {
            foreach (var merge in _merges.Where(item => string.Equals(item.SheetName, region.SheetName,
                         StringComparison.OrdinalIgnoreCase)))
            {
                if (ExcelEntityLayoutValidation.Overlaps(region, merge.Range))
                    throw new ArgumentException($"列表区域与合并区域重叠: {region.SheetName}!{region.Start.Address}");
            }
        }

        return new ExcelEntityLayout<TEntity>(normalizedCells, _listRegions.ToArray(), _merges.ToArray(),
            _relations.ToArray());
    }
}

/// <summary>
/// 固定单元格的映射配置构建器。
/// </summary>
/// <typeparam name="TValue">固定单元格绑定的属性类型。</typeparam>
public sealed class ExcelEntityCellBuilder<TValue>
{
    /// <summary>
    /// 获取固定单元格使用的请求级映射配置快照。
    /// </summary>
    internal ExcelMappingConfiguration MappingConfiguration { get; private set; }

    /// <summary>
    /// 获取固定单元格使用的规范化映射文档快照。
    /// </summary>
    internal ExcelMappingDocument MappingDocument { get; private set; }

    /// <summary>
    /// 设置固定单元格使用的请求级映射配置。
    /// </summary>
    /// <param name="configuration">映射配置。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCellBuilder<TValue> Mapping(ExcelMappingConfiguration configuration)
    {
        MappingConfiguration = configuration == null ? null :
            MappingConfigurationCloner.Clone(configuration, MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置固定单元格使用的规范化映射文档。
    /// </summary>
    /// <param name="document">映射文档。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCellBuilder<TValue> Mapping(ExcelMappingDocument document)
    {
        MappingDocument = document == null ? null : MappingDocumentCloner.Clone(document);
        return this;
    }
}

/// <summary>
/// 列表区域构建器。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
public sealed class ExcelEntityListRegionBuilder<TItem> where TItem : class, new()
{
    /// <summary>
    /// 获取是否在列表区域起始行包含表头。
    /// </summary>
    internal bool IncludeHeader { get; private set; } = true;
    /// <summary>
    /// 获取列表区域可选右下角的 A1 地址。
    /// </summary>
    internal string EndAddress { get; private set; }
    /// <summary>
    /// 获取列表项的请求级映射配置。
    /// </summary>
    internal ExcelMappingConfiguration MappingConfiguration { get; private set; }
    /// <summary>
    /// 获取列表项的规范化映射文档。
    /// </summary>
    internal ExcelMappingDocument MappingDocument { get; private set; }

    /// <summary>
    /// 设置是否在列表区域起始行写入映射表头。
    /// </summary>
    /// <param name="includeHeader">是否写入表头。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> Header(bool includeHeader = true)
    {
        IncludeHeader = includeHeader;
        return this;
    }

    /// <summary>
    /// 设置列表区域的可选右下角边界。
    /// </summary>
    /// <param name="address">A1 右下角地址。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> End(string address)
    {
        EndAddress = ExcelEntityCellReference.Parse(address).Address;
        return this;
    }

    /// <summary>
    /// 设置列表项的请求级映射配置。
    /// </summary>
    /// <param name="configuration">映射配置。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> Mapping(ExcelMappingConfiguration configuration)
    {
        MappingConfiguration = configuration == null ? null :
            MappingConfigurationCloner.Clone(configuration, MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置列表项的规范化映射文档。
    /// </summary>
    /// <param name="document">映射文档。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListRegionBuilder<TItem> Mapping(ExcelMappingDocument document)
    {
        MappingDocument = document == null ? null : MappingDocumentCloner.Clone(document);
        return this;
    }
}

/// <summary>
/// 已完成校验的不可变实体布局。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityLayout<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityLayout{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="cells">固定单元格绑定快照。</param>
    /// <param name="listRegions">列表区域绑定快照。</param>
    /// <param name="merges">固定合并区域快照。</param>
    /// <param name="relations">父子关系绑定快照。</param>
    internal ExcelEntityLayout(IReadOnlyList<ExcelEntityCellBinding<TEntity>> cells,
        IReadOnlyList<ExcelEntityListRegion<TEntity>> listRegions,
        IReadOnlyList<ExcelEntityMergeRegion> merges,
        IReadOnlyList<ExcelRelationRequest> relations)
    {
        Cells = cells;
        ListRegions = listRegions;
        Merges = merges;
        Relations = Array.AsReadOnly((relations ?? Array.Empty<ExcelRelationRequest>()).ToArray());
    }

    /// <summary>
    /// 获取固定单元格绑定。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelEntityCellBinding<TEntity>> Cells { get; }

    /// <summary>
    /// 获取列表区域绑定。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelEntityListRegion<TEntity>> ListRegions { get; }

    /// <summary>
    /// 获取合并区域声明。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelEntityMergeRegion> Merges { get; }

    /// <summary>
    /// 获取聚合对象父子关系绑定。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelRelationRequest> Relations { get; }
}

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
    internal ExcelEntityCellBinding(string sheetName, ExcelEntityCellReference reference, string propertyName,
        Type propertyType, Func<object, object> getter, Action<object, object> setter, string converterName,
        ExcelMappingConfiguration mappingConfiguration = null, ExcelMappingDocument mappingDocument = null)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Reference = reference;
        PropertyName = propertyName ?? throw new ArgumentNullException(nameof(propertyName));
        PropertyType = propertyType ?? throw new ArgumentNullException(nameof(propertyType));
        Getter = getter ?? throw new ArgumentNullException(nameof(getter));
        Setter = setter;
        ConverterName = converterName;
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
}

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
    internal ExcelEntityListRegion(string sheetName, ExcelEntityCellReference start, string propertyName,
        Type itemType, Func<object, IEnumerable> getter, Action<object, object> setter, bool includeHeader,
        ExcelEntityCellReference? end, ExcelMappingConfiguration mappingConfiguration,
        ExcelMappingDocument mappingDocument)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Start = start;
        PropertyName = propertyName ?? throw new ArgumentNullException(nameof(propertyName));
        ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
        Getter = getter ?? throw new ArgumentNullException(nameof(getter));
        Setter = setter;
        IncludeHeader = includeHeader;
        End = end;
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
}

/// <summary>
/// 固定合并区域描述。
/// </summary>
public sealed class ExcelEntityMergeRegion
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityMergeRegion" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="range">合并区域的零基矩形坐标。</param>
    internal ExcelEntityMergeRegion(string sheetName, ExcelEntityCellRange range)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Range = range;
    }

    /// <summary>
    /// 获取目标工作表名称。
    /// </summary>
    public string SheetName { get; }
    /// <summary>
    /// 获取待读取或写入的合并区域。
    /// </summary>
    public ExcelEntityCellRange Range { get; }
}

/// <summary>
/// 零基 Excel 单元格坐标。
/// </summary>
public readonly struct ExcelEntityCellReference : IEquatable<ExcelEntityCellReference>
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCellReference" /> 类型的实例。
    /// </summary>
    /// <param name="row">从零开始的行索引。</param>
    /// <param name="column">从零开始的列索引。</param>
    internal ExcelEntityCellReference(int row, int column)
    {
        if (row < 0)
            throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0)
            throw new ArgumentOutOfRangeException(nameof(column));
        Row = row;
        Column = column;
    }

    /// <summary>
    /// 获取从零开始的行索引。
    /// </summary>
    public int Row { get; }
    /// <summary>
    /// 获取从零开始的列索引。
    /// </summary>
    public int Column { get; }
    /// <summary>
    /// 获取与当前零基坐标对应的 A1 地址。
    /// </summary>
    public string Address => ExcelEntityAddress.Format(Row, Column);

    /// <summary>
    /// 解析单个 A1 单元格地址。
    /// </summary>
    /// <param name="address">要解析的 A1 地址。</param>
    /// <returns>解析得到的零基单元格坐标。</returns>
    public static ExcelEntityCellReference Parse(string address) => ExcelEntityAddress.Parse(address);
    /// <inheritdoc />
    public bool Equals(ExcelEntityCellReference other) => Row == other.Row && Column == other.Column;
    /// <inheritdoc />
    public override bool Equals(object obj) => obj is ExcelEntityCellReference other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked((Row * 397) ^ Column);
    /// <inheritdoc />
    public override string ToString() => Address;
}

/// <summary>
/// 零基 Excel 矩形区域。
/// </summary>
public readonly struct ExcelEntityCellRange : IEquatable<ExcelEntityCellRange>
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCellRange" /> 类型的实例。
    /// </summary>
    /// <param name="first">区域左上角坐标。</param>
    /// <param name="last">区域右下角坐标。</param>
    internal ExcelEntityCellRange(ExcelEntityCellReference first, ExcelEntityCellReference last)
    {
        if (last.Row < first.Row || last.Column < first.Column)
            throw new ArgumentException("区域结束坐标不能位于起始坐标之前。");
        First = first;
        Last = last;
    }

    /// <summary>
    /// 获取区域的左上角坐标。
    /// </summary>
    public ExcelEntityCellReference First { get; }
    /// <summary>
    /// 获取区域的右下角坐标。
    /// </summary>
    public ExcelEntityCellReference Last { get; }
    /// <summary>
    /// 获取与当前零基坐标对应的 A1 区域地址。
    /// </summary>
    public string Address => First.Equals(Last) ? First.Address : $"{First.Address}:{Last.Address}";
    /// <summary>
    /// 判断指定坐标是否位于当前区域内。
    /// </summary>
    /// <param name="reference">要检查的零基单元格坐标。</param>
    /// <returns>坐标位于区域内时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    public bool Contains(ExcelEntityCellReference reference) => reference.Row >= First.Row && reference.Row <= Last.Row
        && reference.Column >= First.Column && reference.Column <= Last.Column;
    /// <summary>
    /// 解析 A1 矩形区域地址。
    /// </summary>
    /// <param name="range">要解析的单元格或矩形区域地址。</param>
    /// <returns>解析得到的零基矩形区域。</returns>
    public static ExcelEntityCellRange Parse(string range) => ExcelEntityAddress.ParseRange(range);
    /// <inheritdoc />
    public bool Equals(ExcelEntityCellRange other) => First.Equals(other.First) && Last.Equals(other.Last);
    /// <inheritdoc />
    public override bool Equals(object obj) => obj is ExcelEntityCellRange other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked((First.GetHashCode() * 397) ^ Last.GetHashCode());
    /// <inheritdoc />
    public override string ToString() => Address;
}

/// <summary>
/// 负责解析实体属性表达式并创建已编译访问委托。
/// </summary>
internal static class ExcelEntityExpression
{
    /// <summary>
    /// 从直接属性访问表达式中取得属性元数据。
    /// </summary>
    /// <param name="expression">待解析的表达式主体。</param>
    /// <returns>表达式直接访问的属性元数据。</returns>
    internal static System.Reflection.PropertyInfo GetProperty(Expression expression)
    {
        if (expression is UnaryExpression unary && unary.NodeType == ExpressionType.Convert)
            expression = unary.Operand;
        if (!(expression is MemberExpression member) || !(member.Member is System.Reflection.PropertyInfo property)
            || member.Expression == null || member.Expression.NodeType != ExpressionType.Parameter)
            throw new ArgumentException("实体布局属性必须是直接属性表达式。", nameof(expression));
        return property;
    }

    /// <summary>
    /// 为可写属性创建接受对象参数的写入委托。
    /// </summary>
    /// <typeparam name="TEntity">声明目标属性的实体类型。</typeparam>
    /// <param name="property">待写入的属性元数据。</param>
    /// <returns>已编译的写入委托；属性不可写时返回 <see langword="null" />。</returns>
    internal static Action<object, object> CreateObjectSetter<TEntity>(System.Reflection.PropertyInfo property)
    {
        if (!property.CanWrite)
            return null;
        var target = Expression.Parameter(typeof(object), "target");
        var value = Expression.Parameter(typeof(object), "value");
        var assignment = Expression.Assign(Expression.Property(Expression.Convert(target, typeof(TEntity)), property),
            Expression.Convert(value, property.PropertyType));
        return Expression.Lambda<Action<object, object>>(assignment, target, value).Compile();
    }

    /// <summary>
    /// 为指定值类型的可写属性创建对象写入委托。
    /// </summary>
    /// <typeparam name="TEntity">声明目标属性的实体类型。</typeparam>
    /// <typeparam name="TValue">属性值类型。</typeparam>
    /// <param name="property">待写入的属性元数据。</param>
    /// <returns>已编译的写入委托；属性不可写时返回 <see langword="null" />。</returns>
    internal static Action<object, object> CreateSetter<TEntity, TValue>(System.Reflection.PropertyInfo property)
        => CreateObjectSetter<TEntity>(property);
}

/// <summary>
/// 提供 Excel A1 单元格和矩形区域地址的解析与格式化。
/// </summary>
internal static class ExcelEntityAddress
{
    /// <summary>
    /// XLSX 工作表允许的最大行数。
    /// </summary>
    private const int MaxRows = 1048576;
    /// <summary>
    /// XLSX 工作表允许的最大列数。
    /// </summary>
    private const int MaxColumns = 16384;

    /// <summary>
    /// 解析并验证单个 A1 单元格地址。
    /// </summary>
    /// <param name="address">待解析的 A1 地址。</param>
    /// <returns>与地址对应的零基单元格坐标。</returns>
    internal static ExcelEntityCellReference Parse(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("单元格地址不能为空。", nameof(address));
        var text = address.Trim().Replace("$", string.Empty);
        var split = 0;
        while (split < text.Length && ((text[split] >= 'A' && text[split] <= 'Z')
            || (text[split] >= 'a' && text[split] <= 'z')))
            split++;
        if (split == 0 || split == text.Length || !int.TryParse(text.Substring(split),
                System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture,
                out var row) || row <= 0 || row > MaxRows)
            throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
        var column = 0;
        foreach (var character in text.Substring(0, split).ToUpperInvariant())
        {
            if (character < 'A' || character > 'Z')
                throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
            if (column > (MaxColumns - (character - 'A' + 1)) / 26)
                throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
            column = column * 26 + character - 'A' + 1;
        }
        if (column <= 0 || column > MaxColumns)
            throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
        return new ExcelEntityCellReference(row - 1, column - 1);
    }

    /// <summary>
    /// 解析并验证 A1 单元格或矩形区域地址。
    /// </summary>
    /// <param name="range">待解析的 A1 地址。</param>
    /// <returns>与地址对应的零基矩形区域。</returns>
    internal static ExcelEntityCellRange ParseRange(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentException("区域地址不能为空。", nameof(range));
        var parts = range.Split(':');
        if (parts.Length > 2)
            throw new ArgumentException($"无效的 A1 区域地址: {range}", nameof(range));
        var first = Parse(parts[0]);
        var last = parts.Length == 1 ? first : Parse(parts[1]);
        return new ExcelEntityCellRange(first, last);
    }

    /// <summary>
    /// 将零基行列坐标格式化为 A1 单元格地址。
    /// </summary>
    /// <param name="row">从零开始的行索引。</param>
    /// <param name="column">从零开始的列索引。</param>
    /// <returns>格式化后的 A1 单元格地址。</returns>
    internal static string Format(int row, int column)
    {
        var value = column + 1;
        var letters = string.Empty;
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            value = (value - 1) / 26;
        }
        return letters + (row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// 提供实体布局名称、包含关系和区域重叠的验证操作。
/// </summary>
internal static class ExcelEntityLayoutValidation
{
    /// <summary>
    /// 验证工作表名称是否满足 Excel 长度和字符约束。
    /// </summary>
    /// <param name="name">待验证的工作表名称。</param>
    internal static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31 || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称无效: {name}", nameof(name));
    }

    /// <summary>
    /// 判断单元格是否位于指定起始坐标和可选结束坐标内。
    /// </summary>
    /// <param name="start">区域起始坐标。</param>
    /// <param name="end">可选的区域右下角坐标。</param>
    /// <param name="cell">待检查的单元格坐标。</param>
    /// <returns>单元格位于区域内时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Contains(ExcelEntityCellReference start, ExcelEntityCellReference? end,
        ExcelEntityCellReference cell) => cell.Row >= start.Row && cell.Column >= start.Column
        && (!end.HasValue || (cell.Row <= end.Value.Row && cell.Column <= end.Value.Column));

    /// <summary>
    /// 判断结束坐标是否位于起始坐标的右下方向。
    /// </summary>
    /// <param name="start">区域起始坐标。</param>
    /// <param name="end">区域结束坐标。</param>
    /// <returns>结束坐标不早于起始坐标时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Contains(ExcelEntityCellReference start, ExcelEntityCellReference end) =>
        end.Row >= start.Row && end.Column >= start.Column;

    /// <summary>
    /// 判断两个有界矩形区域是否存在交集。
    /// </summary>
    /// <param name="left">第一个矩形区域。</param>
    /// <param name="right">第二个矩形区域。</param>
    /// <returns>区域存在交集时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Overlaps(ExcelEntityCellRange left, ExcelEntityCellRange right) =>
        left.First.Row <= right.Last.Row && right.First.Row <= left.Last.Row
        && left.First.Column <= right.Last.Column && right.First.Column <= left.Last.Column;

    /// <summary>
    /// 判断两个列表区域是否存在可确定的交集。
    /// </summary>
    /// <typeparam name="TEntity">列表区域所属的聚合对象类型。</typeparam>
    /// <param name="left">第一个列表区域。</param>
    /// <param name="right">第二个列表区域。</param>
    /// <returns>已知边界或起点表明区域重叠时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Overlaps<TEntity>(ExcelEntityListRegion<TEntity> left,
        ExcelEntityListRegion<TEntity> right) where TEntity : class, new()
    {
        if (!left.End.HasValue && !right.End.HasValue)
            return true;
        if (left.End.HasValue && right.End.HasValue)
            return Overlaps(new ExcelEntityCellRange(left.Start, left.End.Value),
                new ExcelEntityCellRange(right.Start, right.End.Value));
        if (left.Start.Equals(right.Start))
            return true;
        return left.End.HasValue
            ? Contains(left.Start, left.End, right.Start)
            : right.End.HasValue && Contains(right.Start, right.End, left.Start);
    }

    /// <summary>
    /// 判断列表区域与固定矩形区域是否存在可确定的交集。
    /// </summary>
    /// <typeparam name="TEntity">列表区域所属的聚合对象类型。</typeparam>
    /// <param name="region">待检查的列表区域。</param>
    /// <param name="range">待检查的固定矩形区域。</param>
    /// <returns>区域存在可确定交集时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Overlaps<TEntity>(ExcelEntityListRegion<TEntity> region,
        ExcelEntityCellRange range) where TEntity : class, new()
    {
        if (region.End.HasValue)
            return Overlaps(new ExcelEntityCellRange(region.Start, region.End.Value), range);
        return region.Start.Row <= range.Last.Row && region.Start.Column <= range.Last.Column;
    }
}

/// <summary>
/// 为包含工作表名称和整数坐标的元组提供忽略大小写的相等性。
/// </summary>
internal sealed class StringTupleComparer : IEqualityComparer<(string Sheet, int A, int B)>,
    IEqualityComparer<(string Sheet, int A, int B, int C, int D)>
{
    /// <summary>
    /// 获取全局共享的忽略大小写比较器。
    /// </summary>
    internal static readonly StringTupleComparer OrdinalIgnoreCase = new StringTupleComparer();
    /// <inheritdoc />
    public bool Equals((string Sheet, int A, int B) x, (string Sheet, int A, int B) y) =>
        string.Equals(x.Sheet, y.Sheet, StringComparison.OrdinalIgnoreCase) && x.A == y.A && x.B == y.B;
    /// <inheritdoc />
    public int GetHashCode((string Sheet, int A, int B) value) =>
        Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(value.Sheet ?? string.Empty), value.A, value.B);
    /// <inheritdoc />
    public bool Equals((string Sheet, int A, int B, int C, int D) x, (string Sheet, int A, int B, int C, int D) y) =>
        string.Equals(x.Sheet, y.Sheet, StringComparison.OrdinalIgnoreCase) && x.A == y.A && x.B == y.B
        && x.C == y.C && x.D == y.D;
    /// <inheritdoc />
    public int GetHashCode((string Sheet, int A, int B, int C, int D) value) => Combine(
        StringComparer.OrdinalIgnoreCase.GetHashCode(value.Sheet ?? string.Empty), value.A, value.B, value.C, value.D);

    /// <summary>
    /// 按顺序组合多个整数哈希值。
    /// </summary>
    /// <param name="values">待组合的整数哈希值。</param>
    /// <returns>合并后的哈希值。</returns>
    private static int Combine(params int[] values)
    {
        unchecked
        {
            var hash = 17;
            foreach (var value in values)
                hash = hash * 31 + value;
            return hash;
        }
    }
}
