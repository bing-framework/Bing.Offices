using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
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

    /// <summary>
    /// 根据实体属性上的固定单元格特性创建布局。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="configure">在特性布局基础上继续配置的委托。</param>
    /// <returns>不可变布局。</returns>
    /// <remarks>
    /// 特性只扫描公开实例属性；列表区域、合并区域和映射覆盖仍通过委托配置。
    /// </remarks>
    public static ExcelEntityLayout<TEntity> LayoutFromAttributes<TEntity>(
        Action<ExcelEntityLayoutBuilder<TEntity>> configure = null)
        where TEntity : class, new()
    {
        var builder = new ExcelEntityLayoutBuilder<TEntity>();
        builder.CellsFromAttributes();
        configure?.Invoke(builder);
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
    /// 从实体公开实例属性上的固定单元格特性添加绑定。
    /// </summary>
    /// <returns>当前构建器。</returns>
    public ExcelEntityLayoutBuilder<TEntity> CellsFromAttributes()
    {
        var properties = typeof(TEntity).GetProperties(System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.Public)
            .Where(property => property.GetCustomAttributes(typeof(ExcelEntityCellAttribute), true).Length > 0)
            .OrderBy(property => property.Name, StringComparer.Ordinal);
        foreach (var property in properties)
        {
            var attributes = property.GetCustomAttributes(typeof(ExcelEntityCellAttribute), true)
                .Cast<ExcelEntityCellAttribute>().ToArray();
            if (attributes.Length != 1)
                throw new ArgumentException($"属性只能配置一个 ExcelEntityCellAttribute: {property.Name}");
            var attribute = attributes[0];
            var parameter = Expression.Parameter(typeof(TEntity), "entity");
            var getterBody = Expression.Convert(Expression.Property(parameter, property), typeof(object));
            var getter = Expression.Lambda<Func<TEntity, object>>(getterBody, parameter).Compile();
            _cells.Add(new ExcelEntityCellBinding<TEntity>(attribute.SheetName,
                ExcelEntityCellReference.Parse(attribute.Address), property.Name, property.PropertyType,
                value => getter((TEntity)value), ExcelEntityExpression.CreateObjectSetter<TEntity>(property),
                attribute.ConverterName));
        }
        return this;
    }

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
    /// 将实体属性绑定到模板命名锚点。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">命名锚点所属的工作表名称。</param>
    /// <param name="anchorName">模板中的命名范围名称；必须解析为单个单元格。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="converterName">可选的命名值转换器。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 命名锚点在模板预检阶段解析；非模板导出也要求工作簿已包含该名称。
    /// </remarks>
    public ExcelEntityLayoutBuilder<TEntity> CellNamed<TValue>(string sheetName, string anchorName,
        Expression<Func<TEntity, TValue>> property, string converterName = null)
        => AddCell(sheetName, null, property, converterName, null, anchorName);

    /// <summary>
    /// 将实体属性绑定到固定单元格。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="address">A1 单元格地址。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="configure">固定单元格映射配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>配置委托可覆盖该属性的映射计划。</remarks>
    public ExcelEntityLayoutBuilder<TEntity> Cell<TValue>(string sheetName, string address,
        Expression<Func<TEntity, TValue>> property, Action<ExcelEntityCellBuilder<TValue>> configure)
        => AddCell(sheetName, address, property, null, configure);

    /// <summary>
    /// 将实体属性绑定到模板命名锚点。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">命名锚点所属的工作表名称。</param>
    /// <param name="anchorName">模板中的命名范围名称；必须解析为单个单元格。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="configure">固定单元格映射配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>配置委托可覆盖该属性的映射计划；锚点在模板预检阶段解析。</remarks>
    public ExcelEntityLayoutBuilder<TEntity> CellNamed<TValue>(string sheetName, string anchorName,
        Expression<Func<TEntity, TValue>> property, Action<ExcelEntityCellBuilder<TValue>> configure)
        => AddCell(sheetName, null, property, null, configure, anchorName);

    /// <summary>
    /// 添加固定单元格绑定并保存其映射配置快照。
    /// </summary>
    /// <typeparam name="TValue">属性类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="address">A1 单元格地址。</param>
    /// <param name="property">实体属性表达式。</param>
    /// <param name="converterName">可选的命名值转换器。</param>
    /// <param name="configure">固定单元格映射配置委托。</param>
    /// <param name="anchorName">可选的模板命名锚点名称。</param>
    /// <returns>当前构建器。</returns>
    private ExcelEntityLayoutBuilder<TEntity> AddCell<TValue>(string sheetName, string address,
        Expression<Func<TEntity, TValue>> property, string converterName,
        Action<ExcelEntityCellBuilder<TValue>> configure, string anchorName = null)
    {
        if (property == null)
            throw new ArgumentNullException(nameof(property));
        var member = ExcelEntityExpression.GetProperty(property.Body);
        var getter = property.Compile();
        var setter = ExcelEntityExpression.CreateSetter<TEntity, TValue>(member);
        var cellBuilder = new ExcelEntityCellBuilder<TValue>();
        configure?.Invoke(cellBuilder);
        if (anchorName == null && address == null)
            throw new ArgumentException("固定单元格必须指定地址或命名锚点。", nameof(address));
        _cells.Add(new ExcelEntityCellBinding<TEntity>(sheetName,
            address == null ? ExcelEntityCellReference.Parse("A1") : ExcelEntityCellReference.Parse(address),
            member.Name, typeof(TValue), value => getter((TEntity)value), setter, converterName,
            cellBuilder.MappingConfiguration, cellBuilder.MappingDocument, anchorName));
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
        _listRegions.Add(CreateListRegion(sheetName, startAddress, null, member, getter, setter, builder));
        return this;
    }

    /// <summary>
    /// 将实体集合属性绑定到模板命名锚点开始的列表区域。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="sheetName">命名锚点所属的工作表名称。</param>
    /// <param name="anchorName">模板中的命名范围名称；必须解析为单个单元格。</param>
    /// <param name="items">实体集合属性表达式。</param>
    /// <param name="configure">列表区域配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 命名锚点提供列表区域左上角；列表边界由运行时数据和 Footer 预检确定。
    /// </remarks>
    public ExcelEntityLayoutBuilder<TEntity> ListRegionNamed<TItem>(string sheetName, string anchorName,
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
        if (builder.EndAddress != null)
            throw new ArgumentException("命名锚点列表区域不能配置绝对 End 地址。", nameof(configure));
        _listRegions.Add(CreateListRegion(sheetName, null, anchorName, member, getter, setter, builder));
        return this;
    }

    /// <summary>
    /// 创建列表区域绑定并检查尾部公式配置。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="startAddress">列表起始地址。</param>
    /// <param name="anchorName">可选的模板命名锚点。</param>
    /// <param name="member">实体集合属性。</param>
    /// <param name="getter">集合读取委托。</param>
    /// <param name="setter">集合写入委托。</param>
    /// <param name="builder">列表配置构建器。</param>
    /// <returns>列表区域绑定。</returns>
    private static ExcelEntityListRegion<TEntity> CreateListRegion<TItem>(string sheetName,
        string startAddress, string anchorName, System.Reflection.PropertyInfo member,
        Func<TEntity, IEnumerable<TItem>> getter, Action<object, object> setter,
        ExcelEntityListRegionBuilder<TItem> builder) where TItem : class, new()
    {
        if (builder.FooterDefinition?.ContainsContiguousSum == true &&
            (builder.PageSubtotalDefinition != null || builder.GroupSubtotalDefinition != null))
            throw new ArgumentException("最终尾部跨越中间小计行，不能使用连续明细求和公式。");
        return new ExcelEntityListRegion<TEntity>(sheetName,
            startAddress == null ? ExcelEntityCellReference.Parse("A1") :
                ExcelEntityCellReference.Parse(startAddress), member.Name, typeof(TItem),
            value => (IEnumerable)(getter((TEntity)value) ?? Array.Empty<TItem>()), setter,
            builder.IncludeHeader, builder.EndAddress == null ? (ExcelEntityCellReference?)null :
            ExcelEntityCellReference.Parse(builder.EndAddress), builder.MappingConfiguration,
            builder.MappingDocument, builder.DynamicColumnGroups, builder.UnknownDynamicValuePolicy,
            builder.FooterDefinition, builder.CalculatedColumns, builder.GroupSubtotalDefinition,
            builder.PageBreakRows, builder.PageSubtotalDefinition, anchorName, builder.FooterAnchorName);
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

        foreach (var cell in _cells.Where(item => item.AnchorName != null))
            ExcelEntityLayoutValidation.ValidateAnchorName(cell.AnchorName);
        foreach (var region in _listRegions.Where(item => item.AnchorName != null))
            ExcelEntityLayoutValidation.ValidateAnchorName(region.AnchorName);
        foreach (var region in _listRegions.Where(item => item.FooterAnchorName != null))
            ExcelEntityLayoutValidation.ValidateAnchorName(region.FooterAnchorName);

        var duplicateAnchors = _cells.Where(item => item.AnchorName != null)
            .Select(item => (item.SheetName, item.AnchorName))
            .Concat(_listRegions.Where(item => item.AnchorName != null)
                .Select(item => (item.SheetName, item.AnchorName)))
            .GroupBy(item => (SheetName: item.SheetName?.ToUpperInvariant(),
                AnchorName: item.AnchorName?.ToUpperInvariant()))
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateAnchors != null)
            throw new ArgumentException($"实体布局包含重复命名锚点: {duplicateAnchors.First().SheetName}!{duplicateAnchors.First().AnchorName}");
        var footerAnchors = _listRegions.Where(region => region.FooterAnchorName != null)
            .Select(region => region.FooterAnchorName).ToArray();
        var allAnchors = _cells.Where(cell => cell.AnchorName != null).Select(cell => cell.AnchorName)
            .Concat(_listRegions.Where(region => region.AnchorName != null).Select(region => region.AnchorName));
        if (footerAnchors.Concat(allAnchors).GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
            .Any(group => group.Count() > 1))
            throw new ArgumentException("最终尾部命名锚点与实体布局中的其他锚点重复。");

        foreach (var region in _listRegions)
        {
            if (region.End.HasValue && !ExcelEntityLayoutValidation.Contains(region.Start, region.End.Value))
                throw new ArgumentException($"列表区域结束地址不能位于起始地址之前: {region.SheetName}!{region.Start.Address}");
            if (region.DynamicColumnGroups.Count > 0 && region.MappingConfiguration != null &&
                region.MappingConfiguration.DynamicColumns != null &&
                region.MappingConfiguration.DynamicColumns.Count > 0)
                throw new ArgumentException($"列表区域不能同时配置显式动态列组和映射动态列: {region.SheetName}!{region.Start.Address}");
            if (region.DynamicColumnGroups.Count > 0 && region.MappingDocument != null &&
                ((region.MappingDocument.Import?.DynamicColumns?.Count ?? 0) > 0
                    || (region.MappingDocument.Export?.DynamicColumns?.Count ?? 0) > 0))
                throw new ArgumentException($"列表区域不能同时配置显式动态列组和映射动态列: {region.SheetName}!{region.Start.Address}");
            var groupKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var dynamicNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var group in region.DynamicColumnGroups)
            {
                if (!groupKeys.Add(group.GroupKey) || !dynamicNames.Add(group.GroupKey))
                    throw new ArgumentException($"列表区域包含重复动态列组: {group.GroupKey}");
                foreach (var definition in group.Definitions)
                {
                    if (definition == null || string.IsNullOrWhiteSpace(definition.Key)
                        || string.IsNullOrWhiteSpace(definition.Title)
                        || !dynamicNames.Add(definition.Key)
                        || !dynamicNames.Add(definition.Title))
                        throw new ArgumentException("列表区域动态列 Key 或标题无效或重复。");
                    foreach (var alias in definition.Aliases ?? Array.Empty<string>())
                    {
                        if (string.IsNullOrWhiteSpace(alias) || !dynamicNames.Add(alias))
                            throw new ArgumentException($"列表区域动态列标题别名无效或重复: {alias}");
                    }
                }
            }
            foreach (var calculated in region.CalculatedColumns)
            {
                if (calculated == null || string.IsNullOrWhiteSpace(calculated.Key)
                    || string.IsNullOrWhiteSpace(calculated.Title)
                    || !dynamicNames.Add(calculated.Key)
                    || !dynamicNames.Add(calculated.Title))
                    throw new ArgumentException("列表区域计算列 Key 或标题无效或重复。");
                if (calculated.Placement?.BeforeKey != null
                    && calculated.Placement.AfterKey != null)
                    throw new ArgumentException($"计算列位置不能同时指定 Before 和 After: {calculated.Key}");
            }
            if (region.GroupSubtotal != null && region.Footer != null
                && string.Equals(region.GroupSubtotal.Footer.MarkerText, region.Footer.MarkerText,
                    StringComparison.Ordinal))
                throw new ArgumentException("分组小计与最终尾部不能使用相同标记文本。");
            if (region.PageSubtotal != null && !region.PageBreakRows.HasValue)
                throw new ArgumentException("分页小计必须同时配置分页行数。");
            if (region.PageSubtotal != null && region.GroupSubtotal != null)
                throw new ArgumentException("分页小计不能同时配置连续分组小计。");
            if (region.PageSubtotal != null && region.Footer != null
                && string.Equals(region.PageSubtotal.MarkerText, region.Footer.MarkerText,
                    StringComparison.Ordinal))
                throw new ArgumentException("分页小计与最终尾部不能使用相同标记文本。");
            if (region.PageSubtotal != null && region.GroupSubtotal != null
                && string.Equals(region.PageSubtotal.MarkerText, region.GroupSubtotal.Footer.MarkerText,
                    StringComparison.Ordinal))
                throw new ArgumentException("分页小计与分组小计不能使用相同标记文本。");
        }

        var normalizedCells = _cells.Select(cell =>
        {
            if (cell.AnchorName != null)
                return cell;
            var merge = _merges.FirstOrDefault(item => string.Equals(item.SheetName, cell.SheetName,
                StringComparison.OrdinalIgnoreCase) && item.Range.Contains(cell.Reference));
            if (merge == null || merge.Range.First.Equals(cell.Reference))
                return cell;
            return new ExcelEntityCellBinding<TEntity>(cell.SheetName, merge.Range.First, cell.PropertyName,
                cell.PropertyType, cell.Getter, cell.Setter, cell.ConverterName,
                cell.MappingConfiguration, cell.MappingDocument);
        }).ToArray();

        var duplicateCell = normalizedCells.Where(cell => cell.AnchorName == null)
            .GroupBy(cell => (cell.SheetName, cell.Reference.Row, cell.Reference.Column),
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

        foreach (var cell in normalizedCells.Where(item => item.AnchorName == null))
        {
            var merge = _merges.FirstOrDefault(item => string.Equals(item.SheetName, cell.SheetName,
                StringComparison.OrdinalIgnoreCase) && item.Range.Contains(cell.Reference));
            foreach (var region in _listRegions.Where(item => item.AnchorName == null
                         && string.Equals(item.SheetName, cell.SheetName,
                         StringComparison.OrdinalIgnoreCase)))
            {
                if (ExcelEntityLayoutValidation.Contains(region.Start, region.End, cell.Reference))
                    throw new ArgumentException($"固定单元格与列表区域重叠: {cell.SheetName}!{cell.Reference.Address}");
            }
        }

        foreach (var pair in _listRegions.Where(item => item.AnchorName == null)
                     .SelectMany((left, index) => _listRegions.Where(item => item.AnchorName == null).Skip(index + 1)
                     .Where(right => string.Equals(left.SheetName, right.SheetName,
                         StringComparison.OrdinalIgnoreCase)
                         && ExcelEntityLayoutValidation.Overlaps(left, right))))
            throw new ArgumentException($"实体布局包含重叠列表区域: {pair.SheetName}!{pair.Start.Address}");

        foreach (var region in _listRegions)
        {
            if (region.AnchorName != null)
                continue;
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
    /// 按声明顺序保存动态列组配置。
    /// </summary>
    private readonly List<ExcelEntityDynamicColumnGroup<TItem>> _dynamicColumnGroups =
        new List<ExcelEntityDynamicColumnGroup<TItem>>();
    /// <summary>
    /// 按声明顺序保存计算列配置。
    /// </summary>
    private readonly List<ExcelEntityCalculatedColumn<TItem>> _calculatedColumns =
        new List<ExcelEntityCalculatedColumn<TItem>>();
    /// <summary>
    /// 保存可选的连续分组小计配置。
    /// </summary>
    private IExcelEntityGroupSubtotal _groupSubtotal;

    /// <summary>
    /// 保存可选的每页明细小计配置。
    /// </summary>
    private ExcelEntityListFooter<TItem> _pageSubtotal;

    /// <summary>
    /// 保存可选的明细分页行数。
    /// </summary>
    private int? _pageBreakRows;

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
    /// 获取显式动态列组快照。
    /// </summary>
    internal IReadOnlyList<ExcelEntityDynamicColumnGroup<TItem>> DynamicColumnGroups => _dynamicColumnGroups;

    /// <summary>
    /// 获取计算列快照。
    /// </summary>
    internal IReadOnlyList<ExcelEntityCalculatedColumn<TItem>> CalculatedColumns => _calculatedColumns;

    /// <summary>
    /// 获取连续分组小计定义。
    /// </summary>
    internal IExcelEntityGroupSubtotal GroupSubtotalDefinition => _groupSubtotal;

    /// <summary>
    /// 获取每页明细小计定义。
    /// </summary>
    internal ExcelEntityListFooter<TItem> PageSubtotalDefinition => _pageSubtotal;

    /// <summary>
    /// 获取明细分页行数。
    /// </summary>
    internal int? PageBreakRows => _pageBreakRows;

    /// <summary>
    /// 获取未知动态值处理策略。
    /// </summary>
    internal ExcelUnknownDynamicValuePolicy UnknownDynamicValuePolicy { get; private set; }

    /// <summary>
    /// 获取可变明细尾部定义。
    /// </summary>
    internal ExcelEntityListFooter<TItem> FooterDefinition { get; private set; }

    /// <summary>
    /// 获取最终尾部的工作簿命名锚点。
    /// </summary>
    internal string FooterAnchorName { get; private set; }

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

    /// <summary>
    /// 添加一个实体属性字典动态列组。
    /// </summary>
    /// <param name="groupKey">动态列组的稳定标识。</param>
    /// <param name="values">实体属性字典表达式。</param>
    /// <param name="definitions">动态列定义。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 字典属性必须是直接属性表达式；动态组与映射配置中的隐式动态列不能同时使用。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> DynamicColumnGroup(string groupKey,
        Expression<Func<TItem, IDictionary<string, object>>> values,
        IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        if (string.IsNullOrWhiteSpace(groupKey))
            throw new ArgumentException("动态列组标识不能为空。", nameof(groupKey));
        if (values == null)
            throw new ArgumentNullException(nameof(values));
        if (definitions == null || definitions.Count == 0)
            throw new ArgumentException("动态列组至少需要一个定义。", nameof(definitions));
        var property = ExcelEntityExpression.GetProperty(values.Body);
        if (!typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
            throw new ArgumentException("动态列字典属性必须实现 IDictionary<string, object>。", nameof(values));
        var getter = values.Compile();
        var snapshots = definitions.Select(CloneDefinition).ToArray();
        _dynamicColumnGroups.Add(new ExcelEntityDynamicColumnGroup<TItem>(groupKey.Trim(), property,
            value => getter((TItem)value), ExcelEntityExpression.CreateObjectSetter<TItem>(property), snapshots));
        return this;
    }

    /// <summary>
    /// 设置未知动态值的处理策略。
    /// </summary>
    /// <param name="policy">未知动态值处理策略。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 未配置显式动态组时，该设置只作为兼容性描述保留，由既有单字典映射处理。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> UnknownDynamicValues(ExcelUnknownDynamicValuePolicy policy)
    {
        if (!Enum.IsDefined(typeof(ExcelUnknownDynamicValuePolicy), policy))
            throw new ArgumentOutOfRangeException(nameof(policy));
        UnknownDynamicValuePolicy = policy;
        return this;
    }

    /// <summary>
    /// 添加一个根据当前行上下文计算值的导出列。
    /// </summary>
    /// <typeparam name="TValue">计算值类型。</typeparam>
    /// <param name="key">计算列的稳定标识。</param>
    /// <param name="title">计算列标题。</param>
    /// <param name="valueFactory">根据行上下文计算单元格值的委托。</param>
    /// <param name="configure">计算列配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 计算列只参与导出；导入时保留其物理列位置但不绑定到实体属性。委托接收本次导出的明细快照，
    /// 每个单元格最多执行一次；取消异常会原样透传，其他异常带有当前 Sheet、行、列和 Key 上下文。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> CalculatedColumn<TValue>(string key, string title,
        Func<ExcelEntityRowContext<TItem>, TValue> valueFactory,
        Action<ExcelEntityCalculatedColumnBuilder> configure = null)
    {
        if (string.IsNullOrWhiteSpace(key))
            throw new ArgumentException("计算列 Key 不能为空。", nameof(key));
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("计算列标题不能为空。", nameof(title));
        if (valueFactory == null)
            throw new ArgumentNullException(nameof(valueFactory));
        var builder = new ExcelEntityCalculatedColumnBuilder();
        configure?.Invoke(builder);
        _calculatedColumns.Add(new ExcelEntityCalculatedColumn<TItem>(key.Trim(), title.Trim(),
            context => valueFactory(context), typeof(TValue), builder));
        return this;
    }

    /// <summary>
    /// 声明可变明细结束后的尾部内容。
    /// </summary>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <param name="configure">尾部配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 尾部相对地址以 marker 所在单元格为 A1；导入时 marker 用于截断明细边界。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> Footer(string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        var builder = new ExcelEntityListFooterBuilder<TItem>();
        configure?.Invoke(builder);
        FooterDefinition = builder.Build(markerText);
        FooterAnchorName = null;
        return this;
    }

    /// <summary>
    /// 声明带工作簿命名锚点的最终尾部。
    /// </summary>
    /// <param name="anchorName">指向最终尾部标记单元格的名称。</param>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <param name="configure">尾部配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 导出时将名称指向实际标记位置；模板中同名、同工作表的单格名称会被更新。
    /// 导入仍由标记文本确定明细边界。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> FooterNamed(string anchorName, string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        ExcelEntityLayoutValidation.ValidateAnchorName(anchorName);
        Footer(markerText, configure);
        FooterAnchorName = anchorName;
        return this;
    }

    /// <summary>
    /// 按相邻明细的分组键写入连续分组小计。
    /// </summary>
    /// <typeparam name="TKey">分组键类型。</typeparam>
    /// <param name="keySelector">读取列表项分组键的表达式。</param>
    /// <param name="markerText">分组小计结束标记文本。</param>
    /// <param name="configure">分组小计尾部配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 分组按明细原有顺序处理，仅当相邻项目的键发生变化时写入小计，不自动排序或合并非连续分组。
    /// 导入时会跳过已识别的分组小计标记行及其间隔和尾部区域。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> GroupSubtotal<TKey>(
        Expression<Func<TItem, TKey>> keySelector, string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        if (keySelector == null)
            throw new ArgumentNullException(nameof(keySelector));
        if (string.IsNullOrWhiteSpace(markerText))
            throw new ArgumentException("分组小计标记不能为空。", nameof(markerText));
        if (_groupSubtotal != null)
            throw new InvalidOperationException("每个列表区域只能配置一个连续分组小计。");
        if (_pageBreakRows.HasValue)
            throw new InvalidOperationException("分页列表不能同时配置连续分组小计。");
        if (_pageSubtotal != null)
            throw new InvalidOperationException("分页小计不能同时配置连续分组小计。");
        var footerBuilder = new ExcelEntityListFooterBuilder<TItem>();
        configure?.Invoke(footerBuilder);
        var footer = footerBuilder.Build(markerText);
        if (footer.ContainsDetailSum)
            throw new ArgumentException("跨小计明细求和公式只能用于最终尾部。", nameof(configure));
        _groupSubtotal = new ExcelEntityGroupSubtotal<TItem, TKey>(keySelector.Compile(), footer);
        return this;
    }

    /// <summary>
    /// 设置列表明细的水平分页行数。
    /// </summary>
    /// <param name="rowsPerPage">每页包含的明细行数，不包含表头。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// Provider 只在导出工作表中写入分页符，不插入空白行，也不改变导入数据边界。
    /// 每个列表区域最多配置一个分页策略；分页不能与连续分组小计组合。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> PageBreak(int rowsPerPage)
    {
        if (rowsPerPage <= 0)
            throw new ArgumentOutOfRangeException(nameof(rowsPerPage), "每页明细行数必须大于零。");
        if (_pageBreakRows.HasValue)
            throw new InvalidOperationException("每个列表区域只能配置一个分页策略。");
        if (_groupSubtotal != null)
            throw new InvalidOperationException("分页列表不能同时配置连续分组小计。");
        _pageBreakRows = rowsPerPage;
        return this;
    }

    /// <summary>
    /// 声明每个完整分页明细后的聚合小计。
    /// </summary>
    /// <param name="markerText">分页小计标记文本。</param>
    /// <param name="configure">分页小计配置委托。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 分页小计仅在存在下一页明细时写入，聚合委托接收当前页的明细快照；必须同时配置
    /// <see cref="PageBreak"/>，且不能与连续分组小计组合。导入时会跳过分页小计区域。
    /// </remarks>
    public ExcelEntityListRegionBuilder<TItem> PageSubtotal(string markerText,
        Action<ExcelEntityListFooterBuilder<TItem>> configure = null)
    {
        if (_pageSubtotal != null)
            throw new InvalidOperationException("每个列表区域只能配置一个分页小计。");
        if (_groupSubtotal != null)
            throw new InvalidOperationException("分页小计不能同时配置连续分组小计。");
        var footerBuilder = new ExcelEntityListFooterBuilder<TItem>();
        configure?.Invoke(footerBuilder);
        var footer = footerBuilder.Build(markerText);
        if (footer.ContainsDetailSum)
            throw new ArgumentException("跨小计明细求和公式只能用于最终尾部。", nameof(configure));
        _pageSubtotal = footer;
        return this;
    }

    /// <summary>
    /// 复制动态列定义及其可变配置。
    /// </summary>
    /// <param name="source">待复制的动态列定义。</param>
    /// <returns>独立的动态列定义。</returns>
    private static ExcelDynamicColumnDefinition CloneDefinition(ExcelDynamicColumnDefinition source)
    {
        if (source == null)
            throw new ArgumentException("动态列定义不能为空。", nameof(source));
        return new ExcelDynamicColumnDefinition
        {
            Key = source.Key,
            Title = source.Title,
            Aliases = source.Aliases == null ? null : source.Aliases.ToArray(),
            DataType = source.DataType,
            Order = source.Order,
            Placement = source.Placement,
            PhysicalColumnIndex = source.PhysicalColumnIndex,
            NumberFormat = source.NumberFormat,
            HeaderStyle = source.HeaderStyle,
            BodyStyle = source.BodyStyle,
            ConverterName = source.ConverterName,
            ValidatorName = source.ValidatorName,
            ValidationRuleNames = source.ValidationRuleNames == null ? null : source.ValidationRuleNames.ToArray(),
            ValidationRules = source.ValidationRules == null ? null : source.ValidationRules
                .Select(CloneValidation).ToArray(),
            ImageMultiplicity = source.ImageMultiplicity
        };
    }

    /// <summary>
    /// 复制动态列校验配置。
    /// </summary>
    /// <param name="source">待复制的校验配置。</param>
    /// <returns>独立的校验配置；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    private static ExcelMappingDynamicValidationConfiguration CloneValidation(
        ExcelMappingDynamicValidationConfiguration source) => source == null ? null : new ExcelMappingDynamicValidationConfiguration
        {
            Name = source.Name,
            Pattern = source.Pattern,
            Format = source.Format,
            CultureName = source.CultureName,
            Min = source.Min,
            Max = source.Max,
            MaxValue = source.MaxValue,
            MaxLength = source.MaxLength,
            IgnoreEmpty = source.IgnoreEmpty
        };
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
    /// 验证模板命名锚点名称。
    /// </summary>
    /// <param name="name">待验证的命名锚点名称。</param>
    internal static void ValidateAnchorName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("模板命名锚点名称不能为空。", nameof(name));
        if (name.Length > 255 || name.IndexOfAny(new[] { '!', '[', ']', '\'', ' ' }) >= 0)
            throw new ArgumentException($"模板命名锚点名称无效: {name}", nameof(name));
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
