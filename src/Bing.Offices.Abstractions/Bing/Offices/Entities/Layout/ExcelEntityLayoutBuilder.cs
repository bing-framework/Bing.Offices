using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Entities.Layout;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

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
