using System.Reflection;
using Bing.Offices.Entities;
using Bing.Offices.Providers;

namespace Bing.Offices.Exports;

/// <summary>
/// 构建和验证 Excel 导出列计划的无状态职责。
/// </summary>
internal static class NpoiExportColumnPlanner
{
    /// <summary>
    /// 检查动态列定义和列键唯一性。
    /// </summary>
    /// <param name="definitions">待检查的动态列定义。</param>
    /// <param name="columns">已生成的固定和动态列计划。</param>
    internal static void ValidateDynamicColumns(IReadOnlyList<ExcelDynamicColumnDefinition> definitions,
        IReadOnlyList<ExcelColumnPlan> columns)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var physicalIndexes = new HashSet<int>();
        foreach (var column in columns.Where(column => !column.IsDynamic))
            titles.Add(column.Title);
        foreach (var definition in definitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
        {
            if (!keys.Add(definition.Key))
                throw new ArgumentException($"动态列包含重复 Key: {definition.Key}", nameof(definitions));
            if (!titles.Add(definition.Title))
                throw new ArgumentException($"动态列包含重复标题: {definition.Title}", nameof(definitions));
            foreach (var alias in definition.Aliases ?? Array.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(alias) || !titles.Add(alias))
                    throw new ArgumentException($"动态列包含重复或空标题别名: {alias}", nameof(definitions));
            }
            if (definition.PhysicalColumnIndex.HasValue && definition.Placement != null)
                throw new ArgumentException($"动态列 {definition.Key} 不能同时指定相对位置和物理索引。",
                    nameof(definitions));
            if (definition.Placement?.PhysicalColumnIndex != null && definition.PhysicalColumnIndex.HasValue)
                throw new ArgumentException($"动态列 {definition.Key} 不能重复指定物理索引。",
                    nameof(definitions));
            var physicalIndex = definition.PhysicalColumnIndex ?? definition.Placement?.PhysicalColumnIndex;
            if (physicalIndex.HasValue && !physicalIndexes.Add(physicalIndex.Value))
                throw new ArgumentException($"动态列包含重复物理列索引: {physicalIndex.Value}",
                    nameof(definitions));
        }
        ValidateColumns(columns);
    }

    /// <summary>
    /// 验证动态列定义具有导出所需的键和标题。
    /// </summary>
    /// <param name="definitions">待验证的动态列定义。</param>
    internal static void ValidateDynamicDefinitions(IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        foreach (var definition in definitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
        {
            if (definition == null || string.IsNullOrWhiteSpace(definition.Key)
                || string.IsNullOrWhiteSpace(definition.Title))
                throw new ArgumentException("动态列 Key 和 Title 不能为空。", nameof(definitions));
        }
    }

    /// <summary>
    /// 创建 Workbook 请求使用的固定列和 typed 动态列计划。
    /// </summary>
    /// <typeparam name="T">工作表数据项类型。</typeparam>
    /// <param name="typeMap">实体类型的映射计划。</param>
    /// <param name="dynamicColumns">请求级动态列定义。</param>
    /// <param name="calculatedColumns">实体布局声明的计算列。</param>
    /// <returns>按最终物理列顺序排列的列计划。</returns>
    internal static IReadOnlyList<ExcelColumnPlan> CreateColumns<T>(IExcelMappingPlan typeMap,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns,
        IReadOnlyList<IExcelEntityCalculatedColumn> calculatedColumns = null)
        where T : class, new()
    {
        var fixedColumns = new List<ExcelColumnPlan>();
        foreach (var property in typeMap.Columns)
        {
            if (property.Ignored)
                continue;
            if (!property.IsDynamicColumn)
            {
                var reflectionProperty = ResolveProperty<T>(property.Name);
                fixedColumns.Add(new ExcelColumnPlan(property.Title, property, false, -1, null, null,
                    property.ValueConverters, property.ValidationBindings,
                    reflectionProperty: reflectionProperty));
                continue;
            }
        }
        var columns = fixedColumns.ToList();
        var dynamicProperty = typeMap.Columns.FirstOrDefault(property => property.IsDynamicColumn);
        if (dynamicProperty == null)
            return AppendCalculatedColumns(columns, calculatedColumns);
        var dynamicReflectionProperty = ResolveProperty<T>(dynamicProperty.Name);
        var definitions = (dynamicColumns ?? Array.Empty<ExcelDynamicColumnDefinition>())
            .OrderBy(definition => definition.Order)
            .ThenBy(definition => definition.Key, StringComparer.Ordinal)
            .ToList();
        foreach (var definition in definitions)
        {
            var dynamicPlan = typeMap.DynamicColumns.Single(column => string.Equals(column.Key, definition.Key,
                StringComparison.OrdinalIgnoreCase));
            var column = new ExcelColumnPlan(definition.Title, dynamicProperty, true, -1, definition, definition.Key,
                dynamicPlan.ValueConverters, dynamicPlan.ValidationBindings,
                reflectionProperty: dynamicReflectionProperty,
                isUnique: dynamicPlan.IsUnique,
                uniqueIgnoreEmpty: dynamicPlan.UniqueIgnoreEmpty);
            var placement = definition.Placement;
            var physicalIndex = definition.PhysicalColumnIndex ?? placement?.PhysicalColumnIndex;
            if (physicalIndex.HasValue)
            {
                if (physicalIndex.Value > columns.Count)
                    throw new ArgumentOutOfRangeException(nameof(definition.PhysicalColumnIndex),
                        $"动态列 {definition.Key} 的物理索引超出当前列计划。");
                columns.Insert(physicalIndex.Value, column);
                continue;
            }
            if (placement != null && !string.IsNullOrWhiteSpace(placement.BeforeKey))
            {
                var index = columns.FindIndex(item => string.Equals(item.Key, placement.BeforeKey,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                    throw new ArgumentException($"动态列 {definition.Key} 的 Before 目标不存在: {placement.BeforeKey}");
                columns.Insert(index, column);
                continue;
            }
            if (placement != null && !string.IsNullOrWhiteSpace(placement.AfterKey))
            {
                var index = columns.FindIndex(item => string.Equals(item.Key, placement.AfterKey,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                    throw new ArgumentException($"动态列 {definition.Key} 的 After 目标不存在: {placement.AfterKey}");
                columns.Insert(index + 1, column);
                continue;
            }
            columns.Add(column);
        }
        return AppendCalculatedColumns(columns, calculatedColumns);
    }

    /// <summary>
    /// 创建实体列表区域的列计划。
    /// </summary>
    /// <typeparam name="T">列表项类型。</typeparam>
    /// <param name="typeMap">列表项映射计划。</param>
    /// <param name="dynamicColumns">按组展开的动态列定义。</param>
    /// <param name="groups">实体布局声明的动态列组。</param>
    /// <param name="calculatedColumns">实体布局声明的计算列。</param>
    /// <returns>按最终物理列顺序排列的列计划。</returns>
    internal static IReadOnlyList<ExcelColumnPlan> CreateEntityColumns<T>(IExcelMappingPlan typeMap,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups,
        IReadOnlyList<IExcelEntityCalculatedColumn> calculatedColumns = null)
        where T : class, new()
    {
        if (typeMap == null)
            throw new ArgumentNullException(nameof(typeMap));
        groups ??= Array.Empty<ExcelEntityDynamicColumnGroup<T>>();
        if (groups.Count == 0)
            return CreateColumns<T>(typeMap, dynamicColumns, calculatedColumns);

        var dynamicPropertyNames = new HashSet<string>(groups.Select(group => group.Property.Name),
            StringComparer.OrdinalIgnoreCase);
        var columns = new List<ExcelColumnPlan>();
        foreach (var property in typeMap.Columns)
        {
            if (property.Ignored || dynamicPropertyNames.Contains(property.Name) || property.IsDynamicColumn)
                continue;
            columns.Add(new ExcelColumnPlan(property.Title, property, false, -1, null, null,
                property.ValueConverters, property.ValidationBindings,
                reflectionProperty: ResolveProperty<T>(property.Name)));
        }

        var definitions = (dynamicColumns ?? Array.Empty<ExcelDynamicColumnDefinition>())
            .Select((definition, index) => new { Definition = definition, Index = index })
            .OrderBy(item => item.Definition.Order)
            .ThenBy(item => item.Index)
            .Select(item => item.Definition)
            .ToList();
        foreach (var definition in definitions)
        {
            var group = groups.FirstOrDefault(item => item.Definitions.Any(column =>
                string.Equals(column.Key, definition.Key, StringComparison.OrdinalIgnoreCase)));
            if (group == null)
                throw new ArgumentException($"动态列 {definition.Key} 未关联到实体动态列组。",
                    nameof(dynamicColumns));
            var property = typeMap.Columns.FirstOrDefault(item =>
                string.Equals(item.Name, group.Property.Name, StringComparison.OrdinalIgnoreCase));
            if (property == null)
                throw new ArgumentException($"动态列组属性未生成映射: {group.Property.Name}", nameof(groups));
            var dynamicPlan = typeMap.DynamicColumns.SingleOrDefault(column =>
                string.Equals(column.Key, definition.Key, StringComparison.OrdinalIgnoreCase));
            if (dynamicPlan == null)
                throw new ArgumentException($"动态列未生成映射计划: {definition.Key}", nameof(dynamicColumns));
            var column = new ExcelColumnPlan(definition.Title, property, true, -1, definition,
                definition.Key, dynamicPlan.ValueConverters, dynamicPlan.ValidationBindings,
                reflectionProperty: group.Property, isUnique: dynamicPlan.IsUnique,
                uniqueIgnoreEmpty: dynamicPlan.UniqueIgnoreEmpty);
            InsertDynamicColumn(columns, column, definition);
        }
        ValidateDynamicColumns(dynamicColumns, columns);
        return AppendCalculatedColumns(columns, calculatedColumns);
    }

    /// <summary>
    /// 将实体布局计算列加入已有物理列计划。
    /// </summary>
    /// <param name="columns">固定列和动态列计划。</param>
    /// <param name="calculatedColumns">计算列定义。</param>
    /// <returns>包含计算列的物理列计划。</returns>
    private static IReadOnlyList<ExcelColumnPlan> AppendCalculatedColumns(
        IList<ExcelColumnPlan> columns,
        IReadOnlyList<IExcelEntityCalculatedColumn> calculatedColumns)
    {
        var definitions = (calculatedColumns ?? Array.Empty<IExcelEntityCalculatedColumn>())
            .Where(column => column != null)
            .OrderBy(column => column.Order)
            .ThenBy(column => column.Key, StringComparer.Ordinal)
            .ToArray();
        if (definitions.Length == 0)
        {
            ValidateColumns(columns.ToArray());
            return columns.ToArray();
        }

        var keys = new HashSet<string>(columns.Select(column => column.Key),
            StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(columns.Select(column => column.Title),
            StringComparer.OrdinalIgnoreCase);
        foreach (var definition in definitions)
        {
            if (string.IsNullOrWhiteSpace(definition.Key) || string.IsNullOrWhiteSpace(definition.Title))
                throw new ArgumentException("计算列 Key 和 Title 不能为空。", nameof(calculatedColumns));
            if (!keys.Add(definition.Key) || !titles.Add(definition.Title))
                throw new ArgumentException($"计算列 Key 或标题重复: {definition.Key}",
                    nameof(calculatedColumns));
            var column = new ExcelColumnPlan(definition);
            InsertCalculatedColumn(columns, column, definition);
        }
        ValidateColumns(columns.ToArray());
        return columns.ToArray();
    }

    /// <summary>
    /// 按计算列位置声明插入列计划。
    /// </summary>
    /// <param name="columns">当前物理列计划。</param>
    /// <param name="column">待插入的计算列。</param>
    /// <param name="definition">计算列定义。</param>
    private static void InsertCalculatedColumn(IList<ExcelColumnPlan> columns, ExcelColumnPlan column,
        IExcelEntityCalculatedColumn definition)
    {
        var placement = definition.Placement;
        var physicalIndex = definition.PhysicalColumnIndex ?? placement?.PhysicalColumnIndex;
        if (physicalIndex.HasValue)
        {
            if (physicalIndex.Value < 0 || physicalIndex.Value > columns.Count)
                throw new ArgumentOutOfRangeException(nameof(definition),
                    $"计算列 {definition.Key} 的物理索引超出当前列计划。");
            columns.Insert(physicalIndex.Value, column);
            return;
        }
        if (placement != null && !string.IsNullOrWhiteSpace(placement.BeforeKey))
        {
            var index = columns.ToList().FindIndex(item => string.Equals(item.Key, placement.BeforeKey,
                StringComparison.OrdinalIgnoreCase) || string.Equals(item.Title, placement.BeforeKey,
                StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new ArgumentException($"计算列 {definition.Key} 的 Before 目标不存在: {placement.BeforeKey}");
            columns.Insert(index, column);
            return;
        }
        if (placement != null && !string.IsNullOrWhiteSpace(placement.AfterKey))
        {
            var index = columns.ToList().FindIndex(item => string.Equals(item.Key, placement.AfterKey,
                StringComparison.OrdinalIgnoreCase) || string.Equals(item.Title, placement.AfterKey,
                StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new ArgumentException($"计算列 {definition.Key} 的 After 目标不存在: {placement.AfterKey}");
            columns.Insert(index + 1, column);
            return;
        }
        columns.Add(column);
    }

    /// <summary>
    /// 按动态列的物理索引或相对位置插入列计划。
    /// </summary>
    /// <param name="columns">待修改的列计划集合。</param>
    /// <param name="column">待插入的动态列计划。</param>
    /// <param name="definition">动态列定义。</param>
    private static void InsertDynamicColumn(IList<ExcelColumnPlan> columns, ExcelColumnPlan column,
        ExcelDynamicColumnDefinition definition)
    {
        var placement = definition.Placement;
        var physicalIndex = definition.PhysicalColumnIndex ?? placement?.PhysicalColumnIndex;
        if (physicalIndex.HasValue)
        {
            if (physicalIndex.Value > columns.Count)
                throw new ArgumentOutOfRangeException(nameof(definition.PhysicalColumnIndex),
                    $"动态列 {definition.Key} 的物理索引超出当前列计划。");
            columns.Insert(physicalIndex.Value, column);
            return;
        }
        if (placement != null && !string.IsNullOrWhiteSpace(placement.BeforeKey))
        {
            var index = columns.ToList().FindIndex(item => string.Equals(item.Key, placement.BeforeKey,
                StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new ArgumentException($"动态列 {definition.Key} 的 Before 目标不存在: {placement.BeforeKey}");
            columns.Insert(index, column);
            return;
        }
        if (placement != null && !string.IsNullOrWhiteSpace(placement.AfterKey))
        {
            var index = columns.ToList().FindIndex(item => string.Equals(item.Key, placement.AfterKey,
                StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new ArgumentException($"动态列 {definition.Key} 的 After 目标不存在: {placement.AfterKey}");
            columns.Insert(index + 1, column);
            return;
        }
        columns.Add(column);
    }

    /// <summary>
    /// 合并映射计划和请求级样式以生成导出动态列定义。
    /// </summary>
    /// <param name="column">已绑定的动态列映射。</param>
    /// <param name="requestColumn">请求中与映射键匹配的动态列定义。</param>
    /// <param name="mapping">包含默认样式和布局的列映射计划。</param>
    /// <returns>用于生成导出列计划的动态列定义。</returns>
    internal static ExcelDynamicColumnDefinition CreateDynamicDefinition(IExcelDynamicMappingColumn column,
        ExcelDynamicColumnDefinition requestColumn, IExcelMappingPlan mapping)
    {
        var columnIndex = column.ColumnIndex;
        var placementKey = column.PlacementKey;
        if (!columnIndex.HasValue && string.IsNullOrWhiteSpace(placementKey)
            && mapping.DynamicColumns.Count == 1)
        {
            columnIndex = mapping.Layout?.ColumnIndex;
            placementKey = mapping.Layout?.PlacementKey;
        }
        return new ExcelDynamicColumnDefinition
        {
            Key = column.Key,
            Title = column.Title,
            Aliases = column.Aliases,
            DataType = ResolveDynamicType(column.DataTypeName),
            Order = column.Order,
            Placement = CreatePlacement(placementKey),
            PhysicalColumnIndex = columnIndex,
            NumberFormat = column.NumberFormat ?? mapping.Style?.NumberFormat,
            HeaderStyle = requestColumn?.HeaderStyle,
            BodyStyle = requestColumn?.BodyStyle,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = column.ValidationRuleNames,
            ImageMultiplicity = column.ImageMultiplicity
        };
    }

    /// <summary>
    /// 将 before/after 形式的位置键转换为动态列定位规则。
    /// </summary>
    /// <param name="placementKey">映射中保存的相对位置键。</param>
    /// <returns>解析后的定位规则；键为空时为 null。</returns>
    internal static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var key = placementKey.Substring(placementKey.IndexOfAny(new[] { ':', '-' }) + 1);
        return placementKey.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || placementKey.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key)
            : ExcelColumnPlacement.After(key);
    }

    /// <summary>
    /// 解析配置允许的动态列逻辑类型名称。
    /// </summary>
    /// <param name="name">配置中的类型名称；为空时使用 string。</param>
    /// <returns>对应的 CLR 类型。</returns>
    internal static Type ResolveDynamicType(string name)
    {
        switch ((name ?? "string").ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "string": return typeof(string);
            case "boolean": case "bool": return typeof(bool);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int32": case "int": return typeof(int);
            case "int64": case "long": return typeof(long);
            case "single": case "float": return typeof(float);
            case "double": return typeof(double);
            case "decimal": return typeof(decimal);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            default: throw new InvalidOperationException($"动态列数据类型不在允许列表中: {name}");
        }
    }

    /// <summary>
    /// 验证解析后的列标题唯一，确保导出结果可被导入器无歧义读取。
    /// </summary>
    /// <param name="columns">当前请求的导出列。</param>
    internal static void ValidateColumns(IReadOnlyList<ExcelColumnPlan> columns)
    {
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (!titles.Add(column.Title))
                throw new ArgumentException($"导出列标题重复: {column.Title}", nameof(columns));
        }
    }

    /// <summary>
    /// 解析实体上的公共实例属性。
    /// </summary>
    /// <typeparam name="T">导出实体类型。</typeparam>
    /// <param name="name">映射配置中的属性名称。</param>
    /// <returns>匹配的公共实例属性。</returns>
    internal static PropertyInfo ResolveProperty<T>(string name) where T : class, new()
    {
        var property = typeof(T).GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
        if (property == null)
            throw new InvalidOperationException($"无法解析映射属性: {name}");
        return property;
    }
}
