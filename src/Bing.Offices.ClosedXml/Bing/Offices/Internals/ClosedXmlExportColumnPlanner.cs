using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Exports;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// ClosedXML 导出列的统一物理布局计划。
/// </summary>
internal static class ClosedXmlExportColumnPlanner
{
    /// <summary>
    /// 生成固定列和动态列共用的物理布局。
    /// </summary>
    /// <param name="itemType">数据项目类型。</param>
    /// <param name="plan">Core 映射计划。</param>
    /// <param name="requestDefinitions">请求中的动态列定义。</param>
    /// <param name="mappingConfiguration">用于读取固定列物理索引的配置。</param>
    /// <returns>按物理列索引排列的导出列。</returns>
    internal static IReadOnlyList<ClosedXmlExportColumn> Create(Type itemType,
        IExcelMappingPlan plan, IReadOnlyList<ExcelDynamicColumnDefinition> requestDefinitions,
        ExcelMappingConfiguration mappingConfiguration = null)
    {
        if (itemType == null)
            throw new ArgumentNullException(nameof(itemType));
        if (plan == null)
            throw new ArgumentNullException(nameof(plan));

        var configuredIndexes = (mappingConfiguration?.Columns ?? new List<ExcelColumnConfiguration>())
            .Where(column => column != null && !string.IsNullOrWhiteSpace(column.PropertyName)
                && column.ColumnIndex.HasValue)
            .ToDictionary(column => column.PropertyName, column => column.ColumnIndex.Value,
                StringComparer.OrdinalIgnoreCase);
        var fixedColumns = plan.Columns
            .Where(column => !column.Ignored && !column.IsDynamicColumn)
            .Select(column =>
            {
                configuredIndexes.TryGetValue(column.Name, out var physicalIndex);
                return new ClosedXmlExportColumn(
                    column.Name, column.Title ?? column.Name, column, null,
                    ResolveProperty(itemType, column.Name), null, null,
                    plan.Style?.NumberFormat, column.Formatter,
                    configuredIndexes.ContainsKey(column.Name) ? physicalIndex : (int?)null);
            })
            .ToList();
        var columns = new List<ClosedXmlExportColumn>();
        var fixedIndexes = new HashSet<int>();

        // 先建立显式固定列锚点，再让未指定索引的固定列按计划顺序填充
        // 最靠前的空槽，避免默认列被错误追加到最高锚点之后。
        foreach (var fixedColumn in fixedColumns)
        {
            if (!configuredIndexes.TryGetValue(fixedColumn.Key, out var configuredIndex))
                continue;
            if (configuredIndex < 0 || !fixedIndexes.Add(configuredIndex))
                throw new ArgumentOutOfRangeException(nameof(mappingConfiguration),
                    $"固定列 {fixedColumn.Key} 的物理索引无效或重复: {configuredIndex}。");
            while (columns.Count <= configuredIndex)
                columns.Add(null);
            if (columns[configuredIndex] != null)
                throw new ArgumentOutOfRangeException(nameof(mappingConfiguration),
                    $"固定列 {fixedColumn.Key} 的物理索引冲突: {configuredIndex}。");
            columns[configuredIndex] = fixedColumn;
        }

        var nextAvailable = 0;
        foreach (var fixedColumn in fixedColumns)
        {
            if (configuredIndexes.ContainsKey(fixedColumn.Key))
                continue;
            while (nextAvailable < columns.Count && columns[nextAvailable] != null)
                nextAvailable++;
            if (nextAvailable == columns.Count)
                columns.Add(fixedColumn);
            else
                columns[nextAvailable] = fixedColumn;
            nextAvailable++;
        }
        var definitions = (requestDefinitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
            .Where(definition => definition != null)
            .ToArray();
        ValidateDefinitions(definitions);

        var dynamicMappings = plan.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>();
        var mappingKeys = new HashSet<string>(dynamicMappings.Select(column => column.Key),
            StringComparer.OrdinalIgnoreCase);
        foreach (var definition in definitions)
        {
            if (!mappingKeys.Contains(definition.Key))
                throw new BingOfficesConfigurationException(
                    $"动态列未在映射计划中定义: {definition.Key}", stage: BingOfficesStage.Plan);
        }

        var dynamicDefinitions = dynamicMappings.Select(mapping =>
        {
            var request = definitions.FirstOrDefault(definition =>
                string.Equals(definition.Key, mapping.Key, StringComparison.OrdinalIgnoreCase));
            var placement = request?.Placement ?? CreatePlacement(mapping.PlacementKey);
            var physicalIndex = request?.PhysicalColumnIndex ?? placement?.PhysicalColumnIndex ?? mapping.ColumnIndex;
            if (request?.PhysicalColumnIndex.HasValue == true && request.Placement != null)
                throw new ArgumentException($"动态列 {mapping.Key} 不能同时指定相对位置和物理索引。",
                    nameof(requestDefinitions));
            return new DynamicPlan(mapping, request?.Title ?? mapping.Title ?? mapping.Key,
                request?.Order ?? mapping.Order, placement, physicalIndex,
                request?.HeaderStyle, request?.BodyStyle,
                plan.Style?.NumberFormat, request?.NumberFormat ?? mapping.NumberFormat);
        }).OrderBy(item => item.Order).ThenBy(item => item.Mapping.Key, StringComparer.Ordinal)
            .ToArray();

        var keys = new HashSet<string>(fixedColumns.Select(column => column.Key),
            StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(fixedColumns.Select(column => column.Title),
            StringComparer.OrdinalIgnoreCase);
        var dynamicColumns = dynamicDefinitions.Select(dynamic => new
        {
            Plan = dynamic,
            Column = new ClosedXmlExportColumn(dynamic.Mapping.Key, dynamic.Title, null,
                dynamic.Mapping, null, dynamic.HeaderStyle, dynamic.BodyStyle,
                dynamic.MappingNumberFormat, dynamic.NumberFormat, null)
        }).ToArray();
        foreach (var dynamic in dynamicColumns)
        {
            if (!keys.Add(dynamic.Plan.Mapping.Key))
                throw new BingOfficesConfigurationException($"导出列键重复: {dynamic.Plan.Mapping.Key}",
                    stage: BingOfficesStage.Plan);
            if (!titles.Add(dynamic.Plan.Title))
                throw new BingOfficesConfigurationException($"导出列标题重复: {dynamic.Plan.Title}",
                    stage: BingOfficesStage.Plan);
        }

        // 相对位置和默认顺序先形成连续布局；显式物理索引随后锚定到最终布局，
        // 这样既保持 Before/After 语义，也能保留用户明确要求的物理空列。
        foreach (var dynamic in dynamicColumns.Where(item => !item.Plan.PhysicalColumnIndex.HasValue))
        {
            var placement = dynamic.Plan.Placement;
            if (placement != null && !string.IsNullOrWhiteSpace(placement.BeforeKey))
            {
                var index = FindKey(columns, placement.BeforeKey);
                if (index < 0)
                    throw new BingOfficesConfigurationException(
                        $"动态列 {dynamic.Plan.Mapping.Key} 的 Before 目标不存在: {placement.BeforeKey}",
                        stage: BingOfficesStage.Plan);
                columns.Insert(index, dynamic.Column);
            }
            else if (placement != null && !string.IsNullOrWhiteSpace(placement.AfterKey))
            {
                var index = FindKey(columns, placement.AfterKey);
                if (index < 0)
                    throw new BingOfficesConfigurationException(
                        $"动态列 {dynamic.Plan.Mapping.Key} 的 After 目标不存在: {placement.AfterKey}",
                        stage: BingOfficesStage.Plan);
                columns.Insert(index + 1, dynamic.Column);
            }
            else
                columns.Add(dynamic.Column);
        }

        var explicitIndexes = new HashSet<int>();
        foreach (var dynamic in dynamicColumns.Where(item => item.Plan.PhysicalColumnIndex.HasValue)
                     .OrderBy(item => item.Plan.PhysicalColumnIndex.Value)
                     .ThenBy(item => item.Plan.Mapping.Key, StringComparer.Ordinal))
        {
            var index = dynamic.Plan.PhysicalColumnIndex.Value;
            if (index < 0 || !explicitIndexes.Add(index))
                throw new ArgumentOutOfRangeException(nameof(requestDefinitions),
                    $"动态列 {dynamic.Plan.Mapping.Key} 的物理索引无效或重复: {index}。");

            while (columns.Count <= index)
                columns.Add(null);
            if (columns[index] != null)
                throw new BingOfficesConfigurationException(
                    $"动态列 {dynamic.Plan.Mapping.Key} 的物理索引与已有列冲突: {index}。",
                    stage: BingOfficesStage.Plan);
            columns[index] = dynamic.Column;
        }

        var result = new List<ClosedXmlExportColumn>();
        for (var index = 0; index < columns.Count; index++)
        {
            var column = columns[index];
            if (column == null)
                continue;
            column.PhysicalColumnIndex = index;
            result.Add(column);
        }
        return result;
    }

    /// <summary>
    /// 检查请求动态列的键、标题和位置定义是否唯一且有效。
    /// </summary>
    /// <param name="definitions">请求中的动态列定义。</param>
    private static void ValidateDefinitions(IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var definition in definitions)
        {
            if (string.IsNullOrWhiteSpace(definition.Key) || string.IsNullOrWhiteSpace(definition.Title))
                throw new ArgumentException("动态列 Key 和 Title 不能为空。", nameof(definitions));
            if (!keys.Add(definition.Key) || !titles.Add(definition.Title))
                throw new ArgumentException($"动态列 Key 或标题重复: {definition.Key}", nameof(definitions));
            if (definition.PhysicalColumnIndex.HasValue && definition.Placement != null)
                throw new ArgumentException($"动态列 {definition.Key} 不能同时指定相对位置和物理索引。",
                    nameof(definitions));
        }
    }

    /// <summary>
    /// 在当前物理布局中查找列键或标题。
    /// </summary>
    /// <param name="columns">当前物理布局。</param>
    /// <param name="key">待查找的列键或标题。</param>
    /// <returns>列索引；未找到时返回 -1。</returns>
    private static int FindKey(IReadOnlyList<ClosedXmlExportColumn> columns, string key)
    {
        for (var index = 0; index < columns.Count; index++)
        {
            if (columns[index] != null
                && (string.Equals(columns[index].Key, key, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(columns[index].Title, key, StringComparison.OrdinalIgnoreCase)))
                return index;
        }
        return -1;
    }

    /// <summary>
    /// 将配置位置键解析为相对列位置描述。
    /// </summary>
    /// <param name="placementKey">位置配置键。</param>
    /// <returns>解析后的位置；输入为空或格式无效时返回 <see langword="null" />。</returns>
    private static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var separator = placementKey.IndexOfAny(new[] { ':', '-' });
        if (separator < 0 || separator == placementKey.Length - 1)
            return null;
        var key = placementKey.Substring(separator + 1);
        return placementKey.StartsWith("before", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key)
            : ExcelColumnPlacement.After(key);
    }

    /// <summary>
    /// 解析实体上的可读属性。
    /// </summary>
    /// <param name="itemType">数据项目类型。</param>
    /// <param name="name">属性名称。</param>
    /// <returns>可读属性信息。</returns>
    private static PropertyInfo ResolveProperty(Type itemType, string name)
    {
        var property = itemType.GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
        if (property == null || !property.CanRead)
            throw new BingOfficesConfigurationException($"属性不存在或不可读: {name}",
                stage: BingOfficesStage.Plan);
        return property;
    }

    /// <summary>
    /// 保存动态列的最终布局信息。
    /// </summary>
    private sealed class DynamicPlan
    {
        /// <summary>
        /// 初始化一个 <see cref="DynamicPlan" /> 类型的实例。
        /// </summary>
        /// <param name="mapping">Core 动态映射列。</param>
        /// <param name="title">导出标题。</param>
        /// <param name="order">默认顺序。</param>
        /// <param name="placement">相对位置。</param>
        /// <param name="physicalColumnIndex">物理列索引。</param>
        /// <param name="headerStyle">请求表头样式。</param>
        /// <param name="bodyStyle">请求正文样式。</param>
        /// <param name="mappingNumberFormat">映射数字格式。</param>
        /// <param name="numberFormat">请求数字格式。</param>
        internal DynamicPlan(IExcelDynamicMappingColumn mapping, string title, int order,
            ExcelColumnPlacement placement, int? physicalColumnIndex, ExcelCellStyle headerStyle,
            ExcelCellStyle bodyStyle, string mappingNumberFormat, string numberFormat)
        {
            Mapping = mapping;
            Title = title;
            Order = order;
            Placement = placement;
            PhysicalColumnIndex = physicalColumnIndex;
            HeaderStyle = headerStyle;
            BodyStyle = bodyStyle;
            MappingNumberFormat = mappingNumberFormat;
            NumberFormat = numberFormat;
        }

        /// <summary>
        /// 获取 Core 动态映射列。
        /// </summary>
        internal IExcelDynamicMappingColumn Mapping { get; }

        /// <summary>
        /// 获取导出标题。
        /// </summary>
        internal string Title { get; }

        /// <summary>
        /// 获取默认顺序。
        /// </summary>
        internal int Order { get; }

        /// <summary>
        /// 获取相对位置。
        /// </summary>
        internal ExcelColumnPlacement Placement { get; }

        /// <summary>
        /// 获取显式物理列索引。
        /// </summary>
        internal int? PhysicalColumnIndex { get; }

        /// <summary>
        /// 获取请求表头样式。
        /// </summary>
        internal ExcelCellStyle HeaderStyle { get; }

        /// <summary>
        /// 获取请求正文样式。
        /// </summary>
        internal ExcelCellStyle BodyStyle { get; }

        /// <summary>
        /// 获取映射数字格式。
        /// </summary>
        internal string MappingNumberFormat { get; }

        /// <summary>
        /// 获取请求数字格式。
        /// </summary>
        internal string NumberFormat { get; }
    }
}

/// <summary>
/// 描述一个 ClosedXML 导出物理列及其映射来源。
/// </summary>
internal sealed class ClosedXmlExportColumn
{
    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExportColumn" /> 类型的实例。
    /// </summary>
    /// <param name="key">列键。</param>
    /// <param name="title">列标题。</param>
    /// <param name="fixedColumn">固定映射列。</param>
    /// <param name="dynamicColumn">动态映射列。</param>
    /// <param name="property">固定列对应的实体属性。</param>
    /// <param name="headerStyle">列级表头样式。</param>
    /// <param name="bodyStyle">列级正文样式。</param>
    /// <param name="mappingNumberFormat">映射数字格式。</param>
    /// <param name="numberFormat">列级数字格式。</param>
    /// <param name="physicalColumnIndex">初始物理列索引。</param>
    internal ClosedXmlExportColumn(string key, string title, IExcelMappingColumn fixedColumn,
        IExcelDynamicMappingColumn dynamicColumn, PropertyInfo property, ExcelCellStyle headerStyle,
        ExcelCellStyle bodyStyle, string mappingNumberFormat, string numberFormat,
        int? physicalColumnIndex)
    {
        Key = key;
        Title = title;
        Fixed = fixedColumn;
        Dynamic = dynamicColumn;
        Property = property;
        HeaderStyle = headerStyle;
        BodyStyle = bodyStyle;
        MappingNumberFormat = mappingNumberFormat;
        NumberFormat = numberFormat;
        PhysicalColumnIndex = physicalColumnIndex ?? 0;
    }

    /// <summary>
    /// 获取列键。
    /// </summary>
    internal string Key { get; }

    /// <summary>
    /// 获取列标题。
    /// </summary>
    internal string Title { get; }

    /// <summary>
    /// 获取固定映射列；动态列时返回 <see langword="null" />。
    /// </summary>
    internal IExcelMappingColumn Fixed { get; }

    /// <summary>
    /// 获取动态映射列；固定列时返回 <see langword="null" />。
    /// </summary>
    internal IExcelDynamicMappingColumn Dynamic { get; }

    /// <summary>
    /// 获取固定列对应的实体属性。
    /// </summary>
    internal PropertyInfo Property { get; }

    /// <summary>
    /// 获取列级表头样式。
    /// </summary>
    internal ExcelCellStyle HeaderStyle { get; }

    /// <summary>
    /// 获取列级正文样式。
    /// </summary>
    internal ExcelCellStyle BodyStyle { get; }

    /// <summary>
    /// 获取映射数字格式。
    /// </summary>
    internal string MappingNumberFormat { get; }

    /// <summary>
    /// 获取列级数字格式。
    /// </summary>
    internal string NumberFormat { get; }

    /// <summary>
    /// 获取或设置零基物理列索引。
    /// </summary>
    internal int PhysicalColumnIndex { get; set; }
}
