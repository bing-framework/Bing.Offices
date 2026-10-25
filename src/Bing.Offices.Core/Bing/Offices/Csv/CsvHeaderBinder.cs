using Bing.Offices.Providers;

namespace Bing.Offices.Csv;

/// <summary>将 CSV 表头或列位置绑定为固定和动态实体属性列。</summary>
internal static class CsvHeaderBinder
{
    /// <summary>读取首条记录并按标题、别名或属性名创建 CSV 列绑定。</summary>
    /// <param name="records">按顺序提供 CSV 记录的枚举器。</param>
    /// <param name="properties">可绑定的固定属性集合。</param>
    /// <param name="dynamicProperties">可绑定的动态字典属性集合。</param>
    /// <param name="dynamicColumns">已编译的动态列表头映射。</param>
    /// <param name="requireExpectedHeaders">是否要求所有固定属性均出现在表头中。</param>
    /// <param name="maxColumns">允许读取的最大表头列数。</param>
    /// <returns>按源 CSV 列索引排列的绑定列集合。</returns>
    public static IReadOnlyList<CsvColumn> Bind(IEnumerator<IReadOnlyList<string>> records,
        IReadOnlyCollection<CsvPropertyBinding> properties,
        IReadOnlyCollection<CsvPropertyBinding> dynamicProperties,
        IReadOnlyList<IExcelDynamicMappingColumn> dynamicColumns, bool requireExpectedHeaders, int? maxColumns = null)
    {
        if (records == null)
            throw new ArgumentNullException(nameof(records));
        if (!records.MoveNext())
            throw new CsvInvalidHeaderException("CSV 不包含表头。");
        return Bind(records.Current, properties, dynamicProperties, dynamicColumns, requireExpectedHeaders, maxColumns);
    }

    /// <summary>使用已读取的首条记录创建列绑定。</summary>
    /// <param name="header">CSV 首条记录中的表头文本。</param>
    /// <param name="properties">可绑定的固定属性集合。</param>
    /// <param name="dynamicProperties">可绑定的动态字典属性集合。</param>
    /// <param name="dynamicColumns">已编译的动态列表头映射。</param>
    /// <param name="requireExpectedHeaders">是否要求所有固定属性均出现在表头中。</param>
    /// <param name="maxColumns">允许读取的最大表头列数。</param>
    /// <returns>按源 CSV 列索引排列的绑定列集合。</returns>
    public static IReadOnlyList<CsvColumn> Bind(IReadOnlyList<string> header,
        IReadOnlyCollection<CsvPropertyBinding> properties,
        IReadOnlyCollection<CsvPropertyBinding> dynamicProperties,
        IReadOnlyList<IExcelDynamicMappingColumn> dynamicColumns, bool requireExpectedHeaders, int? maxColumns = null)
    {
        if (header == null)
            throw new ArgumentNullException(nameof(header));
        if (maxColumns.HasValue && header.Count > maxColumns.Value)
            throw new CsvResourceLimitException($"CSV 表头超过最大列数: {maxColumns.Value}");

        var columns = new List<CsvColumn>();
        var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var boundProperties = new HashSet<string>(StringComparer.Ordinal);
        var boundDynamicKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var propertyLookup = CreatePropertyLookup(properties);
        var dynamicLookup = dynamicProperties.Count == 1 ? CreateDynamicLookup(dynamicColumns) : null;
        for (var index = 0; index < header.Count; index++)
        {
            var headerValue = header[index];
            if (!headers.Add(headerValue))
                throw new CsvInvalidHeaderException($"CSV 包含重复表头: {headerValue}");
            propertyLookup.TryGetValue(headerValue, out var property);
            IExcelDynamicMappingColumn dynamicColumn = null;
            if (property == null && dynamicProperties.Count == 1)
            {
                dynamicLookup.TryGetValue(headerValue, out dynamicColumn);
                if (dynamicColumns == null || dynamicColumns.Count == 0 || dynamicColumn != null)
                    property = dynamicProperties.First();
            }
            if (property == null)
                continue;
            if (!property.Property.CanWrite)
                throw new CsvInvalidHeaderException($"导入模板属性不可写入: {property.Name}");
            // 不同标题和别名也可能指向同一目标，必须在写入实体之前拒绝歧义。
            if (!property.IsDynamicColumn && !boundProperties.Add(property.Name))
                throw new CsvInvalidHeaderException($"CSV 多个表头映射到同一属性: {property.Name}");
            if (dynamicColumn != null && !boundDynamicKeys.Add(dynamicColumn.Key))
                throw new CsvInvalidHeaderException($"CSV 多个表头映射到同一动态列: {dynamicColumn.Key}");
            columns.Add(new CsvColumn(index, property, headerValue, property.IsDynamicColumn, dynamicColumn));
        }
        if (requireExpectedHeaders)
        {
            var missing = properties.Where(property => columns.All(column => column.Property != property))
                .Select(property => property.Title);
            if (missing.Any())
                throw new CsvInvalidHeaderException($"CSV 不存在列: {string.Join(",", missing)}");
        }
        return columns;
    }

    /// <summary>仅为本次绑定建立固定属性索引，冲突词条保留映射顺序中的第一个属性。</summary>
    private static Dictionary<string, CsvPropertyBinding> CreatePropertyLookup(IEnumerable<CsvPropertyBinding> properties)
    {
        var result = new Dictionary<string, CsvPropertyBinding>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            AddFirst(result, property.Title, property);
            foreach (var alias in property.Aliases)
                AddFirst(result, alias, property);
            AddFirst(result, property.Name, property);
        }
        return result;
    }

    /// <summary>仅为本次绑定建立动态标题与别名索引，固定属性仍先于动态定义匹配。</summary>
    private static Dictionary<string, IExcelDynamicMappingColumn> CreateDynamicLookup(
        IReadOnlyList<IExcelDynamicMappingColumn> definitions)
    {
        var result = new Dictionary<string, IExcelDynamicMappingColumn>(StringComparer.OrdinalIgnoreCase);
        foreach (var definition in definitions ?? Array.Empty<IExcelDynamicMappingColumn>())
        {
            AddFirst(result, definition.Title, definition);
            foreach (var alias in definition.Aliases)
                AddFirst(result, alias, definition);
        }
        return result;
    }

    /// <summary>一个词条可以属于多个目标，必须与原 FirstOrDefault 的首次匹配规则一致。</summary>
    private static void AddFirst<T>(IDictionary<string, T> lookup, string token, T value)
    {
        if (token != null && !lookup.ContainsKey(token))
            lookup.Add(token, value);
    }

    /// <summary>在 CSV 不包含表头时，按导出顺序展开固定属性和显式动态列。</summary>
    /// <param name="properties">按映射顺序排列的固定属性与动态字典属性。</param>
    /// <param name="dynamicColumns">动态列定义；按 Order、Key 排序，与导出器一致。</param>
    /// <returns>按零基列索引排列的绑定列集合。</returns>
    public static IReadOnlyList<CsvColumn> BindByPosition(IReadOnlyList<CsvPropertyBinding> properties,
        IReadOnlyList<IExcelDynamicMappingColumn> dynamicColumns)
    {
        if (properties == null)
            throw new ArgumentNullException(nameof(properties));
        var columns = new List<CsvColumn>();
        foreach (var property in properties)
        {
            if (!property.Property.CanWrite)
                throw new CsvInvalidHeaderException($"导入模板属性不可写入: {property.Name}");
            if (!property.IsDynamicColumn)
                columns.Add(new CsvColumn(columns.Count, property, property.Title, false));
            else
            {
                if (dynamicColumns == null || dynamicColumns.Count == 0)
                    throw new CsvInvalidHeaderException("无表头 CSV 导入必须显式配置动态列。");
                foreach (var column in dynamicColumns.OrderBy(column => column.Order)
                             .ThenBy(column => column.Key, StringComparer.Ordinal))
                    columns.Add(new CsvColumn(columns.Count, property, column.Title, true, column));
            }
        }
        return columns;
    }
}

/// <summary>表示一个 CSV 源列及其固定或动态实体属性绑定。</summary>
internal sealed class CsvColumn
{
    /// <summary>初始化一个 <see cref="CsvColumn" /> 类型的实例。</summary>
    /// <param name="index">源 CSV 中的零基列索引。</param>
    /// <param name="property">固定或动态目标属性绑定。</param>
    /// <param name="headerName">源 CSV 表头名称。</param>
    /// <param name="isDynamic">是否将字段写入动态值字典。</param>
    /// <param name="dynamicColumn">与表头匹配的动态列计划。</param>
    public CsvColumn(int index, CsvPropertyBinding property, string headerName, bool isDynamic,
        IExcelDynamicMappingColumn dynamicColumn = null)
    {
        Index = index;
        Property = property;
        HeaderName = headerName;
        IsDynamic = isDynamic;
        DynamicColumn = dynamicColumn;
    }

    /// <summary>获取源 CSV 中的零基列索引。</summary>
    public int Index { get; }
    /// <summary>获取目标实体属性绑定。</summary>
    public CsvPropertyBinding Property { get; }
    /// <summary>获取源 CSV 表头名称。</summary>
    public string HeaderName { get; }
    /// <summary>获取是否将字段写入动态值字典。</summary>
    public bool IsDynamic { get; }
    /// <summary>获取与表头匹配的动态列计划；固定列时为 null。</summary>
    public IExcelDynamicMappingColumn DynamicColumn { get; }
}
