using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Mappings;
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
    /// <param name="headerMatch">是否要求所有固定属性均出现在表头中。</param>
    /// <param name="maxColumns">允许读取的最大表头列数。</param>
    /// <returns>按源 CSV 列索引排列的绑定列集合。</returns>
    public static IReadOnlyList<CsvColumn> Bind(IEnumerator<IReadOnlyList<string>> records,
        IReadOnlyCollection<CsvPropertyBinding> properties,
        IReadOnlyCollection<CsvPropertyBinding> dynamicProperties,
        IReadOnlyList<IExcelDynamicMappingColumn> dynamicColumns, bool headerMatch, int? maxColumns = null)
    {
        if (records == null)
            throw new ArgumentNullException(nameof(records));
        if (!records.MoveNext())
            throw new CsvInvalidHeaderException("CSV 不包含表头。");
        if (maxColumns.HasValue && records.Current.Count > maxColumns.Value)
            throw new CsvResourceLimitException($"CSV 表头超过最大列数: {maxColumns.Value}");

        var columns = new List<CsvColumn>();
        var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < records.Current.Count; index++)
        {
            var header = records.Current[index];
            if (!headers.Add(header))
                throw new CsvInvalidHeaderException($"CSV 包含重复表头: {header}");
            var property = properties.FirstOrDefault(candidate =>
                string.Equals(candidate.Title, header, StringComparison.OrdinalIgnoreCase)
                || candidate.Aliases.Any(alias => string.Equals(alias, header, StringComparison.OrdinalIgnoreCase))
                || string.Equals(candidate.Name, header, StringComparison.OrdinalIgnoreCase));
            IExcelDynamicMappingColumn dynamicColumn = null;
            if (property == null && dynamicProperties.Count == 1)
            {
                dynamicColumn = dynamicColumns?.FirstOrDefault(candidate =>
                    string.Equals(candidate.Title, header, StringComparison.OrdinalIgnoreCase)
                    || candidate.Aliases.Any(alias => string.Equals(alias, header,
                        StringComparison.OrdinalIgnoreCase)));
                if (dynamicColumns == null || dynamicColumns.Count == 0 || dynamicColumn != null)
                    property = dynamicProperties.First();
            }
            if (property == null)
                continue;
            if (!property.Property.CanWrite)
                throw new CsvInvalidHeaderException($"导入模板属性不可写入: {property.Name}");
            columns.Add(new CsvColumn(index, property, header, property.IsDynamicColumn, dynamicColumn));
        }
        if (headerMatch)
        {
            var missing = properties.Where(property => columns.All(column => column.Property != property))
                .Select(property => property.Title);
            if (missing.Any())
                throw new CsvInvalidHeaderException($"CSV 不存在列: {string.Join(",", missing)}");
        }
        return columns;
    }

    /// <summary>在 CSV 不包含表头时按固定属性声明顺序创建列绑定。</summary>
    /// <param name="properties">按目标 CSV 列顺序排列的固定属性集合。</param>
    /// <returns>按零基列索引排列的绑定列集合。</returns>
    public static IReadOnlyList<CsvColumn> BindByPosition(IReadOnlyList<CsvPropertyBinding> properties)
    {
        if (properties == null)
            throw new ArgumentNullException(nameof(properties));
        return properties.Select((property, index) =>
                new CsvColumn(index, property, property.Title, false)).ToList();
    }
}

/// <summary>表示一个 CSV 源列及其固定或动态实体属性绑定。</summary>
internal sealed class CsvColumn
{
    /// <summary>使用列索引、属性绑定和可选动态列计划创建 CSV 列。</summary>
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
