using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;

namespace Bing.Offices.Csv;

internal static class CsvHeaderBinder
{
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

    public static IReadOnlyList<CsvColumn> BindByPosition(IReadOnlyList<CsvPropertyBinding> properties)
    {
        if (properties == null)
            throw new ArgumentNullException(nameof(properties));
        return properties.Select((property, index) =>
                new CsvColumn(index, property, property.Title, false)).ToList();
    }
}

internal sealed class CsvColumn
{
    public CsvColumn(int index, CsvPropertyBinding property, string headerName, bool isDynamic,
        IExcelDynamicMappingColumn dynamicColumn = null)
    {
        Index = index;
        Property = property;
        HeaderName = headerName;
        IsDynamic = isDynamic;
        DynamicColumn = dynamicColumn;
    }

    public int Index { get; }
    public CsvPropertyBinding Property { get; }
    public string HeaderName { get; }
    public bool IsDynamic { get; }
    public IExcelDynamicMappingColumn DynamicColumn { get; }
}
