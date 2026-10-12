using System.Globalization;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Exports.PivotTables;

/// <summary>
/// 校验高级透视来源的字段、聚合和筛选值。
/// </summary>
internal static class ClosedXmlAdvancedPivotSourceValidator
{
    /// <summary>
    /// 校验高级透视来源。
    /// </summary>
    /// <param name="range">包含表头的来源区域。</param>
    /// <param name="definition">已快照的透视定义。</param>
    /// <param name="headers">已验证唯一性的原始表头。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>使用来源实际字面值的报表筛选选项。</returns>
    internal static IReadOnlyDictionary<string, IReadOnlyList<XLCellValue>> Validate(
        IXLRange range, ExcelPivotTableDefinition definition,
        IReadOnlyList<string> headers, CancellationToken cancellationToken)
    {
        var columns = headers.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index + 1, StringComparer.OrdinalIgnoreCase);
        var required = definition.RowFields.Select(field => field.Field)
            .Concat(definition.ColumnFields.Select(field => field.Field))
            .Concat(definition.ValueFields.Select(field => field.Field))
            .Concat(definition.Filters.Select(field => field.Field));
        foreach (var field in required)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!columns.ContainsKey(field))
                throw Invalid(definition, range.FirstCell(), $"表头缺少字段 {field}。");
        }

        var numericFields = definition.ValueFields.Where(value => value.Aggregation != ExcelPivotAggregation.Count)
            .Select(value => value.Field).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var seenNumeric = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var filterItems = definition.Filters.ToDictionary(filter => filter.Field,
            _ => new Dictionary<string, XLCellValue>(StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase);
        var budget = new ClosedXmlPivotLayoutBudget(definition, columns);
        foreach (var row in range.Rows().Skip(1))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (row.Cells().All(cell => cell.IsEmpty()))
                throw Invalid(definition, row.FirstCell(), "包含空数据行。");
            foreach (var field in numericFields)
            {
                var cell = row.Cell(columns[field]);
                if (cell.IsEmpty())
                    continue;
                if (!cell.CachedValue.IsNumber)
                    throw Invalid(definition, cell, $"数值聚合字段 {field} 必须为数值。");
                seenNumeric.Add(field);
            }
            foreach (var filter in definition.Filters)
            {
                var item = row.Cell(columns[filter.Field]).CachedValue;
                filterItems[filter.Field].TryAdd(ValueKey(item), item);
            }
            budget.Include(row);
        }

        foreach (var field in numericFields)
            if (!seenNumeric.Contains(field))
                throw Invalid(definition, range.Cell(2, columns[field]), $"数值聚合字段 {field} 没有数值。");
        var resolved = new Dictionary<string, IReadOnlyList<XLCellValue>>(StringComparer.OrdinalIgnoreCase);
        foreach (var filter in definition.Filters)
        {
            var selected = new List<XLCellValue>();
            foreach (var value in filter.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!filterItems[filter.Field].TryGetValue(ValueKey(XLCellValue.FromObject(value,
                        CultureInfo.InvariantCulture)), out var actual))
                    throw Invalid(definition, range.Cell(1, columns[filter.Field]),
                        $"筛选字段 {filter.Field} 的指定值不在来源中。");
                selected.Add(actual);
            }
            resolved.Add(filter.Field, selected.AsReadOnly());
        }
        budget.Validate(range.FirstCell());
        return resolved;
    }

    /// <summary>
    /// 生成保留 Excel 类型的比较键。
    /// </summary>
    /// <param name="value">字面单元格值。</param>
    /// <returns>按不变文化表示的类型和值。</returns>
    internal static string ValueKey(XLCellValue value) =>
        value.Type + ":" + value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// 创建带原始来源坐标的配置异常。
    /// </summary>
    /// <param name="definition">透视定义。</param>
    /// <param name="cell">违规来源单元格。</param>
    /// <param name="reason">错误原因。</param>
    /// <returns>包含透视名称和原始坐标的异常。</returns>
    internal static BingOfficesConfigurationException Invalid(ExcelPivotTableDefinition definition,
        IXLCell cell, string reason) => new(
        $"透视表 {definition.Name} 的来源 {definition.SourceSheet}!{cell.Address} {reason}",
        BingOfficesStage.Preflight, definition.SourceSheet,
        cell.Address.RowNumber, cell.Address.ColumnNumber);
}
