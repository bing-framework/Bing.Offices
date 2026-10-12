using System.Globalization;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 校验透视来源、聚合、筛选及展开容量。
/// </summary>
internal static class NpoiPivotSourceValidator
{
    /// <summary>
    /// 创建完整类型化来源计划。
    /// </summary>
    /// <param name="workbook">已写入业务数据的 XSSF 工作簿。</param>
    /// <param name="definition">请求中的不可变透视定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>包含全部来源记录及实际筛选值的计划。</returns>
    internal static NpoiPivotSourcePlan Validate(XSSFWorkbook workbook,
        ExcelPivotTableDefinition definition, CancellationToken token) =>
        Validate(workbook, definition, token, null);

    /// <summary>
    /// 创建完整类型化来源计划并观察成功读取的来源数据行。
    /// </summary>
    /// <param name="workbook">已写入业务数据的 XSSF 工作簿。</param>
    /// <param name="definition">请求中的不可变透视定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="afterSourceRowRead">每个来源数据行读取完成后的观察回调，可为空。</param>
    /// <returns>包含全部来源记录及实际筛选值的计划。</returns>
    internal static NpoiPivotSourcePlan Validate(XSSFWorkbook workbook,
        ExcelPivotTableDefinition definition, CancellationToken token, Action<int> afterSourceRowRead)
    {
        token.ThrowIfCancellationRequested();
        definition.Validate();
        var range = definition.SourceRange;
        if (range.EndRow >= 1048576 || range.EndColumn >= 16384)
            throw Invalid(definition, range.StartRow, range.StartColumn, "来源区域超出 XLSX 行列边界。");
        var sheet = workbook.GetSheet(definition.SourceSheet) as XSSFSheet
            ?? throw Invalid(definition, range.StartRow, range.StartColumn, "来源工作表不存在。");
        var headers = new string[range.EndColumn - range.StartColumn + 1];
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var column = 0; column < headers.Length; column++)
        {
            token.ThrowIfCancellationRequested();
            var value = ReadValue(sheet, definition, range.StartRow, range.StartColumn + column);
            var name = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(name) || !names.Add(name))
                throw Invalid(definition, range.StartRow, range.StartColumn + column, "表头为空或重复。");
            headers[column] = name;
        }
        var columns = headers.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.OrdinalIgnoreCase);
        var advanced = definition.ValueFields.Count > 0;
        var required = advanced
            ? definition.RowFields.Select(field => field.Field).Concat(definition.ColumnFields.Select(field => field.Field))
                .Concat(definition.ValueFields.Select(field => field.Field)).Concat(definition.Filters.Select(field => field.Field))
            : new[] { definition.RowField, definition.ValueField };
        foreach (var field in required)
            if (!columns.ContainsKey(field))
                throw Invalid(definition, range.StartRow, range.StartColumn, $"表头缺少字段 {field}。");
        var numericFields = advanced
            ? definition.ValueFields.Where(value => value.Aggregation != ExcelPivotAggregation.Count)
                .Select(value => value.Field).Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
            : new[] { definition.ValueField };
        var numericSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var filterItems = definition.Filters.ToDictionary(filter => filter.Field,
            _ => new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase);
        var rows = new List<IReadOnlyList<object>>();
        var budget = new NpoiPivotLayoutBudget(definition, columns);
        for (var rowIndex = range.StartRow + 1; rowIndex <= range.EndRow; rowIndex++)
        {
            token.ThrowIfCancellationRequested();
            var row = new object[headers.Length];
            for (var column = 0; column < row.Length; column++)
            {
                token.ThrowIfCancellationRequested();
                var value = ReadValue(sheet, definition, rowIndex, range.StartColumn + column);
                row[column] = value;
            }
            afterSourceRowRead?.Invoke(rowIndex);
            token.ThrowIfCancellationRequested();
            if (row.All(value => value == null))
                throw Invalid(definition, rowIndex, range.StartColumn, "包含空数据行。");
            foreach (var field in numericFields)
            {
                var column = columns[field];
                if (row[column] == null)
                    continue;
                if (row[column] is not double)
                    throw Invalid(definition, rowIndex, range.StartColumn + column, $"数值聚合字段 {field} 必须为数值。");
                numericSeen.Add(field);
            }
            foreach (var filter in definition.Filters)
            {
                var value = row[columns[filter.Field]];
                filterItems[filter.Field].TryAdd(ValueKey(value), value);
            }
            budget.Include(row);
            rows.Add(row);
        }
        foreach (var field in numericFields)
            if (!numericSeen.Contains(field))
                throw Invalid(definition, range.StartRow + 1, range.StartColumn + columns[field], $"数值聚合字段 {field} 没有数值。");
        var resolved = new Dictionary<string, IReadOnlyList<object>>(StringComparer.OrdinalIgnoreCase);
        foreach (var filter in definition.Filters)
        {
            var values = new List<object>();
            foreach (var expected in filter.Values)
            {
                token.ThrowIfCancellationRequested();
                if (!filterItems[filter.Field].TryGetValue(ValueKey(expected), out var actual))
                    throw Invalid(definition, range.StartRow, range.StartColumn + columns[filter.Field],
                        $"筛选字段 {filter.Field} 的指定值不在来源中。");
                values.Add(actual);
            }
            resolved.Add(filter.Field, values.AsReadOnly());
        }
        var outputSize = budget.Validate();
        token.ThrowIfCancellationRequested();
        return new NpoiPivotSourcePlan(definition, sheet, headers, rows, resolved, outputSize.Rows, outputSize.Columns);
    }

    /// <summary>
    /// 读取单元格的实际字面值。
    /// </summary>
    /// <param name="sheet">来源工作表。</param>
    /// <param name="definition">透视定义。</param>
    /// <param name="row">原始零基行号。</param>
    /// <param name="column">原始零基列号。</param>
    /// <returns>文本、有限数值、布尔、日期或空值。</returns>
    private static object ReadValue(XSSFSheet sheet, ExcelPivotTableDefinition definition, int row, int column)
    {
        var cell = sheet.GetRow(row)?.GetCell(column);
        if (cell == null || cell.CellType == CellType.Blank)
            return null;
        switch (cell.CellType)
        {
            case CellType.Formula:
                throw Invalid(definition, row, column, "包含公式，请先计算并写入字面值。");
            case CellType.Error:
                throw Invalid(definition, row, column, "包含 Excel 错误值。");
            case CellType.String: return cell.StringCellValue;
            case CellType.Boolean: return cell.BooleanCellValue;
            case CellType.Numeric:
                var number = cell.NumericCellValue;
                if (!double.IsFinite(number))
                    throw Invalid(definition, row, column, "包含非有限数值。");
                if (!DateUtil.IsCellDateFormatted(cell))
                    return number;
                if (!DateUtil.IsValidExcelDate(number))
                    throw Invalid(definition, row, column, "日期序列无效。");
                try
                {
                    return DateUtil.GetJavaDate(number, sheet.Workbook.IsDate1904());
                }
                catch (ArgumentException exception)
                {
                    throw Invalid(definition, row, column, "日期序列超出可表示范围。", exception);
                }
            default: throw Invalid(definition, row, column, "单元格类型不受支持。");
        }
    }

    /// <summary>
    /// 生成保留类型的轴和筛选比较键。
    /// </summary>
    /// <param name="value">来源或已验证的筛选字面值。</param>
    /// <returns>使用不变文化的类型和值表示。</returns>
    internal static string ValueKey(object value) => value switch
    {
        null => "blank:",
        string text => "text:" + text,
        bool boolean => "boolean:" + (boolean ? "1" : "0"),
        DateTime date => "date:" + date.Ticks.ToString(CultureInfo.InvariantCulture),
        _ => "number:" + NormalizeNumber(value).ToString("R", CultureInfo.InvariantCulture)
    };

    /// <summary>
    /// 将合法筛选数值归一为 Excel 数值。
    /// </summary>
    /// <param name="value">已验证的有限数值。</param>
    /// <returns>统一正负零的 double 数值。</returns>
    private static double NormalizeNumber(object value)
    {
        var number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
        return number == 0 ? 0 : number;
    }

    /// <summary>
    /// 创建保留原始来源坐标的配置异常。
    /// </summary>
    /// <param name="definition">透视定义。</param>
    /// <param name="row">原始零基行号。</param>
    /// <param name="column">原始零基列号。</param>
    /// <param name="reason">错误原因。</param>
    /// <param name="inner">原始处理异常。</param>
    /// <returns>包含一基坐标的配置异常。</returns>
    internal static BingOfficesConfigurationException Invalid(ExcelPivotTableDefinition definition,
        int row, int column, string reason, Exception inner = null) => new(
        $"透视表 {definition.Name} 的来源 {definition.SourceSheet} ({row + 1},{column + 1}) {reason}",
        BingOfficesStage.Preflight, definition.SourceSheet, row + 1, column + 1, inner);
}
