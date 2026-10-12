using System.Globalization;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Exports.PivotTables;

/// <summary>
/// 校验透视表来源数据。
/// </summary>
internal static class ClosedXmlPivotSourceValidator
{
    /// <summary>
    /// XLSX 工作表允许的最大行数。
    /// </summary>
    private const int MaximumRowCount = 1048576;

    /// <summary>
    /// 校验透视来源区域。
    /// </summary>
    /// <param name="range">包含表头的来源区域。</param>
    /// <param name="definition">透视表定义。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按来源实际类型和值解析的报表筛选选项。</returns>
    internal static IReadOnlyDictionary<string, IReadOnlyList<XLCellValue>> Validate(
        IXLRange range, ExcelPivotTableDefinition definition,
        CancellationToken cancellationToken)
    {
        // 在任何 Value/GetString 读取及缓存创建前检查全部来源字段，避免隐式公式求值。
        foreach (var cell in range.CellsUsed(XLCellsUsedOptions.Contents))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.HasFormula)
                throw InvalidSource(definition, cell, "包含公式，请先计算并写入字面值。");
            if (cell.CachedValue.Type == XLDataType.Error)
                throw InvalidSource(definition, cell, "包含 Excel 错误值。");
        }

        var headerCells = range.FirstRow().Cells().ToArray();
        var headers = new string[headerCells.Length];
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < headerCells.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var header = headerCells[index].GetString();
            if (string.IsNullOrWhiteSpace(header) || !names.Add(header))
                throw InvalidSource(definition, headerCells[index], "表头为空或重复。");
            headers[index] = header;
        }

        if (definition.ValueFields.Count > 0)
        {
            return ClosedXmlAdvancedPivotSourceValidator.Validate(range, definition, headers, cancellationToken);
        }

        if (!names.Contains(definition.RowField) || !names.Contains(definition.ValueField))
            throw InvalidSource(definition, headerCells[0],
                $"表头缺少字段 {definition.RowField} 或 {definition.ValueField}。");

        var valueColumn = Array.FindIndex(headers, header => string.Equals(header,
            definition.ValueField, StringComparison.OrdinalIgnoreCase)) + 1;
        var rowColumn = Array.FindIndex(headers, header => string.Equals(header,
            definition.RowField, StringComparison.OrdinalIgnoreCase)) + 1;
        var groups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var hasNumericValue = false;
        foreach (var row in range.Rows().Skip(1))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (row.Cells().All(cell => cell.IsEmpty()))
                throw InvalidSource(definition, row.FirstCell(), "包含空数据行。");

            var group = row.Cell(rowColumn).Value;
            groups.Add(group.Type + ":" + group.ToString(CultureInfo.InvariantCulture));
            if ((long)definition.TargetRow + groups.Count + 2 > MaximumRowCount)
                throw InvalidSource(definition, row.Cell(rowColumn), "展开区域超出 XLSX 行边界。");

            var value = row.Cell(valueColumn);
            if (value.IsEmpty())
                continue;
            if (!value.Value.IsNumber)
                throw InvalidSource(definition, value, "求和字段必须为数值。");
            hasNumericValue = true;
        }

        if (!hasNumericValue)
            throw InvalidSource(definition, range.Cell(2, valueColumn), "求和字段没有数值。");
        var expandedCells = (groups.Count + 2L) * 2;
        if (definition.MaxExpandedCells.HasValue && expandedCells > definition.MaxExpandedCells.Value)
            throw InvalidSource(definition, range.FirstCell(),
                $"完整展开上界 {expandedCells} 超过 MaxExpandedCells {definition.MaxExpandedCells.Value}。");
        return new Dictionary<string, IReadOnlyList<XLCellValue>>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 创建带来源坐标的配置异常。
    /// </summary>
    /// <param name="definition">透视表定义。</param>
    /// <param name="cell">发生错误的来源单元格。</param>
    /// <param name="reason">错误原因。</param>
    /// <returns>包含透视表名称和原始坐标的配置异常。</returns>
    private static BingOfficesConfigurationException InvalidSource(
        ExcelPivotTableDefinition definition, IXLCell cell, string reason) => new(
        $"透视表 {definition.Name} 的来源 {definition.SourceSheet}!{cell.Address} {reason}",
        BingOfficesStage.Preflight, definition.SourceSheet,
        cell.Address.RowNumber, cell.Address.ColumnNumber);
}
