using Bing.Offices.Exports;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 计算透视完整展开的保守空间预算。
/// </summary>
internal sealed class NpoiPivotLayoutBudget
{
    /// <summary>
    /// 当前不可变透视定义。
    /// </summary>
    private readonly ExcelPivotTableDefinition _definition;

    /// <summary>
    /// 来源字段到区域内零基列号的映射。
    /// </summary>
    private readonly IReadOnlyDictionary<string, int> _columns;

    /// <summary>
    /// 行轴各层级的类型化前缀集合。
    /// </summary>
    private readonly HashSet<string>[] _rows;

    /// <summary>
    /// 列轴各层级的类型化前缀集合。
    /// </summary>
    private readonly HashSet<string>[] _columnsPrefixes;

    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotLayoutBudget" /> 类型的实例。
    /// </summary>
    /// <param name="definition">不可变透视定义。</param>
    /// <param name="columns">来源字段到区域内零基列号的映射。</param>
    internal NpoiPivotLayoutBudget(ExcelPivotTableDefinition definition, IReadOnlyDictionary<string, int> columns)
    {
        _definition = definition;
        _columns = columns;
        _rows = CreatePrefixes(definition.ValueFields.Count == 0 ? 1 : definition.RowFields.Count);
        _columnsPrefixes = CreatePrefixes(definition.ColumnFields.Count);
    }

    /// <summary>
    /// 纳入一条来源记录的轴层级。
    /// </summary>
    /// <param name="row">包含全部来源字段的字面值记录。</param>
    internal void Include(IReadOnlyList<object> row)
    {
        if (_definition.ValueFields.Count == 0)
            _rows[0].Add(NpoiPivotSourceValidator.ValueKey(row[_columns[_definition.RowField]]));
        else
            IncludeAxis(row, _definition.RowFields, _rows);
        IncludeAxis(row, _definition.ColumnFields, _columnsPrefixes);
    }

    /// <summary>
    /// 校验保守展开面积和 XLSX 行列边界。
    /// </summary>
    /// <returns>包含筛选区的完整展开行列上界。</returns>
    internal (int Rows, int Columns) Validate()
    {
        try
        {
            long height;
            long width;
            if (_definition.ValueFields.Count == 0)
            {
                height = _rows[0].Count + 2L;
                width = 2;
            }
            else
            {
                var rowItems = CountAxis(_definition.RowFields, _rows);
                var columnItems = CountAxis(_definition.ColumnFields, _columnsPrefixes);
                var filterHeight = _definition.Filters.Count == 0 ? 0L : _definition.Filters.Count + 1L;
                var headerHeight = Math.Max(1L, _definition.ColumnFields.Count)
                    + (_definition.ValueFields.Count > 1 ? 1L : 0L);
                height = checked(filterHeight + headerHeight + rowItems + 1);
                width = checked(Math.Max(1L, _definition.RowFields.Count)
                    + checked(columnItems + (_definition.ColumnFields.Count > 0 ? 1L : 0L))
                    * _definition.ValueFields.Count);
                if (_definition.Filters.Count > 0)
                    width = Math.Max(2, width);
            }
            if ((long)_definition.TargetRow + height > 1048576
                || (long)_definition.TargetColumn + width > 16384)
                throw NpoiPivotSourceValidator.Invalid(_definition, _definition.SourceRange.StartRow,
                    _definition.SourceRange.StartColumn, "完整展开上界超出 XLSX 行列边界。");
            var area = checked(height * width);
            if (_definition.MaxExpandedCells.HasValue && area > _definition.MaxExpandedCells.Value)
                throw NpoiPivotSourceValidator.Invalid(_definition, _definition.SourceRange.StartRow,
                    _definition.SourceRange.StartColumn,
                    $"完整展开上界 {area} 超过 MaxExpandedCells {_definition.MaxExpandedCells.Value}。");
            return (checked((int)height), checked((int)width));
        }
        catch (OverflowException exception)
        {
            throw NpoiPivotSourceValidator.Invalid(_definition, _definition.SourceRange.StartRow,
                _definition.SourceRange.StartColumn, "完整展开上界计算溢出。", exception);
        }
    }

    /// <summary>
    /// 纳入来源记录的类型化轴前缀。
    /// </summary>
    /// <param name="row">完整来源记录。</param>
    /// <param name="fields">有序轴字段。</param>
    /// <param name="prefixes">各层级独立的前缀集合。</param>
    private void IncludeAxis(IReadOnlyList<object> row, IReadOnlyList<ExcelPivotAxisFieldDefinition> fields,
        IReadOnlyList<HashSet<string>> prefixes)
    {
        var prefix = string.Empty;
        for (var index = 0; index < fields.Count; index++)
        {
            var key = NpoiPivotSourceValidator.ValueKey(row[_columns[fields[index].Field]]);
            prefix += key.Length + ":" + key;
            prefixes[index].Add(prefix);
        }
    }

    /// <summary>
    /// 计算轴叶项和非叶层级自动小计数量。
    /// </summary>
    /// <param name="fields">有序轴字段。</param>
    /// <param name="prefixes">各层级已出现的组合。</param>
    /// <returns>完整轴项目的保守数量。</returns>
    private static long CountAxis(IReadOnlyList<ExcelPivotAxisFieldDefinition> fields,
        IReadOnlyList<HashSet<string>> prefixes)
    {
        if (fields.Count == 0)
            return 1;
        long count = prefixes[fields.Count - 1].Count;
        for (var index = 0; index < fields.Count - 1; index++)
            if (fields[index].SubtotalMode == ExcelPivotSubtotalMode.Automatic)
                count = checked(count + prefixes[index].Count);
        return count;
    }

    /// <summary>
    /// 创建忽略文本大小写的独立层级集合。
    /// </summary>
    /// <param name="count">轴层级数量。</param>
    /// <returns>各层级独立的前缀集合。</returns>
    private static HashSet<string>[] CreatePrefixes(int count) => Enumerable.Range(0, count)
        .Select(_ => new HashSet<string>(StringComparer.OrdinalIgnoreCase)).ToArray();
}
