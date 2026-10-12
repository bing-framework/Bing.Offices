using Bing.Offices.Exports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Exports.PivotTables;

/// <summary>
/// 计算高级透视完整展开的保守空间预算。
/// </summary>
internal sealed class ClosedXmlPivotLayoutBudget
{
    /// <summary>
    /// 当前不可变透视定义。
    /// </summary>
    private readonly ExcelPivotTableDefinition _definition;

    /// <summary>
    /// 来源表头到区域内一基列号的映射。
    /// </summary>
    private readonly IReadOnlyDictionary<string, int> _columns;

    /// <summary>
    /// 各行轴层级的已出现前缀。
    /// </summary>
    private readonly HashSet<string>[] _rowPrefixes;

    /// <summary>
    /// 各列轴层级的已出现前缀。
    /// </summary>
    private readonly HashSet<string>[] _columnPrefixes;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlPivotLayoutBudget"/> 类型的实例。
    /// </summary>
    /// <param name="definition">已快照的透视定义。</param>
    /// <param name="columns">来源字段的一基列号。</param>
    internal ClosedXmlPivotLayoutBudget(ExcelPivotTableDefinition definition,
        IReadOnlyDictionary<string, int> columns)
    {
        _definition = definition;
        _columns = columns;
        _rowPrefixes = CreatePrefixes(definition.RowFields.Count);
        _columnPrefixes = CreatePrefixes(definition.ColumnFields.Count);
    }

    /// <summary>
    /// 纳入一条来源记录的轴层级。
    /// </summary>
    /// <param name="row">来源数据行。</param>
    internal void Include(IXLRangeRow row)
    {
        IncludeAxis(row, _definition.RowFields, _rowPrefixes);
        IncludeAxis(row, _definition.ColumnFields, _columnPrefixes);
    }

    /// <summary>
    /// 校验保守展开面积与格式边界。
    /// </summary>
    /// <param name="source">用于定位预算错误的来源单元格。</param>
    internal void Validate(IXLCell source)
    {
        // 筛选不减少完整缓存的展开上界；多度量统一放在列轴。
        var rowItems = CountAxis(_definition.RowFields, _rowPrefixes);
        var columnItems = CountAxis(_definition.ColumnFields, _columnPrefixes);
        var filterHeight = _definition.Filters.Count == 0 ? 0L : _definition.Filters.Count + 1L;
        var headerHeight = Math.Max(1L, _definition.ColumnFields.Count)
            + (_definition.ValueFields.Count > 1 ? 1L : 0L);
        var height = checked(filterHeight + headerHeight + rowItems + 1);
        var width = checked(Math.Max(1L, _definition.RowFields.Count)
            + (columnItems + (_definition.ColumnFields.Count > 0 ? 1L : 0L))
            * _definition.ValueFields.Count);
        if (_definition.Filters.Count > 0)
            width = Math.Max(2, width);
        if ((long)_definition.TargetRow + height > 1048576
            || (long)_definition.TargetColumn + width > 16384)
            throw ClosedXmlAdvancedPivotSourceValidator.Invalid(_definition, source,
                "完整展开上界超出 XLSX 行列边界。");
        var area = checked(height * width);
        if (_definition.MaxExpandedCells.HasValue && area > _definition.MaxExpandedCells.Value)
            throw ClosedXmlAdvancedPivotSourceValidator.Invalid(_definition, source,
                $"完整展开上界 {area} 超过 MaxExpandedCells {_definition.MaxExpandedCells.Value}。");
    }

    /// <summary>
    /// 记录一条记录的轴前缀。
    /// </summary>
    /// <param name="row">来源记录。</param>
    /// <param name="fields">有序轴字段。</param>
    /// <param name="prefixes">对应层级的前缀集合。</param>
    private void IncludeAxis(IXLRangeRow row, IReadOnlyList<ExcelPivotAxisFieldDefinition> fields,
        IReadOnlyList<HashSet<string>> prefixes)
    {
        var prefix = string.Empty;
        for (var index = 0; index < fields.Count; index++)
        {
            var key = ClosedXmlAdvancedPivotSourceValidator.ValueKey(row.Cell(_columns[fields[index].Field]).CachedValue);
            // 长度前缀避免业务文本包含分隔符时合并不同轴组合。
            prefix += key.Length + ":" + key;
            prefixes[index].Add(prefix);
        }
    }

    /// <summary>
    /// 计算轴叶项及非叶层级自动小计的数量。
    /// </summary>
    /// <param name="fields">有序轴字段。</param>
    /// <param name="prefixes">已出现的各层级组合。</param>
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
    /// 创建各层级独立的前缀集合。
    /// </summary>
    /// <param name="count">轴层级数。</param>
    /// <returns>忽略文本大小写的独立前缀集合。</returns>
    private static HashSet<string>[] CreatePrefixes(int count) => Enumerable.Range(0, count)
        .Select(_ => new HashSet<string>(StringComparer.OrdinalIgnoreCase)).ToArray();
}
