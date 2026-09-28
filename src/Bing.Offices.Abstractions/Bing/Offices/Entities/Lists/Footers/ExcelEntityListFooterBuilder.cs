using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Bing.Offices.Entities.Layout;
using Bing.Offices.Exports;
using Bing.Offices.Styles;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体列表区域尾部构建器。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
public sealed class ExcelEntityListFooterBuilder<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityListFooterBuilder{TItem}" /> 类型的实例。
    /// </summary>
    public ExcelEntityListFooterBuilder()
    {
    }

    /// <summary>
    /// 保存尾部单元格定义。
    /// </summary>
    private readonly List<(ExcelEntityCellReference Reference,
        Func<int, IReadOnlyList<TItem>, object> ValueFactory, ExcelCellStyle Style,
        string NumberFormat)> _cells =
        new List<(ExcelEntityCellReference, Func<int, IReadOnlyList<TItem>, object>,
            ExcelCellStyle, string)>();
    /// <summary>
    /// 保存尾部相对合并区域。
    /// </summary>
    private readonly List<ExcelEntityCellRange> _merges = new List<ExcelEntityCellRange>();
    /// <summary>
    /// 保存明细与尾部之间的空行数量。
    /// </summary>
    private int _gapRows;
    /// <summary>
    /// 保存结束标记单元格样式。
    /// </summary>
    private ExcelCellStyle _markerStyle;

    /// <summary>
    /// 记录是否声明连续明细求和公式。
    /// </summary>
    private bool _containsContiguousSum;

    /// <summary>
    /// 记录是否声明跨小计明细求和公式。
    /// </summary>
    private bool _containsDetailSum;

    /// <summary>
    /// 设置明细与尾部之间的空行数量。
    /// </summary>
    /// <param name="count">空行数量。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListFooterBuilder<TItem> GapRows(int count)
    {
        if (count < 0)
            throw new ArgumentOutOfRangeException(nameof(count));
        _gapRows = count;
        return this;
    }

    /// <summary>
    /// 添加尾部单元格。
    /// </summary>
    /// <param name="address">相对尾部起点的 A1 地址。</param>
    /// <param name="value">单元格值。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListFooterBuilder<TItem> Cell(string address, object value,
        ExcelCellStyle style = null, string numberFormat = null)
    {
        if (value == null)
            return Cell<object>(address, _ => null, style, numberFormat);
        return Cell(address, _ => value, style, numberFormat);
    }

    /// <summary>
    /// 添加尾部单元格。
    /// </summary>
    /// <typeparam name="TValue">单元格值类型。</typeparam>
    /// <param name="address">相对尾部起点的 A1 地址。</param>
    /// <param name="valueFactory">根据明细读取单元格值的委托。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListFooterBuilder<TItem> Cell<TValue>(string address,
        Func<IReadOnlyList<TItem>, TValue> valueFactory,
        ExcelCellStyle style = null, string numberFormat = null)
    {
        if (valueFactory == null)
            throw new ArgumentNullException(nameof(valueFactory));
        var reference = ExcelEntityCellReference.Parse(address);
        _cells.Add((reference, (_, items) => valueFactory(items), style, numberFormat));
        return this;
    }

    /// <summary>
    /// 添加显式公式尾部单元格。
    /// </summary>
    /// <param name="address">相对尾部起点的 A1 地址。</param>
    /// <param name="formula">以等号开头的 Excel A1 公式。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 公式原样交给 Provider 写入，不由本库求值或随尾部位置改写引用。
    /// 原有 <see cref="Cell(string, object, ExcelCellStyle, string)" /> 的写入行为保持不变。
    /// </remarks>
    public ExcelEntityListFooterBuilder<TItem> Formula(string address, string formula,
        ExcelCellStyle style = null, string numberFormat = null)
    {
        var text = formula?.Trim();
        if (string.IsNullOrWhiteSpace(text) || text[0] != '=' ||
            string.IsNullOrWhiteSpace(text.Substring(1)))
            throw new ArgumentException("尾部公式必须以等号开头且包含公式内容。", nameof(formula));
        return Cell(address, new ExcelEntityFooterFormulaValue(text.Substring(1)), style, numberFormat);
    }

    /// <summary>
    /// 添加对上方连续明细行求和的公式单元格。
    /// </summary>
    /// <param name="address">相对尾部起点的 A1 地址。</param>
    /// <param name="sourceColumn">待求和的工作表列字母。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 仅对本尾部上方连续的明细行求和；空明细写入公式 <c>=0</c>。
    /// 明细与公式单元格之间的空行以及公式单元格相对 marker 的行偏移均会计入引用。
    /// 最终尾部若含分页或分组小计，须改用聚合值或显式公式。
    /// </remarks>
    public ExcelEntityListFooterBuilder<TItem> FormulaSumContiguousRowsAbove(string address,
        string sourceColumn, ExcelCellStyle style = null, string numberFormat = null)
    {
        var column = NormalizeSourceColumn(sourceColumn);
        var reference = ExcelEntityCellReference.Parse(address);
        _cells.Add((reference, (gapRows, items) =>
        {
            if (items.Count == 0)
                return new ExcelEntityFooterFormulaValue("0");
            var rowsBetween = (long)gapRows + reference.Row;
            var firstOffset = rowsBetween + items.Count;
            var lastOffset = rowsBetween + 1;
            return new ExcelEntityFooterFormulaValue(
                $"SUM(INDEX({column}:{column},ROW()-{firstOffset}):" +
                $"INDEX({column}:{column},ROW()-{lastOffset}))");
        }, style, numberFormat));
        _containsContiguousSum = true;
        return this;
    }

    /// <summary>
    /// 添加跨小计的明细总计公式单元格。
    /// </summary>
    /// <param name="address">相对最终尾部起点的 A1 地址。</param>
    /// <param name="sourceColumn">待求和的工作表列字母。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>
    /// 仅用于最终 Footer。Provider 根据实际写出的明细行段生成公式，不计入分页或分组小计行。
    /// 空明细写入 <c>=0</c>；公式由工作簿计算引擎求值。
    /// </remarks>
    public ExcelEntityListFooterBuilder<TItem> FormulaSumDetailRowsAbove(string address,
        string sourceColumn, ExcelCellStyle style = null, string numberFormat = null)
    {
        var column = NormalizeSourceColumn(sourceColumn);
        var reference = ExcelEntityCellReference.Parse(address);
        var value = new ExcelEntityFooterDetailSumValue(column);
        _cells.Add((reference, (_, _) => value, style, numberFormat));
        _containsDetailSum = true;
        return this;
    }

    private static string NormalizeSourceColumn(string sourceColumn)
    {
        var column = sourceColumn?.Trim().ToUpperInvariant();
        if (string.IsNullOrEmpty(column) || column.Any(character => character < 'A' || character > 'Z'))
            throw new ArgumentException("求和来源必须是工作表列字母。", nameof(sourceColumn));
        ExcelEntityCellReference.Parse(column + "1");
        return column;
    }

    /// <summary>
    /// 声明尾部相对合并区域。
    /// </summary>
    /// <param name="range">相对尾部起点的 A1 区域地址。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListFooterBuilder<TItem> Merge(string range)
    {
        _merges.Add(ExcelEntityCellRange.Parse(range));
        return this;
    }

    /// <summary>
    /// 设置明细结束标记的样式。
    /// </summary>
    /// <param name="style">标记单元格样式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityListFooterBuilder<TItem> MarkerStyle(ExcelCellStyle style)
    {
        _markerStyle = style;
        return this;
    }

    /// <summary>
    /// 构建尾部定义并校验单元格占位。
    /// </summary>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <returns>不可变尾部定义。</returns>
    internal ExcelEntityListFooter<TItem> Build(string markerText)
    {
        if (string.IsNullOrWhiteSpace(markerText))
            throw new ArgumentException("尾部标记不能为空。", nameof(markerText));
        var gapRows = _gapRows;
        var cells = _cells.Select(definition => new ExcelEntityListFooterCell<TItem>(
            definition.Reference, items => definition.ValueFactory(gapRows, items),
            definition.Style, definition.NumberFormat)).ToArray();
        var duplicateCell = _cells.GroupBy(cell => (cell.Reference.Row, cell.Reference.Column))
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateCell != null)
            throw new ArgumentException($"尾部包含重复单元格: {duplicateCell.First().Reference.Address}");
        var duplicateMerge = _merges.GroupBy(merge => merge.Address, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateMerge != null)
            throw new ArgumentException($"尾部包含重复合并区域: {duplicateMerge.First().Address}");
        foreach (var pair in _merges.SelectMany((left, index) => _merges.Skip(index + 1)
                     .Where(right => ExcelEntityLayoutValidation.Overlaps(left, right))))
            throw new ArgumentException($"尾部包含重叠合并区域: {pair.Address}");
        if (_cells.Any(cell => cell.Reference.Row == 0 && cell.Reference.Column == 0))
            throw new ArgumentException("尾部 A1 由结束标记保留，不能重复配置。");
        var marker = ExcelEntityCellReference.Parse("A1");
        if (_merges.Any(merge => merge.Contains(marker) && !merge.First.Equals(marker)))
            throw new ArgumentException("尾部合并区域只能将结束标记作为左上角。");
        foreach (var cell in _cells)
        {
            if (_merges.Any(merge => merge.Contains(cell.Reference) && !merge.First.Equals(cell.Reference)))
                throw new ArgumentException($"尾部单元格位于合并区域非左上角: {cell.Reference.Address}");
        }
        return new ExcelEntityListFooter<TItem>(markerText.Trim(), gapRows, cells,
            _merges.ToArray(), _markerStyle, _containsContiguousSum, _containsDetailSum);
    }
}
