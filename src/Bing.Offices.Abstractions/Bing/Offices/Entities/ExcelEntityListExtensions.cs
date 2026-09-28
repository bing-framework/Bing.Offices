using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Bing.Offices.Exports;
using Bing.Offices.Styles;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体列表动态列组的 Provider SPI 描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityDynamicColumnGroup
{
    /// <summary>
    /// 获取动态列组的稳定标识。
    /// </summary>
    string GroupKey { get; }
    /// <summary>
    /// 获取动态值字典属性。
    /// </summary>
    System.Reflection.PropertyInfo Property { get; }
    /// <summary>
    /// 获取读取动态值字典的委托。
    /// </summary>
    Func<object, IDictionary<string, object>> Getter { get; }
    /// <summary>
    /// 获取写入动态值字典的委托。
    /// </summary>
    Action<object, object> Setter { get; }
    /// <summary>
    /// 获取动态列定义集合。
    /// </summary>
    IReadOnlyList<ExcelDynamicColumnDefinition> Definitions { get; }
}

/// <summary>
/// 实体列表计算列的 Provider SPI 描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityCalculatedColumn
{
    /// <summary>
    /// 获取计算列的稳定标识。
    /// </summary>
    string Key { get; }

    /// <summary>
    /// 获取计算列标题。
    /// </summary>
    string Title { get; }

    /// <summary>
    /// 获取计算列值类型。
    /// </summary>
    Type ValueType { get; }

    /// <summary>
    /// 获取计算列的相对位置。
    /// </summary>
    ExcelColumnPlacement Placement { get; }

    /// <summary>
    /// 获取计算列的默认排序值。
    /// </summary>
    int Order { get; }

    /// <summary>
    /// 获取显式物理列索引。
    /// </summary>
    int? PhysicalColumnIndex { get; }

    /// <summary>
    /// 获取计算列表头样式。
    /// </summary>
    ExcelCellStyle HeaderStyle { get; }

    /// <summary>
    /// 获取计算列正文样式。
    /// </summary>
    ExcelCellStyle BodyStyle { get; }

    /// <summary>
    /// 获取计算列数字格式。
    /// </summary>
    string NumberFormat { get; }

    /// <summary>
    /// 为一次导出创建计算值委托。
    /// </summary>
    /// <param name="items">本次导出的明细快照。</param>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <returns>接收项目、项目索引、工作表行号和列号并返回单元格值的委托。</returns>
    Func<object, int, int, int, object> CreateEvaluator(IReadOnlyList<object> items, string sheetName);
}

/// <summary>
/// 实体列表连续分组小计的 Provider SPI 描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityGroupSubtotal
{
    /// <summary>
    /// 获取分组键类型。
    /// </summary>
    Type KeyType { get; }

    /// <summary>
    /// 获取当前项目的分组键。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <returns>列表项的分组键。</returns>
    object GetKey(object item);

    /// <summary>
    /// 获取分组小计尾部定义。
    /// </summary>
    IExcelEntityListFooter Footer { get; }
}

/// <summary>
/// 实体列表连续分组小计定义。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
/// <typeparam name="TKey">分组键类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityGroupSubtotal<TItem, TKey> : IExcelEntityGroupSubtotal
    where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityGroupSubtotal{TItem, TKey}" /> 类型的实例。
    /// </summary>
    /// <param name="keySelector">分组键读取委托。</param>
    /// <param name="footer">分组小计尾部定义。</param>
    internal ExcelEntityGroupSubtotal(Func<TItem, TKey> keySelector, IExcelEntityListFooter footer)
    {
        KeySelector = keySelector ?? throw new ArgumentNullException(nameof(keySelector));
        Footer = footer ?? throw new ArgumentNullException(nameof(footer));
    }

    /// <inheritdoc />
    public Type KeyType => typeof(TKey);

    /// <summary>
    /// 获取分组键读取委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<TItem, TKey> KeySelector { get; }

    /// <inheritdoc />
    public object GetKey(object item) => KeySelector((TItem)item);

    /// <inheritdoc />
    public IExcelEntityListFooter Footer { get; }
}

/// <summary>
/// 实体列表计算列读取当前行时使用的上下文。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityRowContext<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityRowContext{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <param name="items">本次导出的明细快照。</param>
    /// <param name="index">当前列表项的零基索引。</param>
    /// <param name="rowIndex">当前工作表行的零基索引。</param>
    /// <param name="columnIndex">当前工作表列的零基索引。</param>
    /// <param name="sheetName">当前工作表名称。</param>
    /// <param name="columnKey">当前计算列标识。</param>
    internal ExcelEntityRowContext(TItem item, IReadOnlyList<TItem> items, int index,
        int rowIndex, int columnIndex, string sheetName, string columnKey)
    {
        Item = item ?? throw new ArgumentNullException(nameof(item));
        Items = items ?? throw new ArgumentNullException(nameof(items));
        Index = index;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        ColumnKey = columnKey ?? throw new ArgumentNullException(nameof(columnKey));
    }

    /// <summary>
    /// 获取当前列表项。
    /// </summary>
    public TItem Item { get; }

    /// <summary>
    /// 获取本次导出的明细快照。
    /// </summary>
    public IReadOnlyList<TItem> Items { get; }

    /// <summary>
    /// 获取当前列表项的零基索引。
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// 获取当前工作表行的零基索引。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// 获取当前工作表行的一基行号。
    /// </summary>
    public int RowNumber => RowIndex + 1;

    /// <summary>
    /// 获取当前工作表列的零基索引。
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// 获取当前工作表列的一基列号。
    /// </summary>
    public int ColumnNumber => ColumnIndex + 1;

    /// <summary>
    /// 获取当前工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取当前计算列标识。
    /// </summary>
    public string ColumnKey { get; }
}

/// <summary>
/// 实体列表计算列的配置构建器。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityCalculatedColumnBuilder
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCalculatedColumnBuilder" /> 类型的实例。
    /// </summary>
    public ExcelEntityCalculatedColumnBuilder()
    {
    }

    /// <summary>
    /// 保存计算列的排序值。
    /// </summary>
    private int _order;

    /// <summary>
    /// 保存计算列的位置配置。
    /// </summary>
    private ExcelColumnPlacement _placement;

    /// <summary>
    /// 保存计算列表头样式。
    /// </summary>
    private ExcelCellStyle _headerStyle;

    /// <summary>
    /// 保存计算列正文样式。
    /// </summary>
    private ExcelCellStyle _bodyStyle;

    /// <summary>
    /// 保存计算列数字格式。
    /// </summary>
    private string _numberFormat;

    /// <summary>
    /// 设置计算列的默认排序值。
    /// </summary>
    /// <param name="order">排序值，值越小越靠前。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder Order(int order)
    {
        _order = order;
        return this;
    }

    /// <summary>
    /// 设置计算列的相对位置或物理列位置。
    /// </summary>
    /// <param name="placement">列位置。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder Placement(ExcelColumnPlacement placement)
    {
        _placement = placement ?? throw new ArgumentNullException(nameof(placement));
        return this;
    }

    /// <summary>
    /// 设置计算列表头样式。
    /// </summary>
    /// <param name="style">表头样式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder HeaderStyle(ExcelCellStyle style)
    {
        _headerStyle = style;
        return this;
    }

    /// <summary>
    /// 设置计算列正文样式。
    /// </summary>
    /// <param name="style">正文样式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder BodyStyle(ExcelCellStyle style)
    {
        _bodyStyle = style;
        return this;
    }

    /// <summary>
    /// 设置计算列数字格式。
    /// </summary>
    /// <param name="numberFormat">数字格式字符串。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCalculatedColumnBuilder NumberFormat(string numberFormat)
    {
        _numberFormat = numberFormat;
        return this;
    }

    /// <summary>
    /// 获取计算列排序值。
    /// </summary>
    internal int OrderValue => _order;

    /// <summary>
    /// 获取计算列相对位置。
    /// </summary>
    internal ExcelColumnPlacement PlacementValue => _placement;

    /// <summary>
    /// 获取计算列表头样式。
    /// </summary>
    internal ExcelCellStyle HeaderStyleValue => _headerStyle;

    /// <summary>
    /// 获取计算列正文样式。
    /// </summary>
    internal ExcelCellStyle BodyStyleValue => _bodyStyle;

    /// <summary>
    /// 获取计算列数字格式。
    /// </summary>
    internal string NumberFormatValue => _numberFormat;
}

/// <summary>
/// 实体列表计算列定义。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityCalculatedColumn<TItem> : IExcelEntityCalculatedColumn
    where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCalculatedColumn{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="key">计算列标识。</param>
    /// <param name="title">计算列标题。</param>
    /// <param name="valueFactory">计算值委托。</param>
    /// <param name="valueType">计算值类型。</param>
    /// <param name="builder">列配置构建器。</param>
    internal ExcelEntityCalculatedColumn(string key, string title,
        Func<ExcelEntityRowContext<TItem>, object> valueFactory, Type valueType,
        ExcelEntityCalculatedColumnBuilder builder)
    {
        Key = key ?? throw new ArgumentNullException(nameof(key));
        Title = title ?? throw new ArgumentNullException(nameof(title));
        ValueFactory = valueFactory ?? throw new ArgumentNullException(nameof(valueFactory));
        ValueType = valueType ?? throw new ArgumentNullException(nameof(valueType));
        Order = builder?.OrderValue ?? 0;
        Placement = builder?.PlacementValue;
        HeaderStyle = builder?.HeaderStyleValue;
        BodyStyle = builder?.BodyStyleValue;
        NumberFormat = builder?.NumberFormatValue;
        PhysicalColumnIndex = Placement?.PhysicalColumnIndex;
    }

    /// <inheritdoc />
    public string Key { get; }

    /// <inheritdoc />
    public string Title { get; }

    /// <inheritdoc />
    public Type ValueType { get; }

    /// <inheritdoc />
    public ExcelColumnPlacement Placement { get; }

    /// <inheritdoc />
    public int Order { get; }

    /// <inheritdoc />
    public int? PhysicalColumnIndex { get; }

    /// <inheritdoc />
    public ExcelCellStyle HeaderStyle { get; }

    /// <inheritdoc />
    public ExcelCellStyle BodyStyle { get; }

    /// <inheritdoc />
    public string NumberFormat { get; }

    /// <summary>
    /// 获取计算值委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<ExcelEntityRowContext<TItem>, object> ValueFactory { get; }

    /// <inheritdoc />
    public Func<object, int, int, int, object> CreateEvaluator(IReadOnlyList<object> items,
        string sheetName)
    {
        var snapshot = Array.AsReadOnly((items ?? Array.Empty<object>()).Cast<TItem>().ToArray());
        return (item, index, rowNumber, columnNumber) => ValueFactory(new ExcelEntityRowContext<TItem>(
            (TItem)item, snapshot, index, rowNumber - 1, columnNumber - 1, sheetName, Key));
    }
}

/// <summary>
/// 实体列表尾部的 Provider SPI 描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityListFooter
{
    /// <summary>
    /// 获取明细结束标记文本。
    /// </summary>
    string MarkerText { get; }
    /// <summary>
    /// 获取明细与尾部之间的空行数量。
    /// </summary>
    int GapRows { get; }
    /// <summary>
    /// 获取尾部单元格集合。
    /// </summary>
    IReadOnlyList<IExcelEntityListFooterCell> Cells { get; }
    /// <summary>
    /// 获取尾部合并区域集合。
    /// </summary>
    IReadOnlyList<ExcelEntityCellRange> Merges { get; }
    /// <summary>
    /// 获取标记单元格样式。
    /// </summary>
    ExcelCellStyle MarkerStyle { get; }
}

/// <summary>
/// 实体列表尾部单元格的 Provider SPI 描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityListFooterCell
{
    /// <summary>
    /// 获取尾部相对单元格坐标。
    /// </summary>
    ExcelEntityCellReference Reference { get; }
    /// <summary>
    /// 根据明细快照计算单元格值。
    /// </summary>
    /// <param name="items">本次导入或导出的明细快照。</param>
    /// <returns>尾部单元格值。</returns>
    object Evaluate(IReadOnlyList<object> items);
    /// <summary>
    /// 获取单元格样式。
    /// </summary>
    ExcelCellStyle Style { get; }
    /// <summary>
    /// 获取数字格式。
    /// </summary>
    string NumberFormat { get; }
}

/// <summary>
/// 实体列表尾部的显式公式值。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityFooterFormulaValue
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityFooterFormulaValue" /> 类型的实例。
    /// </summary>
    /// <param name="formulaA1">不含等号的 Excel A1 公式文本。</param>
    internal ExcelEntityFooterFormulaValue(string formulaA1)
    {
        FormulaA1 = formulaA1 ?? throw new ArgumentNullException(nameof(formulaA1));
    }

    /// <summary>
    /// 获取不含等号的 Excel A1 公式文本。
    /// </summary>
    public string FormulaA1 { get; }
}

/// <summary>
/// 最终尾部跨小计明细求和的 Provider SPI 值。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityFooterDetailSumValue
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityFooterDetailSumValue" /> 类型的实例。
    /// </summary>
    /// <param name="sourceColumn">待求和的工作表列字母。</param>
    internal ExcelEntityFooterDetailSumValue(string sourceColumn)
    {
        SourceColumn = sourceColumn ?? throw new ArgumentNullException(nameof(sourceColumn));
    }

    /// <summary>
    /// 获取待求和的工作表列字母。
    /// </summary>
    public string SourceColumn { get; }

    /// <summary>
    /// 根据一基明细行段生成不含等号的公式。
    /// </summary>
    /// <param name="detailRows">按写出顺序排列的一基闭区间明细行段。</param>
    /// <returns>不含等号的 Excel A1 公式。</returns>
    /// <remarks>
    /// 每段只包含真实明细行，间隔、小计和尾部行应由 Provider 排除。空明细生成 <c>0</c>。
    /// </remarks>
    public string ToFormulaA1(IReadOnlyList<(int FirstRow, int LastRow)> detailRows)
    {
        if (detailRows == null)
            throw new ArgumentNullException(nameof(detailRows));
        if (detailRows.Count == 0)
            return "0";
        var formula = new StringBuilder();
        var previousLastRow = 0;
        for (var index = 0; index < detailRows.Count; index++)
        {
            var (firstRow, lastRow) = detailRows[index];
            if (firstRow <= previousLastRow || lastRow < firstRow)
                throw new ArgumentException("明细行段必须是递增且不重叠的一基闭区间。", nameof(detailRows));
            if (index % 255 == 0)
            {
                if (index > 0)
                    formula.Append("+");
                formula.Append("SUM(");
            }
            else
                formula.Append(",");
            formula.Append(SourceColumn).Append(firstRow).Append(":")
                .Append(SourceColumn).Append(lastRow);
            if (index % 255 == 254 || index == detailRows.Count - 1)
                formula.Append(")");
            if (formula.Length > 8192)
                throw new ArgumentException("明细行段生成的公式超出 Excel 长度限制。", nameof(detailRows));
            previousLastRow = lastRow;
        }
        return formula.ToString();
    }
}

/// <summary>
/// 实体列表区域使用的动态列来源描述。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityDynamicColumnGroup<TItem> : IExcelEntityDynamicColumnGroup where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityDynamicColumnGroup{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="groupKey">动态列组的稳定标识。</param>
    /// <param name="property">动态值字典属性。</param>
    /// <param name="getter">读取动态值字典的委托。</param>
    /// <param name="setter">写入动态值字典的委托。</param>
    /// <param name="definitions">动态列定义。</param>
    internal ExcelEntityDynamicColumnGroup(string groupKey, System.Reflection.PropertyInfo property,
        Func<object, IDictionary<string, object>> getter, Action<object, object> setter,
        IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        GroupKey = groupKey ?? throw new ArgumentNullException(nameof(groupKey));
        Property = property ?? throw new ArgumentNullException(nameof(property));
        Getter = getter ?? throw new ArgumentNullException(nameof(getter));
        Setter = setter;
        Definitions = definitions ?? throw new ArgumentNullException(nameof(definitions));
    }

    /// <inheritdoc />
    public string GroupKey { get; }

    /// <inheritdoc />
    [EditorBrowsable(EditorBrowsableState.Never)]
    public System.Reflection.PropertyInfo Property { get; }

    /// <inheritdoc />
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<object, IDictionary<string, object>> Getter { get; }

    /// <inheritdoc />
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Action<object, object> Setter { get; }

    /// <inheritdoc />
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelDynamicColumnDefinition> Definitions { get; }
}

/// <summary>
/// 实体列表区域的导出尾部定义。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityListFooter<TItem> : IExcelEntityListFooter where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityListFooter{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="markerText">明细结束标记文本。</param>
    /// <param name="gapRows">明细与尾部之间的空行数量。</param>
    /// <param name="cells">尾部单元格定义。</param>
    /// <param name="merges">尾部相对合并区域。</param>
    /// <param name="markerStyle">标记单元格样式。</param>
    /// <param name="containsContiguousSum">是否包含连续明细求和公式。</param>
    /// <param name="containsDetailSum">是否包含跨小计明细求和公式。</param>
    internal ExcelEntityListFooter(string markerText, int gapRows,
        IReadOnlyList<ExcelEntityListFooterCell<TItem>> cells,
        IReadOnlyList<ExcelEntityCellRange> merges, ExcelCellStyle markerStyle,
        bool containsContiguousSum, bool containsDetailSum)
    {
        MarkerText = markerText ?? throw new ArgumentNullException(nameof(markerText));
        GapRows = gapRows;
        Cells = cells ?? throw new ArgumentNullException(nameof(cells));
        Merges = merges ?? throw new ArgumentNullException(nameof(merges));
        MarkerStyle = markerStyle;
        ContainsContiguousSum = containsContiguousSum;
        ContainsDetailSum = containsDetailSum;
    }

    /// <inheritdoc />
    public string MarkerText { get; }

    /// <inheritdoc />
    public int GapRows { get; }

    /// <summary>
    /// 获取尾部单元格定义。
    /// </summary>
    public IReadOnlyList<ExcelEntityListFooterCell<TItem>> Cells { get; }

    /// <inheritdoc />
    IReadOnlyList<IExcelEntityListFooterCell> IExcelEntityListFooter.Cells =>
        Cells.Cast<IExcelEntityListFooterCell>().ToArray();

    /// <inheritdoc />
    public IReadOnlyList<ExcelEntityCellRange> Merges { get; }

    /// <inheritdoc />
    public ExcelCellStyle MarkerStyle { get; }

    /// <summary>
    /// 获取尾部是否包含连续明细求和公式。
    /// </summary>
    internal bool ContainsContiguousSum { get; }

    /// <summary>
    /// 获取尾部是否包含跨小计明细求和公式。
    /// </summary>
    internal bool ContainsDetailSum { get; }
}

/// <summary>
/// 实体列表区域尾部的单元格定义。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelEntityListFooterCell<TItem> : IExcelEntityListFooterCell where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityListFooterCell{TItem}" /> 类型的实例。
    /// </summary>
    /// <param name="reference">相对单元格坐标。</param>
    /// <param name="valueFactory">读取输出值的委托。</param>
    /// <param name="style">单元格样式。</param>
    /// <param name="numberFormat">数字格式。</param>
    internal ExcelEntityListFooterCell(ExcelEntityCellReference reference,
        Func<IReadOnlyList<TItem>, object> valueFactory, ExcelCellStyle style, string numberFormat)
    {
        Reference = reference;
        ValueFactory = valueFactory ?? throw new ArgumentNullException(nameof(valueFactory));
        Style = style;
        NumberFormat = numberFormat;
    }

    /// <inheritdoc />
    public ExcelEntityCellReference Reference { get; }

    /// <summary>
    /// 获取根据明细读取输出值的委托。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public Func<IReadOnlyList<TItem>, object> ValueFactory { get; }

    /// <inheritdoc />
    object IExcelEntityListFooterCell.Evaluate(IReadOnlyList<object> items)
    {
        return ValueFactory(items.Cast<TItem>().ToArray());
    }

    /// <inheritdoc />
    public ExcelCellStyle Style { get; }

    /// <inheritdoc />
    public string NumberFormat { get; }
}

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
