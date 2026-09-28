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
