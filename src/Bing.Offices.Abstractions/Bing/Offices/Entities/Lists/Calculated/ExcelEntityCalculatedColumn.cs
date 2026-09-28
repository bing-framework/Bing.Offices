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
