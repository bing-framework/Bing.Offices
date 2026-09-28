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
