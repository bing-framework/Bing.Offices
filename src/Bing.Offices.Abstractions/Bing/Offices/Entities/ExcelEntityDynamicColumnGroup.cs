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
