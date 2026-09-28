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
