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
