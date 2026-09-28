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
