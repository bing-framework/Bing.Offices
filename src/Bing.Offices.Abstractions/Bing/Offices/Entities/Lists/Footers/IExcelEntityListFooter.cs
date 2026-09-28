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
