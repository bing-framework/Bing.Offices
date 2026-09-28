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
