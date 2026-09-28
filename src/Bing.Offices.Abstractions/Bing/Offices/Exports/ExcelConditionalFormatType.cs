using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 条件格式类型。
/// </summary>
public enum ExcelConditionalFormatType
{
    /// <summary>
    /// 单元格值比较。
    /// </summary>
    CellValue,
    /// <summary>
    /// 简单公式。
    /// </summary>
    Formula,
    /// <summary>
    /// 色阶。
    /// </summary>
    ColorScale,
    /// <summary>
    /// 数据条。
    /// </summary>
    DataBar,
    /// <summary>
    /// 图标集。
    /// </summary>
    IconSet
}
