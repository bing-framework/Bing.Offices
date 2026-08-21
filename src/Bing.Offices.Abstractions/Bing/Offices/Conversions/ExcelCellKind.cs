namespace Bing.Offices.Conversions;

/// <summary>
/// 与具体 Excel 提供程序无关的逻辑单元格类型。
/// </summary>
public enum ExcelCellKind
{
    /// <summary>
    /// 空单元格。
    /// </summary>
    Empty = 0,

    /// <summary>
    /// 文本单元格。
    /// </summary>
    Text = 1,

    /// <summary>
    /// 数值单元格。
    /// </summary>
    Number = 2,

    /// <summary>
    /// 布尔单元格。
    /// </summary>
    Boolean = 3,

    /// <summary>
    /// 日期或时间单元格。
    /// </summary>
    DateTime = 4,

    /// <summary>
    /// 公式单元格。
    /// </summary>
    Formula = 5,

    /// <summary>
    /// 错误单元格。
    /// </summary>
    Error = 6
}
