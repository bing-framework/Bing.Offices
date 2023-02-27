using System.ComponentModel;

namespace Bing.Offices;

/// <summary>
/// Excel 格式
/// </summary>
public enum ExcelFormat
{
    /// <summary>
    /// 97-2003
    /// </summary>
    [Description("Excel2003")]
    Xls = 0,
    /// <summary>
    /// 2007
    /// </summary>
    [Description("Excel2007+")]
    Xlsx = 1
}