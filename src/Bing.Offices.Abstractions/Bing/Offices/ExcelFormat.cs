using System.ComponentModel;

namespace Bing.Offices;

/// <summary>
/// Excel 格式。
/// </summary>
public enum ExcelFormat
{
    /// <summary>
    /// Excel 97-2003 二进制工作簿格式。
    /// </summary>
    [Description("Excel2003")]
    Xls = 0,
    /// <summary>
    /// Excel 2007 及更高版本的 Open XML 工作簿格式。
    /// </summary>
    [Description("Excel2007+")]
    Xlsx = 1,
    /// <summary>
    /// Excel 二进制工作簿格式。
    /// </summary>
    [Description("ExcelBinary")]
    Xlsb = 2,
    /// <summary>
    /// 含 VBA 项目的 Open XML 工作簿格式。
    /// </summary>
    [Description("ExcelMacroEnabled")]
    Xlsm = 3,
    /// <summary>
    /// OpenDocument 表格格式。
    /// </summary>
    [Description("OpenDocumentSpreadsheet")]
    Ods = 4
}
