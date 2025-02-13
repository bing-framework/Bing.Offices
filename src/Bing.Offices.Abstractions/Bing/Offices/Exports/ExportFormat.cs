using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 导出格式
/// </summary>
public enum ExportFormat
{
    /// <summary>
    /// 97-2003
    /// </summary>
    [Description("Excel2003")]
    Xls,
    /// <summary>
    /// 2007
    /// </summary>
    [Description("Excel2007+")]
    Xlsx,
}
