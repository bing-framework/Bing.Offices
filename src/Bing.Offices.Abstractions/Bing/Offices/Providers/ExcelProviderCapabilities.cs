using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 声明的静态能力集合。
/// </summary>
[Flags]
[EditorBrowsable(EditorBrowsableState.Never)]
public enum ExcelProviderCapabilities
{
    /// <summary>
    /// 不声明任何 Excel Provider 能力。
    /// </summary>
    None = 0,
    /// <summary>
    /// 支持列表导入或导出操作。
    /// </summary>
    List = 1,
    /// <summary>
    /// 支持 Workbook 级多工作表操作。
    /// </summary>
    Workbook = 2,
    /// <summary>
    /// 支持 Entity 固定单元格和列表区域布局。
    /// </summary>
    Entity = 4,
    /// <summary>
    /// 支持基于既有工作簿模板的操作。
    /// </summary>
    Template = 8,
    /// <summary>
    /// 支持读取或写入合并区域。
    /// </summary>
    Merge = 16,
    /// <summary>
    /// 支持可取消的异步外围 IO。
    /// </summary>
    Async = 32,
    /// <summary>
    /// 支持 XLS/HSSF 工作簿格式。
    /// </summary>
    Xls = 64,
    /// <summary>
    /// 支持 XLSX/XSSF 工作簿格式。
    /// </summary>
    Xlsx = 128,
    /// <summary>
    /// 支持 XLSB 工作簿格式。
    /// </summary>
    Xlsb = 256,
    /// <summary>
    /// 支持 XLSM 工作簿格式。
    /// </summary>
    Xlsm = 512,
    /// <summary>
    /// 支持 ODS 工作簿格式。
    /// </summary>
    Ods = 1024
}
