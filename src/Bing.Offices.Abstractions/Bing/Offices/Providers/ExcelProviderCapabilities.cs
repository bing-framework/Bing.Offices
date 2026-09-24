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
    Xlsx = 128
}

/// <summary>
/// Provider 的只读能力描述。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelProviderCapabilities
{
    /// <summary>
    /// 获取用于诊断和错误上下文的 Provider 稳定名称。
    /// </summary>
    string ProviderName { get; }
    /// <summary>
    /// 获取 Provider 静态声明的能力集合。
    /// </summary>
    ExcelProviderCapabilities Capabilities { get; }
    /// <summary>
    /// 判断 Provider 是否同时支持指定能力组合。
    /// </summary>
    /// <param name="capabilities">要检查的一个或多个能力标志。</param>
    /// <returns>当全部能力均受支持时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    bool Supports(ExcelProviderCapabilities capabilities);
}
