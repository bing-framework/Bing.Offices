using Bing.Offices.Exceptions;
using Bing.Offices.Providers;

namespace Bing.Offices.Exports;

/// <summary>
/// 工作簿导出执行模式。
/// </summary>
public enum ExcelWorkbookExportMode
{
    /// <summary>
    /// 使用完整工作簿导出契约。
    /// </summary>
    CompleteWorkbook,

    /// <summary>
    /// 使用前向流式导出契约。
    /// </summary>
    ForwardStreaming
}
