using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 的方向化能力描述。
/// </summary>
/// <remarks>
/// 该接口是对历史位标志的可选补充，用于描述只读 Provider、分批导入和格式方向。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelProviderCapabilityDescriptor : IExcelProviderCapabilities
{
    /// <summary>
    /// 获取可读取的工作簿格式。
    /// </summary>
    IReadOnlyList<ExcelFormat> ReadFormats { get; }

    /// <summary>
    /// 获取可写入的工作簿格式。
    /// </summary>
    IReadOnlyList<ExcelFormat> WriteFormats { get; }

    /// <summary>
    /// 获取是否支持完整 Workbook 导入。
    /// </summary>
    bool SupportsCompleteWorkbookImport { get; }

    /// <summary>
    /// 获取是否支持按批次导入。
    /// </summary>
    bool SupportsBatchImport { get; }

    /// <summary>
    /// 获取是否支持完整 Workbook 导出。
    /// </summary>
    bool SupportsCompleteWorkbookExport { get; }

    /// <summary>
    /// 获取是否覆盖真实异步外围 IO。
    /// </summary>
    bool SupportsTrueAsyncIo { get; }

    /// <summary>
    /// 获取 Provider 的明确限制说明。
    /// </summary>
    IReadOnlyList<string> Limitations { get; }
}
