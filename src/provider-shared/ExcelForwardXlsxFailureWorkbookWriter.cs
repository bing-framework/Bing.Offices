using Bing.Offices.Imports;
using Bing.Offices.IO;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 写出 Forward XLSX 的低保留错误行工作簿。
/// </summary>
internal static class ExcelForwardXlsxFailureWorkbookWriter
{
    /// <summary>
    /// 同步写出错误行工作簿。
    /// </summary>
    internal static void Write(ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter)
    {
        ExcelDataReaderFailureWorkbookWriter.Write(options, errors, capture, cancellationToken,
            fileExportCommitter, summarySheetName: "_ImportErrors");
    }

    /// <summary>
    /// 异步写出错误行工作簿。
    /// </summary>
    internal static Task WriteAsync(ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter)
        => ExcelDataReaderFailureWorkbookWriter.WriteAsync(options, errors, capture,
            cancellationToken, fileExportCommitter, summarySheetName: "_ImportErrors");
}
