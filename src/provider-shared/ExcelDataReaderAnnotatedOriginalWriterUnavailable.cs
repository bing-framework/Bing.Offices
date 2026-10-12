using System.IO;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 为复用错误行序列化器提供不可用的原始工作簿实现占位。
/// </summary>
/// <remarks>
/// NPOI 和 ClosedXML Forward 只链接 ErrorRowsOnly 所需的序列化代码；AnnotatedOriginal
/// 仍由各 Provider 的原有 DOM 路径负责，因此该实现不应被调用。
/// </remarks>
internal static class ExcelDataReaderAnnotatedOriginalWriter
{
    /// <summary>
    /// 报告当前 Provider 未链接原始工作簿序列化实现。
    /// </summary>
    internal static MemoryStream Serialize(Stream source, ExcelImportFailureOptions options,
        IReadOnlyList<ExcelImportError> errors, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        throw new BingOfficesUnsupportedFeatureException(
            "当前 Provider 的 Forward 失败输出不支持 AnnotatedOriginal 序列化。",
            provider: "Excel", operation: BingOfficesOperation.Import,
            stage: BingOfficesStage.Plan);
    }
}
