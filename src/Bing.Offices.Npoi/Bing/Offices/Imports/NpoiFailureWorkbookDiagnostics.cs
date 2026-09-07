using System.Diagnostics;

namespace Bing.Offices.Imports;

/// <summary>处理失败工作簿可选元数据降级与诊断接收器隔离。</summary>
internal static class NpoiFailureWorkbookDiagnostics
{
    /// <summary>复制 Provider 可选行元数据；不支持时生成结构化降级诊断。</summary>
    internal static void CopyOptionalRowMetadata(Action copy, string propertyName, int rowIndex,
        Action<ExcelImportFailureDiagnostic> diagnosticSink)
    {
        try
        {
            copy();
        }
        catch (NotImplementedException exception)
        {
            var diagnosticException = new NotImplementedException(
                $"失败工作簿无法复制第 {rowIndex + 1} 行的 {propertyName} 元数据。", exception);
            var diagnostic = new ExcelImportFailureDiagnostic(
                "FailureWorkbookRowMetadataUnsupported", null, diagnosticException);
            if (diagnosticSink == null)
            {
                WriteTrace(diagnostic.Code, propertyName, rowIndex, diagnostic.Exception);
                return;
            }
            try
            {
                diagnosticSink(diagnostic);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception sinkException) when (sinkException is not OutOfMemoryException
                && sinkException is not StackOverflowException)
            {
                WriteTrace("FailureWorkbookDiagnosticSinkFailed", propertyName, rowIndex, sinkException);
            }
        }
    }

    private static void WriteTrace(string code, string propertyName, int rowIndex, Exception exception) =>
        Trace.WriteLine($"code={code};property={propertyName};row={rowIndex + 1};exception={exception.GetType().FullName}");
}
