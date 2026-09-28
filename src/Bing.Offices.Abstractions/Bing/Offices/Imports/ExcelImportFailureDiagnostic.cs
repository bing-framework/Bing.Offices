namespace Bing.Offices.Imports;

/// <summary>
/// 失败工作簿输出边界的结构化诊断。
/// </summary>
public sealed class ExcelImportFailureDiagnostic
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelImportFailureDiagnostic" /> 类型的实例。
    /// </summary>
    /// <param name="code">诊断代码。</param>
    /// <param name="temporaryPath">未包含工作簿内容的临时文件路径。</param>
    /// <param name="exception">清理或输出阶段捕获的异常。</param>
    public ExcelImportFailureDiagnostic(string code, string temporaryPath, Exception exception)
    {
        Code = code;
        TemporaryPath = temporaryPath;
        Exception = exception;
    }

    /// <summary>
    /// 获取诊断代码。
    /// </summary>
    public string Code { get; }

    /// <summary>
    /// 获取未包含工作簿内容的临时文件路径。
    /// </summary>
    public string TemporaryPath { get; }

    /// <summary>
    /// 获取清理异常。
    /// </summary>
    public Exception Exception { get; }
}
