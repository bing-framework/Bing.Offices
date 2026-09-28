namespace Bing.Offices.Imports;

/// <summary>
/// 导入失败工作簿输出模式。
/// </summary>
public enum ExcelImportFailureWorkbookMode
{
    /// <summary>
    /// 不生成失败工作簿。
    /// </summary>
    None,
    /// <summary>
    /// 在原工作簿副本上标记错误。
    /// </summary>
    AnnotatedOriginal,
    /// <summary>
    /// 只输出包含失败行的工作簿。
    /// </summary>
    ErrorRowsOnly
}
