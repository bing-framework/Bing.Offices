namespace Bing.Offices.Imports;

/// <summary>
/// 校验失败后的处理模式。
/// </summary>
public enum ExcelValidationFailureMode
{
    /// <summary>
    /// 校验失败后继续处理后续单元格和行。
    /// </summary>
    Continue = 0,
    /// <summary>
    /// 首次校验失败后停止当前行处理。
    /// </summary>
    StopOnFirstFailure = 1
}
