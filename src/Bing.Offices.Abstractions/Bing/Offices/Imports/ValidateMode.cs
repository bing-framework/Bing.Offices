namespace Bing.Offices.Imports;

/// <summary>
/// 校验模式
/// </summary>
public enum ValidateMode
{
    /// <summary>
    /// 当前行校验失败后继续收集该行可发现的错误；后续行仍会处理。
    /// </summary>
    Continue = 0,
    /// <summary>
    /// 当前行遇到首个校验错误后停止该行校验；不会停止整个 Sheet 或 Workbook。
    /// </summary>
    StopOnFirstFailure = 1
}
