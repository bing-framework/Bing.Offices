namespace Bing.Offices.Imports;

/// <summary>
/// 校验模式
/// </summary>
public enum ValidateMode
{
    /// <summary>
    /// 校验失败后继续校验
    /// </summary>
    Continue = 0,
    /// <summary>
    /// 校验失败后停止校验
    /// </summary>
    StopOnFirstFailure = 1
}