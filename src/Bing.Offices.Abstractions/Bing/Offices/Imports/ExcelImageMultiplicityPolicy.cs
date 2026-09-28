namespace Bing.Offices.Imports;

/// <summary>
/// 图片列多图片处理策略。
/// </summary>
public enum ExcelImageMultiplicityPolicy
{
    /// <summary>
    /// 只绑定第一张图片。
    /// </summary>
    First,
    /// <summary>
    /// 绑定全部图片。
    /// </summary>
    All,
    /// <summary>
    /// 出现多张图片时报告错误。
    /// </summary>
    Fail
}
