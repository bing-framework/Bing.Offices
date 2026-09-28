namespace Bing.Offices.Imports;

/// <summary>
/// 单元格文本空白规范化策略。
/// </summary>
public enum ExcelWhitespacePolicy
{
    /// <summary>
    /// 保留原始文本。
    /// </summary>
    Preserve,
    /// <summary>
    /// 移除首尾空白。
    /// </summary>
    Trim,
    /// <summary>
    /// 移除全部 Unicode 空白字符。
    /// </summary>
    RemoveAll
}
