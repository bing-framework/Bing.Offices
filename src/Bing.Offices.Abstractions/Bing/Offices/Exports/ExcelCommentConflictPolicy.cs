namespace Bing.Offices.Exports;

using System;

/// <summary>
/// 批注冲突处理策略。
/// </summary>
public enum ExcelCommentConflictPolicy
{
    /// <summary>
    /// 保留已有批注。
    /// </summary>
    Preserve,
    /// <summary>
    /// 追加新批注文本。
    /// </summary>
    Append,
    /// <summary>
    /// 替换已有批注。
    /// </summary>
    Replace,
    /// <summary>
    /// 出现冲突时失败。
    /// </summary>
    Fail
}
