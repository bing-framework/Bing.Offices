namespace Bing.Offices.Imports;

/// <summary>
/// 导入失败批注与原有批注冲突时的处理策略。
/// </summary>
public enum ExcelImportCommentConflictPolicy
{
    /// <summary>
    /// 保留已有批注，不追加失败信息。
    /// </summary>
    Preserve,
    /// <summary>
    /// 在已有批注后追加失败信息。
    /// </summary>
    Append,
    /// <summary>
    /// 用失败信息替换已有批注。
    /// </summary>
    Replace,
    /// <summary>
    /// 存在已有批注时直接失败。
    /// </summary>
    Fail
}
