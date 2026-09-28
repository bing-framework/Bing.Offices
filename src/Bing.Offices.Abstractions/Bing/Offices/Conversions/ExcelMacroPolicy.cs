namespace Bing.Offices.Conversions;

/// <summary>
/// 宏项目处理策略。
/// </summary>
public enum ExcelMacroPolicy
{
    /// <summary>
    /// 遇到宏项目时拒绝操作。
    /// </summary>
    Reject,
    /// <summary>
    /// 保留宏二进制项目，但不解析、不执行或修改 VBA。
    /// </summary>
    Preserve,
    /// <summary>
    /// 显式删除宏项目。
    /// </summary>
    Strip
}
