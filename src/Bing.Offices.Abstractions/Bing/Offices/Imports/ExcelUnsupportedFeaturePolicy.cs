namespace Bing.Offices.Imports;

/// <summary>
/// 不支持的工作簿特性处理策略。
/// </summary>
public enum ExcelUnsupportedFeaturePolicy
{
    /// <summary>
    /// 报告为导入错误。
    /// </summary>
    Report,
    /// <summary>
    /// 直接拒绝导入。
    /// </summary>
    Fail
}
