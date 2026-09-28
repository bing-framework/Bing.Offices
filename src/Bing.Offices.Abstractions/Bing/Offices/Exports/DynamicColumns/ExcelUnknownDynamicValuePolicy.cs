namespace Bing.Offices.Exports;

/// <summary>
/// 未知动态值处理策略。
/// </summary>
public enum ExcelUnknownDynamicValuePolicy
{
    /// <summary>
    /// 忽略未在定义中声明的值。
    /// </summary>
    Ignore,

    /// <summary>
    /// 遇到未声明值时拒绝本次导出。
    /// </summary>
    Fail
}
