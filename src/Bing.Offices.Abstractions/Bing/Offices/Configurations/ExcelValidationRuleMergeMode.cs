namespace Bing.Offices.Configurations;

/// <summary>
/// 命名校验规则集合的合并方式。
/// </summary>
public enum ExcelValidationRuleMergeMode
{
    /// <summary>
    /// 替换低优先级规则集合。
    /// </summary>
    Replace,

    /// <summary>
    /// 在低优先级规则后追加。
    /// </summary>
    Append
}
