namespace Bing.Offices.Configurations;

/// <summary>
/// 显示值映射集合的合并方式。
/// </summary>
public enum ExcelValueMappingMergeMode
{
    /// <summary>
    /// 替换低优先级映射。
    /// </summary>
    Replace,

    /// <summary>
    /// 在低优先级映射后追加。
    /// </summary>
    Append
}
