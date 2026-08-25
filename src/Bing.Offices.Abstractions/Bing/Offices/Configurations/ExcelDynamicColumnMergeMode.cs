namespace Bing.Offices.Configurations;

/// <summary>
/// 动态列集合的合并方式。
/// </summary>
public enum ExcelDynamicColumnMergeMode
{
    /// <summary>替换低优先级动态列集合。</summary>
    Replace = 0,

    /// <summary>按稳定 Key 更新或追加动态列。</summary>
    Append = 1
}
