using System.ComponentModel;

namespace Bing.Offices.Formula;

/// <summary>
/// 公式读取语义。
/// </summary>
public enum ExcelFormulaReadMode
{
    /// <summary>
    /// 只返回缓存值。
    /// </summary>
    CachedValue,
    /// <summary>
    /// 只返回公式文本。
    /// </summary>
    FormulaText,
    /// <summary>
    /// 同时返回公式文本和缓存值。
    /// </summary>
    FormulaAndCachedValue
}
