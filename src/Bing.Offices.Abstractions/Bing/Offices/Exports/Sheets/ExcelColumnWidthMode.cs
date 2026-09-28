namespace Bing.Offices.Exports;

using System;

/// <summary>
/// 导出列宽计算模式。
/// </summary>
public enum ExcelColumnWidthMode
{
    /// <summary>
    /// 不修改列宽。
    /// </summary>
    None,
    /// <summary>
    /// 使用固定字符宽度。
    /// </summary>
    Fixed,
    /// <summary>
    /// 使用提供程序自动计算。
    /// </summary>
    AutoFit,
    /// <summary>
    /// 使用受样本限制的自适应估算。
    /// </summary>
    Adaptive
}
