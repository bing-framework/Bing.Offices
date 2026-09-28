using System.ComponentModel;

namespace Bing.Offices.Formula;

/// <summary>
/// 公式计算语义。
/// </summary>
public enum ExcelFormulaCalculationMode
{
    /// <summary>
    /// 不执行计算。
    /// </summary>
    DoNotCalculate,
    /// <summary>
    /// 重新计算 Provider 明确支持的函数。
    /// </summary>
    RecalculateSupported,
    /// <summary>
    /// 遇到不支持函数、循环或外部引用时失败。
    /// </summary>
    MustRecalculate,
    /// <summary>
    /// 标记 Excel 打开时重新计算。
    /// </summary>
    CalculateOnOpen
}
