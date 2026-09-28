using System.ComponentModel;

namespace Bing.Offices.Formula;

/// <summary>
/// 公式处理请求。
/// </summary>
public sealed class ExcelFormulaRequest
{
    /// <summary>
    /// 获取或初始化公式读取模式。
    /// </summary>
    public ExcelFormulaReadMode ReadMode { get; init; } = ExcelFormulaReadMode.CachedValue;
    /// <summary>
    /// 获取或初始化公式计算模式。
    /// </summary>
    public ExcelFormulaCalculationMode CalculationMode { get; init; } = ExcelFormulaCalculationMode.DoNotCalculate;
    /// <summary>
    /// 获取或初始化可选的 Sheet 名称；为空表示全部 Sheet。
    /// </summary>
    public string SheetName { get; init; }
}
