using System.ComponentModel;

namespace Bing.Offices.Formula;

/// <summary>
/// 公式处理结果。
/// </summary>
public sealed class ExcelFormulaResult
{
    /// <summary>
    /// 获取或初始化公式单元格结果。
    /// </summary>
    public IReadOnlyList<ExcelFormulaCellResult> Cells { get; init; } = Array.Empty<ExcelFormulaCellResult>();
    /// <summary>
    /// 获取或初始化是否已完成请求的计算语义。
    /// </summary>
    public bool CalculationCompleted { get; init; }
}
