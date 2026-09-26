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

/// <summary>
/// 单个公式结果。
/// </summary>
public sealed class ExcelFormulaCellResult
{
    /// <summary>
    /// 获取或初始化Sheet 名称。
    /// </summary>
    public string SheetName { get; init; }
    /// <summary>
    /// 获取或初始化零基行索引。
    /// </summary>
    public int RowIndex { get; init; }
    /// <summary>
    /// 获取或初始化零基列索引。
    /// </summary>
    public int ColumnIndex { get; init; }
    /// <summary>
    /// 获取或初始化公式文本；未请求时为空。
    /// </summary>
    public string Formula { get; init; }
    /// <summary>
    /// 获取或初始化缓存值；未请求或不存在时为空。
    /// </summary>
    public object CachedValue { get; init; }
    /// <summary>
    /// 获取或初始化计算后值。
    /// </summary>
    public object CalculatedValue { get; init; }
    /// <summary>
    /// 获取或初始化结构化错误代码。
    /// </summary>
    public string ErrorCode { get; init; }
}

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

/// <summary>
/// 公式分析与计算能力。
/// </summary>
public interface IExcelFormulaProcessor
{
    /// <summary>
    /// 同步分析或计算公式。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="format">输入工作簿格式。</param>
    /// <param name="request">公式读取与计算请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含公式单元格结果与计算完成状态的报告。</returns>
    ExcelFormulaResult Process(Stream source, Bing.Offices.ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步分析或计算公式。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="format">输入工作簿格式。</param>
    /// <param name="request">公式读取与计算请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果为包含公式单元格结果与计算完成状态的报告。</returns>
    Task<ExcelFormulaResult> ProcessAsync(Stream source, Bing.Offices.ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken = default);
}
