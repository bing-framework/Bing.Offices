using System.ComponentModel;

namespace Bing.Offices.Formula;

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
