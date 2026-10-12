using System;
using System.Collections.Generic;
using System.Linq;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体布局分批导入完成摘要。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityBatchImportSummary<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityBatchImportSummary{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="result">完整实体导入结果。</param>
    /// <param name="regions">按布局顺序排列的列表区域摘要。</param>
    /// <param name="globalErrors">无法归属单个列表区域的实体级错误。</param>
    public ExcelEntityBatchImportSummary(ExcelEntityImportResult<TEntity> result,
        IReadOnlyList<ExcelEntityBatchRegionSummary> regions,
        IReadOnlyList<Imports.ExcelImportError> globalErrors)
        : this(result, regions, globalErrors, hasCompleteResult: true)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityBatchImportSummary{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="result">实体导入结果；摘要模式下实体为空聚合对象，仅包含错误和工作表诊断。</param>
    /// <param name="regions">按布局顺序排列的列表区域摘要。</param>
    /// <param name="globalErrors">无法归属单个列表区域的实体级错误。</param>
    /// <param name="hasCompleteResult">是否保留完整实体集合和来源行索引。</param>
    public ExcelEntityBatchImportSummary(ExcelEntityImportResult<TEntity> result,
        IReadOnlyList<ExcelEntityBatchRegionSummary> regions,
        IReadOnlyList<Imports.ExcelImportError> globalErrors, bool hasCompleteResult)
    {
        Result = result ?? throw new ArgumentNullException(nameof(result));
        Regions = Array.AsReadOnly((regions ?? Array.Empty<ExcelEntityBatchRegionSummary>()).ToArray());
        GlobalErrors = Array.AsReadOnly((globalErrors ?? Array.Empty<Imports.ExcelImportError>()).ToArray());
        HasCompleteResult = hasCompleteResult;
    }

    /// <summary>
    /// 获取完整实体导入结果。
    /// </summary>
    public ExcelEntityImportResult<TEntity> Result { get; }

    /// <summary>
    /// 获取是否保留完整实体集合和来源行索引。
    /// </summary>
    /// <remarks>
    /// <see langword="false" /> 表示调用方选择了摘要模式；此时仍可读取批次回调、统计和错误，
    /// 但不应依赖 <see cref="Result" /> 中的实体字段、列表集合或来源行索引。
    /// </remarks>
    public bool HasCompleteResult { get; }

    /// <summary>
    /// 获取按布局顺序排列的列表区域摘要。
    /// </summary>
    public IReadOnlyList<ExcelEntityBatchRegionSummary> Regions { get; }

    /// <summary>
    /// 获取无法归属单个列表区域的实体级错误。
    /// </summary>
    public IReadOnlyList<Imports.ExcelImportError> GlobalErrors { get; }

    /// <summary>
    /// 获取所有列表区域读取的正文行数。
    /// </summary>
    public int RowsRead => Regions.Sum(region => region.Summary.RowsRead);

    /// <summary>
    /// 获取所有列表区域成功实体行数。
    /// </summary>
    public int SucceededRows => Regions.Sum(region => region.Summary.SucceededRows);

    /// <summary>
    /// 获取所有列表区域失败行数。
    /// </summary>
    public int FailedRows => Regions.Sum(region => region.Summary.FailedRows);

    /// <summary>
    /// 获取完整实体导入收集的错误数量。
    /// </summary>
    public int ErrorCount => Result.Errors.Count;

    /// <summary>
    /// 获取错误集合是否因达到数量上限而截断。
    /// </summary>
    public bool ErrorsTruncated => Result.ErrorsTruncated;

    /// <summary>
    /// 获取本次实体导入是否成功。
    /// </summary>
    public bool IsSuccess => Result.IsSuccess;
}
