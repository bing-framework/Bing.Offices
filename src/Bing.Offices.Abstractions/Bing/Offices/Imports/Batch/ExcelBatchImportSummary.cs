using System.Globalization;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// 分批导入完成摘要。
/// </summary>
public sealed class ExcelBatchImportSummary
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelBatchImportSummary"/> 类型的实例。
    /// </summary>
    /// <param name="rowsRead">读取的正文行数。</param>
    /// <param name="succeededRows">成功实体行数。</param>
    /// <param name="failedRows">失败行数。</param>
    /// <param name="errorCount">收集的错误数量。</param>
    /// <param name="errorsTruncated">是否达到错误预算。</param>
    /// <param name="resourceLimitExceeded">是否达到资源预算。</param>
    public ExcelBatchImportSummary(int rowsRead, int succeededRows, int failedRows,
        int errorCount, bool errorsTruncated, bool resourceLimitExceeded)
    {
        RowsRead = rowsRead;
        SucceededRows = succeededRows;
        FailedRows = failedRows;
        ErrorCount = errorCount;
        ErrorsTruncated = errorsTruncated;
        ResourceLimitExceeded = resourceLimitExceeded;
    }

    /// <summary>
    /// 获取读取的正文行数。
    /// </summary>
    public int RowsRead { get; }
    /// <summary>
    /// 获取成功实体行数。
    /// </summary>
    public int SucceededRows { get; }
    /// <summary>
    /// 获取失败行数。
    /// </summary>
    public int FailedRows { get; }
    /// <summary>
    /// 获取收集的错误数量。
    /// </summary>
    public int ErrorCount { get; }
    /// <summary>
    /// 获取是否达到错误预算。
    /// </summary>
    public bool ErrorsTruncated { get; }
    /// <summary>
    /// 获取是否达到资源预算。
    /// </summary>
    public bool ResourceLimitExceeded { get; }
    /// <summary>
    /// 获取是否没有导入错误或资源失败。
    /// </summary>
    public bool IsSuccess => !ErrorsTruncated && !ResourceLimitExceeded && FailedRows == 0
        && ErrorCount == 0;
}
