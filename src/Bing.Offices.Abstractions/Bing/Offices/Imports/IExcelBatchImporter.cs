using System.Globalization;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// 按单个工作表分批导入实体的公共接口。
/// </summary>
public interface IExcelBatchImporter
{
    /// <summary>
    /// 同步读取指定工作表，并按批次交付已转换实体。
    /// </summary>
    /// <typeparam name="TItem">工作表实体类型。</typeparam>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="request">分批导入请求。</param>
    /// <param name="onBatch">每个批次的同步回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>导入完成摘要。</returns>
    ExcelBatchImportSummary ImportBatches<TItem>(Stream source,
        ExcelBatchImportRequest<TItem> request, Action<ExcelImportBatch<TItem>> onBatch,
        CancellationToken cancellationToken = default)
        where TItem : class, new();

    /// <summary>
    /// 异步读取指定工作表，并按批次交付已转换实体。
    /// </summary>
    /// <typeparam name="TItem">工作表实体类型。</typeparam>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="request">分批导入请求。</param>
    /// <param name="onBatch">每个批次的异步回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>导入完成摘要任务。</returns>
    Task<ExcelBatchImportSummary> ImportBatchesAsync<TItem>(Stream source,
        ExcelBatchImportRequest<TItem> request,
        Func<ExcelImportBatch<TItem>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default)
        where TItem : class, new();
}

/// <summary>
/// 单个分批导入请求。
/// </summary>
/// <typeparam name="TItem">工作表实体类型。</typeparam>
public sealed class ExcelBatchImportRequest<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelBatchImportRequest{TItem}"/> 类型的实例。
    /// </summary>
    /// <param name="sheetName">要读取的工作表名称。</param>
    public ExcelBatchImportRequest(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(sheetName));
        SheetName = sheetName;
    }

    /// <summary>
    /// 获取工作表名称。
    /// </summary>
    public string SheetName { get; }
    /// <summary>
    /// 获取或初始化表头行的零基索引。
    /// </summary>
    public int HeaderRowIndex { get; init; }
    /// <summary>
    /// 获取或初始化数据起始行的零基索引。
    /// </summary>
    public int DataRowStartIndex { get; init; } = 1;
    /// <summary>
    /// 获取或初始化单批最大实体数量。
    /// </summary>
    public int BatchSize { get; init; } = 1000;
    /// <summary>
    /// 获取或初始化是否要求映射列全部存在。
    /// </summary>
    public bool RequireExpectedHeaders { get; init; } = true;
    /// <summary>
    /// 获取或初始化校验失败处理模式。
    /// </summary>
    public ExcelValidationFailureMode ValidationFailureMode { get; init; } = ExcelValidationFailureMode.StopOnFirstFailure;
    /// <summary>
    /// 获取或初始化映射方向的校验模式。
    /// </summary>
    public ExcelImportValidationMode ValidationMode { get; init; } = ExcelImportValidationMode.ConfiguredRules;
    /// <summary>
    /// 获取或初始化转换区域性。
    /// </summary>
    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;
    /// <summary>
    /// 获取或初始化请求级映射配置。
    /// </summary>
    public ExcelMappingConfiguration MappingConfiguration { get; init; }
    /// <summary>
    /// 获取或初始化规范化映射文档。
    /// </summary>
    public ExcelMappingDocument MappingDocument { get; init; }
    /// <summary>
    /// 获取或初始化导入资源限制。
    /// </summary>
    public ExcelResourceLimits ResourceLimits { get; init; } = new ExcelResourceLimits();
    /// <summary>
    /// 获取或初始化表头比较策略。
    /// </summary>
    public ExcelNameComparison HeaderComparison { get; init; } = ExcelNameComparison.OrdinalIgnoreCase;
    /// <summary>
    /// 获取或初始化表头空白处理策略。
    /// </summary>
    public ExcelWhitespacePolicy HeaderWhitespace { get; init; } = ExcelWhitespacePolicy.Trim;
    /// <summary>
    /// 获取或初始化正文空白处理策略。
    /// </summary>
    public ExcelWhitespacePolicy BodyWhitespace { get; init; } = ExcelWhitespacePolicy.Trim;

    /// <summary>
    /// 验证请求参数。
    /// </summary>
    public void Validate()
    {
        if (HeaderRowIndex < 0 || DataRowStartIndex <= HeaderRowIndex)
            throw new ArgumentOutOfRangeException(nameof(DataRowStartIndex));
        if (BatchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(BatchSize));
        if (!Enum.IsDefined(typeof(ExcelValidationFailureMode), ValidationFailureMode))
            throw new ArgumentOutOfRangeException(nameof(ValidationFailureMode));
        if (!Enum.IsDefined(typeof(ExcelImportValidationMode), ValidationMode))
            throw new ArgumentOutOfRangeException(nameof(ValidationMode));
        if (Culture == null)
            throw new ArgumentNullException(nameof(Culture));
        ResourceLimits?.Validate();
    }
}

/// <summary>
/// 已完成的一批导入结果。
/// </summary>
/// <typeparam name="TItem">工作表实体类型。</typeparam>
public sealed class ExcelImportBatch<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelImportBatch{TItem}"/> 类型的实例。
    /// </summary>
    /// <param name="batchNumber">从一开始的批次号。</param>
    /// <param name="firstRowIndex">批次第一行的一基物理行号。</param>
    /// <param name="items">成功实体。</param>
    /// <param name="errors">结构化错误。</param>
    public ExcelImportBatch(int batchNumber, int firstRowIndex, IReadOnlyList<TItem> items,
        IReadOnlyList<ExcelImportError> errors)
    {
        BatchNumber = batchNumber;
        FirstRowIndex = firstRowIndex;
        Items = items;
        Errors = errors;
    }

    /// <summary>
    /// 获取从一开始的批次号。
    /// </summary>
    public int BatchNumber { get; }
    /// <summary>
    /// 获取批次第一行的一基物理行号。
    /// </summary>
    public int FirstRowIndex { get; }
    /// <summary>
    /// 获取本批成功实体。
    /// </summary>
    public IReadOnlyList<TItem> Items { get; }
    /// <summary>
    /// 获取本批结构化错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }
}

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
