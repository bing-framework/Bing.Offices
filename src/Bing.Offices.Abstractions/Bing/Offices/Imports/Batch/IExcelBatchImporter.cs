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
