namespace Bing.Offices.Imports;

/// <summary>
/// 按多个工作表分批导入 Workbook 模型的公共接口。
/// </summary>
public interface IExcelWorkbookBatchImporter
{
    /// <summary>
    /// 按源工作簿物理顺序读取已配置工作表，并交付已转换实体批次。
    /// </summary>
    /// <remarks>
    /// 通过同步回调交付批次。
    /// </remarks>
    /// <typeparam name="TWorkbook">承载各工作表配置的 Workbook 模型类型。</typeparam>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="onBatch">每个工作表批次的同步回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含各工作表和总计统计的导入摘要。</returns>
    ExcelWorkbookBatchImportSummary ImportWorkbookBatches<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request,
        Action<ExcelWorkbookImportBatch> onBatch,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new();

    /// <summary>
    /// 按源工作簿物理顺序读取已配置工作表，并交付已转换实体批次。
    /// </summary>
    /// <remarks>
    /// 通过异步回调交付批次。
    /// </remarks>
    /// <typeparam name="TWorkbook">承载各工作表配置的 Workbook 模型类型。</typeparam>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="onBatch">每个工作表批次的异步回调。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含各工作表和总计统计的导入摘要任务。</returns>
    Task<ExcelWorkbookBatchImportSummary> ImportWorkbookBatchesAsync<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new();
}
