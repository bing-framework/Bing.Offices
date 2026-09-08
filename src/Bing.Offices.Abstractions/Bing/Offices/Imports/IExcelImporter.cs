namespace Bing.Offices.Imports;

/// <summary>
/// Excel 流式导入器。
/// </summary>
public interface IExcelImporter
{
    /// <summary>
    /// 从一个 Workbook 中按各 Sheet 独立计划导入根模型。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 类型。</typeparam>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new();

    /// <summary>
    /// 从一个 Workbook 异步读取文件/流 IO，并按各 Sheet 独立计划导入根模型。
    /// NPOI Workbook DOM 的解析阶段仍为同步阶段。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 类型。</typeparam>
    /// <param name="source">调用方拥有的输入流。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步导入结果。</returns>
    Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new();

}
