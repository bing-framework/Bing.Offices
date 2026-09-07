namespace Bing.Offices.Exports;

/// <summary>
/// Excel 流式导出器。
/// </summary>
public interface IExcelExporter
{
    /// <summary>
    /// 将 Workbook 请求写入目标流。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <param name="destination">调用方拥有的目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 将 Workbook 请求以原子方式写入文件；文件提交异常也在 exporter 观察边界内分发。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    void ExportToFile(ExcelWorkbookExportRequest request, string path,
        CancellationToken cancellationToken = default);

}
