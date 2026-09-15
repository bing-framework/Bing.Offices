using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Extensions;

/// <summary>
/// Workbook Request 的文件和字节数组便利扩展。
/// </summary>
public static class ExcelStreamExtensions
{
    /// <summary>
    /// 将 Workbook 导出请求写为 Excel 字节数组。
    /// </summary>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="request">Workbook 导出请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含 Excel 内容的字节数组。</returns>
    public static byte[] ExportToBytes(this IExcelExporter exporter, ExcelWorkbookExportRequest request,
        CancellationToken cancellationToken = default)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        using var destination = new MemoryStream();
        exporter.Export(request, destination, cancellationToken);
        return destination.ToArray();
    }

    /// <summary>
    /// 异步将 Workbook 请求导出为 Excel 字节数组。
    /// </summary>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="request">Workbook 导出请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含 Excel 内容的字节数组任务。</returns>
    public static async Task<byte[]> ExportToBytesAsync(this IExcelExporter exporter,
        ExcelWorkbookExportRequest request, CancellationToken cancellationToken = default)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        using var destination = new MemoryStream();
        await exporter.ExportAsync(request, destination, cancellationToken).ConfigureAwait(false);
        return destination.ToArray();
    }

    /// <summary>
    /// 从 Excel 字节数组导入 Workbook。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="content">Excel 文件字节数组。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>导入结果。</returns>
    public static ExcelWorkbookImportResult<TWorkbook> ImportFromBytes<TWorkbook>(this IExcelImporter importer,
        byte[] content, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken = default) where TWorkbook : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (content == null)
            throw new ArgumentNullException(nameof(content));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        using var source = new MemoryStream(content, writable: false);
        return importer.Import(source, request, cancellationToken);
    }

    /// <summary>
    /// 从 Excel 文件导入 Workbook。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="path">源文件路径。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>导入结果。</returns>
    public static ExcelWorkbookImportResult<TWorkbook> ImportFromFile<TWorkbook>(this IExcelImporter importer,
        string path, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken = default) where TWorkbook : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("源文件路径不能为空。", nameof(path));
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return importer.Import(source, request, cancellationToken);
    }

    /// <summary>
    /// 异步从 Excel 字节数组导入 Workbook。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="content">Excel 文件字节数组。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步导入结果任务。</returns>
    public static Task<ExcelWorkbookImportResult<TWorkbook>> ImportFromBytesAsync<TWorkbook>(
        this IExcelImporter importer, byte[] content, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken = default) where TWorkbook : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (content == null)
            throw new ArgumentNullException(nameof(content));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        return ImportFromBytesAsyncCore(importer, content, request, cancellationToken);
    }

    /// <summary>通过内存流执行 Excel 字节数组的异步 Workbook 导入。</summary>
    /// <typeparam name="TWorkbook">Workbook 根实体类型。</typeparam>
    /// <param name="importer">执行导入的 Excel 导入器。</param>
    /// <param name="content">Excel 文件内容。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">导入过程中检查的取消令牌。</param>
    /// <returns>包含 Workbook 实例、错误和截断状态的导入结果。</returns>
    private static async Task<ExcelWorkbookImportResult<TWorkbook>> ImportFromBytesAsyncCore<TWorkbook>(
        IExcelImporter importer, byte[] content, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        using var source = new MemoryStream(content, writable: false);
        return await importer.ImportAsync(source, request, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 异步从 Excel 文件导入 Workbook。
    /// </summary>
    /// <typeparam name="TWorkbook">根 Workbook 类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="path">源文件路径。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步导入结果任务。</returns>
    public static async Task<ExcelWorkbookImportResult<TWorkbook>> ImportFromFileAsync<TWorkbook>(
        this IExcelImporter importer, string path, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken = default) where TWorkbook : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("源文件路径不能为空。", nameof(path));
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        return await importer.ImportAsync(source, request, cancellationToken).ConfigureAwait(false);
    }
}
