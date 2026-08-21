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
    /// 将 Workbook 导出请求写入 Excel 文件。
    /// </summary>
    public static void ExportToFile(this IExcelExporter exporter, ExcelWorkbookExportRequest request, string path,
        CancellationToken cancellationToken = default)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        using var destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        exporter.Export(request, destination, cancellationToken);
    }

    /// <summary>
    /// 从 Excel 字节数组导入 Workbook。
    /// </summary>
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
}
