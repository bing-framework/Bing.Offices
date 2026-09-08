using Bing.Offices.Csv;

namespace Bing.Offices.Extensions;

/// <summary>
/// CSV 流式导入导出的便利扩展。
/// </summary>
public static class CsvStreamExtensions
{
    /// <summary>
    /// 将实体集合导出为 CSV 字节数组。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="exporter">CSV 导出器。</param>
    /// <param name="data">实体集合。</param>
    /// <param name="options">导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public static byte[] ExportToBytes<T>(this ICsvExporter exporter, IEnumerable<T> data,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        using var destination = new MemoryStream();
        exporter.Export(data, destination, options, cancellationToken);
        return destination.ToArray();
    }

    /// <summary>
    /// 将实体集合导出为 CSV 文件。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="exporter">CSV 导出器。</param>
    /// <param name="data">实体集合。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="options">导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public static void ExportToFile<T>(this ICsvExporter exporter, IEnumerable<T> data, string path,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        exporter.ExportToFile(data, path, options, cancellationToken);
    }

    /// <summary>异步将实体集合导出为 CSV 字节数组。</summary>
    public static async Task<byte[]> ExportToBytesAsync<T>(this ICsvExporter exporter, IEnumerable<T> data,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        using var destination = new MemoryStream();
        await exporter.ExportAsync(data, destination, options, cancellationToken).ConfigureAwait(false);
        return destination.ToArray();
    }

    /// <summary>异步将实体集合导出为 CSV 文件。</summary>
    public static Task ExportToFileAsync<T>(this ICsvExporter exporter, IEnumerable<T> data, string path,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        return exporter.ExportToFileAsync(data, path, options, cancellationToken);
    }

    /// <summary>
    /// 从 CSV 字节数组导入实体集合。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="importer">CSV 导入器。</param>
    /// <param name="content">CSV 文件字节数组。</param>
    /// <param name="options">导入选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public static CsvImportResult<T> ImportFromBytes<T>(this ICsvImporter importer, byte[] content,
        CsvImportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (content == null)
            throw new ArgumentNullException(nameof(content));
        using var source = new MemoryStream(content, writable: false);
        return importer.Import(source, options, cancellationToken);
    }

    /// <summary>
    /// 从 CSV 文件导入实体集合。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="importer">CSV 导入器。</param>
    /// <param name="path">源文件路径。</param>
    /// <param name="options">导入选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public static CsvImportResult<T> ImportFromFile<T>(this ICsvImporter importer, string path,
        CsvImportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("源文件路径不能为空。", nameof(path));
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return importer.Import(source, options, cancellationToken);
    }

    /// <summary>异步从 CSV 字节数组导入实体集合。</summary>
    public static Task<CsvImportResult<T>> ImportFromBytesAsync<T>(this ICsvImporter importer, byte[] content,
        CsvImportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (content == null)
            throw new ArgumentNullException(nameof(content));
        return ImportFromBytesAsyncCore(importer, content, options, cancellationToken);
    }

    private static async Task<CsvImportResult<T>> ImportFromBytesAsyncCore<T>(ICsvImporter importer, byte[] content,
        CsvImportOptions<T> options, CancellationToken cancellationToken) where T : class, new()
    {
        using var source = new MemoryStream(content, writable: false);
        return await importer.ImportAsync(source, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>异步从 CSV 文件导入实体集合。</summary>
    public static async Task<CsvImportResult<T>> ImportFromFileAsync<T>(this ICsvImporter importer, string path,
        CsvImportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("源文件路径不能为空。", nameof(path));
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        return await importer.ImportAsync(source, options, cancellationToken).ConfigureAwait(false);
    }
}
