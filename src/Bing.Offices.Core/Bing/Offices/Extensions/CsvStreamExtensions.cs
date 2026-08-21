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
    /// <remarks>
    /// 该兼容入口不提供真实异步 I/O，仅委托当前 Stream-first 导出器。
    /// </remarks>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="exporter">CSV 导出器。</param>
    /// <param name="data">实体集合。</param>
    /// <param name="options">导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    [Obsolete("请使用 ExportToBytes；主导出契约为同步 Stream-first API。")]
    public static Task<byte[]> ExportToBytesAsync<T>(this ICsvExporter exporter, IEnumerable<T> data,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default) where T : class, new() =>
        Task.FromResult(exporter.ExportToBytes(data, options, cancellationToken));

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
        using var destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        exporter.Export(data, destination, options, cancellationToken);
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
}
