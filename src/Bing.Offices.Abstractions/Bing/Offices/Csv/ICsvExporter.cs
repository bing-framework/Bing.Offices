namespace Bing.Offices.Csv;

/// <summary>
/// CSV 流式导出器。
/// </summary>
public interface ICsvExporter
{
    /// <summary>
    /// 将实体集合写入调用方拥有的目标流。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="data">实体集合。</param>
    /// <param name="destination">目标流。调用完成后保持打开。</param>
    /// <param name="options">CSV 导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    void Export<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();

    /// <summary>
    /// 将实体集合写入调用方拥有的目标流。
    /// </summary>
    /// <remarks>
    /// 使用真正异步的 Writer/Stream IO 执行导出。
    /// </remarks>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="data">实体集合。</param>
    /// <param name="destination">目标流。调用完成后保持打开。</param>
    /// <param name="options">CSV 导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    Task ExportAsync<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();

    /// <summary>
    /// 将实体集合以原子方式写入 CSV 文件。
    /// </summary>
    /// <remarks>文件提交异常也在 exporter 观察边界内分发。</remarks>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="data">待导出的实体集合。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="options">CSV 导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    void ExportToFile<T>(IEnumerable<T> data, string path, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();

    /// <summary>
    /// 将实体集合以原子方式写入 CSV 文件。
    /// </summary>
    /// <remarks>
    /// 使用异步文件/流 IO 执行外围写入。
    /// </remarks>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="data">待导出的实体集合。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="options">CSV 导出选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    Task ExportToFileAsync<T>(IEnumerable<T> data, string path, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();
}
