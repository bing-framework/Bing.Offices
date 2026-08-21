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
}
