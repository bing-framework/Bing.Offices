namespace Bing.Offices.Csv;

/// <summary>
/// CSV 流式导入器。
/// </summary>
public interface ICsvImporter
{
    /// <summary>
    /// 从调用方拥有的输入流读取实体集合。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="source">输入流。调用完成后保持打开。</param>
    /// <param name="options">CSV 导入选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>导入得到的实体集合及逐行错误信息。</returns>
    CsvImportResult<T> Import<T>(Stream source, CsvImportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();

    /// <summary>
    /// 从调用方拥有的输入流读取实体集合。
    /// </summary>
    /// <remarks>
    /// 使用真正异步的 Reader/Stream IO 执行导入。
    /// </remarks>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="source">输入流。调用完成后保持打开。</param>
    /// <param name="options">CSV 导入选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步导入结果。</returns>
    Task<CsvImportResult<T>> ImportAsync<T>(Stream source, CsvImportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();
}
