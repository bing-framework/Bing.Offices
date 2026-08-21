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
    CsvImportResult<T> Import<T>(Stream source, CsvImportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new();
}
