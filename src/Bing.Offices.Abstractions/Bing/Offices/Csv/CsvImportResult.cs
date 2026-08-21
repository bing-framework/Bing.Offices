namespace Bing.Offices.Csv;

/// <summary>
/// CSV 导入结果。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
public sealed class CsvImportResult<T>
{
    /// <summary>
    /// 初始化一个<see cref="CsvImportResult{T}"/>类型的实例。
    /// </summary>
    /// <param name="items">成功导入的实体集合。</param>
    /// <param name="errors">导入错误集合。</param>
    public CsvImportResult(IReadOnlyList<T> items, IReadOnlyList<CsvImportError> errors)
    {
        Items = items ?? throw new ArgumentNullException(nameof(items));
        Errors = errors ?? throw new ArgumentNullException(nameof(errors));
    }

    /// <summary>
    /// 获取成功导入的实体集合。
    /// </summary>
    public IReadOnlyList<T> Items { get; }

    /// <summary>
    /// 获取导入错误集合。
    /// </summary>
    public IReadOnlyList<CsvImportError> Errors { get; }
}
