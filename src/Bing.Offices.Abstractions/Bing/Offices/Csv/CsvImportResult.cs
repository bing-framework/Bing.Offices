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
    /// <param name="isTruncated">是否因资源限制而提前截断。</param>
    /// <param name="maxErrors">触发截断的最大错误数。</param>
    public CsvImportResult(IReadOnlyList<T> items, IReadOnlyList<CsvImportError> errors,
        bool isTruncated = false, int? maxErrors = null)
    {
        Items = items ?? throw new ArgumentNullException(nameof(items));
        Errors = errors ?? throw new ArgumentNullException(nameof(errors));
        IsTruncated = isTruncated;
        MaxErrors = maxErrors;
    }

    /// <summary>
    /// 获取成功导入的实体集合。
    /// </summary>
    public IReadOnlyList<T> Items { get; }

    /// <summary>
    /// 获取导入错误集合。
    /// </summary>
    public IReadOnlyList<CsvImportError> Errors { get; }

    /// <summary>获取结果是否因资源限制而提前截断。</summary>
    public bool IsTruncated { get; }

    /// <summary>获取触发截断的最大错误数。</summary>
    public int? MaxErrors { get; }
}
