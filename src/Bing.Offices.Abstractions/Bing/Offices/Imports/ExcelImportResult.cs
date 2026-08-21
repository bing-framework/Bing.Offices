using System.Linq;

namespace Bing.Offices.Imports;

/// <summary>
/// Excel 流式导入结果。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
internal sealed class ExcelImportResult<T>
{
    /// <summary>
    /// 初始化一个<see cref="ExcelImportResult{T}"/>类型的实例。
    /// </summary>
    /// <param name="items">成功导入的实体。</param>
    /// <param name="errors">导入错误。</param>
    public ExcelImportResult(IReadOnlyList<T> items, IReadOnlyList<ExcelImportError> errors)
    {
        Items = items;
        Errors = errors;
    }

    /// <summary>
    /// 获取成功导入的实体。
    /// </summary>
    public IReadOnlyList<T> Items { get; }

    /// <summary>
    /// 获取导入错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>
    /// 获取导入是否没有错误。
    /// </summary>
    public bool IsSuccess => Errors.Count == 0;

    /// <summary>
    /// 获取成功导入的行数。
    /// </summary>
    public int SuccessCount => Items.Count;

    /// <summary>
    /// 获取失败行数。一个行包含多个错误时只按一个失败行计数。
    /// </summary>
    public int FailureCount => Errors.Select(error => (error.SheetName, error.RowIndex)).Distinct().Count();

    /// <summary>
    /// 获取处理的行总数。
    /// </summary>
    public int TotalCount => SuccessCount + FailureCount;
}
