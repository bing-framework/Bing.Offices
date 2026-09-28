using System.Globalization;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// 已完成的一批导入结果。
/// </summary>
/// <typeparam name="TItem">工作表实体类型。</typeparam>
public sealed class ExcelImportBatch<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelImportBatch{TItem}"/> 类型的实例。
    /// </summary>
    /// <param name="batchNumber">从一开始的批次号。</param>
    /// <param name="firstRowIndex">批次第一行的一基物理行号。</param>
    /// <param name="items">成功实体。</param>
    /// <param name="errors">结构化错误。</param>
    public ExcelImportBatch(int batchNumber, int firstRowIndex, IReadOnlyList<TItem> items,
        IReadOnlyList<ExcelImportError> errors)
    {
        BatchNumber = batchNumber;
        FirstRowIndex = firstRowIndex;
        Items = items;
        Errors = errors;
    }

    /// <summary>
    /// 获取从一开始的批次号。
    /// </summary>
    public int BatchNumber { get; }
    /// <summary>
    /// 获取批次第一行的一基物理行号。
    /// </summary>
    public int FirstRowIndex { get; }
    /// <summary>
    /// 获取本批成功实体。
    /// </summary>
    public IReadOnlyList<TItem> Items { get; }
    /// <summary>
    /// 获取本批结构化错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }
}
