namespace Bing.Offices.Imports;

/// <summary>
/// 多工作表分批导入完成摘要。
/// </summary>
public sealed class ExcelWorkbookBatchImportSummary
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookBatchImportSummary" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按源工作簿物理顺序排列的工作表统计。</param>
    public ExcelWorkbookBatchImportSummary(IReadOnlyList<ExcelBatchImportSheetSummary> sheets)
    {
        if (sheets == null)
            throw new ArgumentNullException(nameof(sheets));
        Sheets = sheets.ToArray();
    }

    /// <summary>
    /// 获取按源工作簿物理顺序排列的工作表统计。
    /// </summary>
    public IReadOnlyList<ExcelBatchImportSheetSummary> Sheets { get; }

    /// <summary>
    /// 获取所有工作表读取的正文行数。
    /// </summary>
    public int RowsRead => Sheets.Sum(sheet => sheet.Summary.RowsRead);

    /// <summary>
    /// 获取所有工作表成功实体行数。
    /// </summary>
    public int SucceededRows => Sheets.Sum(sheet => sheet.Summary.SucceededRows);

    /// <summary>
    /// 获取所有工作表失败行数。
    /// </summary>
    public int FailedRows => Sheets.Sum(sheet => sheet.Summary.FailedRows);

    /// <summary>
    /// 获取所有工作表收集的错误数量。
    /// </summary>
    public int ErrorCount => Sheets.Sum(sheet => sheet.Summary.ErrorCount);

    /// <summary>
    /// 获取任一工作表是否达到错误预算。
    /// </summary>
    public bool ErrorsTruncated => Sheets.Any(sheet => sheet.Summary.ErrorsTruncated);

    /// <summary>
    /// 获取任一工作表是否达到资源预算。
    /// </summary>
    public bool ResourceLimitExceeded => Sheets.Any(sheet => sheet.Summary.ResourceLimitExceeded);

    /// <summary>
    /// 获取所有工作表是否都成功完成导入。
    /// </summary>
    public bool IsSuccess => Sheets.All(sheet => sheet.Summary.IsSuccess);
}
