namespace Bing.Offices.Imports;

/// <summary>
/// 单个工作表的分批导入摘要。
/// </summary>
public sealed class ExcelBatchImportSheetSummary
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelBatchImportSheetSummary" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">请求的工作表名称。</param>
    /// <param name="itemType">该工作表实体类型。</param>
    /// <param name="summary">该工作表的分批统计。</param>
    public ExcelBatchImportSheetSummary(string sheetName, Type itemType, ExcelBatchImportSummary summary)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(sheetName));
        ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
        SheetName = sheetName;
        Summary = summary ?? throw new ArgumentNullException(nameof(summary));
    }

    /// <summary>
    /// 获取请求的工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取该工作表实体类型。
    /// </summary>
    public Type ItemType { get; }

    /// <summary>
    /// 获取该工作表的分批统计。
    /// </summary>
    public ExcelBatchImportSummary Summary { get; }
}
