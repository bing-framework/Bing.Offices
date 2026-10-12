namespace Bing.Offices.Imports;

/// <summary>
/// 工作簿分批导入交付的单个工作表批次。
/// </summary>
public sealed class ExcelWorkbookImportBatch
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookImportBatch" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">来源工作表名称。</param>
    /// <param name="itemType">批次实体类型。</param>
    /// <param name="batchNumber">整个 Workbook 中从一开始的批次号。</param>
    /// <param name="firstRowIndex">批次第一行的一基物理行号。</param>
    /// <param name="items">成功实体。</param>
    /// <param name="errors">结构化错误。</param>
    public ExcelWorkbookImportBatch(string sheetName, Type itemType, int batchNumber,
        int firstRowIndex, IReadOnlyList<object> items, IReadOnlyList<ExcelImportError> errors)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(sheetName));
        SheetName = sheetName;
        ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
        BatchNumber = batchNumber;
        FirstRowIndex = firstRowIndex;
        Items = items ?? Array.Empty<object>();
        Errors = errors ?? Array.Empty<ExcelImportError>();
    }

    /// <summary>
    /// 获取来源工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取批次实体类型。
    /// </summary>
    public Type ItemType { get; }

    /// <summary>
    /// 获取整个工作簿中从一开始的批次号。
    /// </summary>
    public int BatchNumber { get; }

    /// <summary>
    /// 获取批次第一行的一基物理行号。
    /// </summary>
    public int FirstRowIndex { get; }

    /// <summary>
    /// 获取本批成功实体。
    /// </summary>
    public IReadOnlyList<object> Items { get; }

    /// <summary>
    /// 获取本批结构化错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }
}
