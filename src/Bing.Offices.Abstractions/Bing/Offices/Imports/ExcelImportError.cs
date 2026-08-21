namespace Bing.Offices.Imports;

/// <summary>
/// Excel 导入期间发现的单元格错误。
/// </summary>
public sealed class ExcelImportError
{
    /// <summary>
    /// 初始化一个<see cref="ExcelImportError"/>类型的实例。
    /// </summary>
    /// <param name="code">错误代码。</param>
    /// <param name="message">错误消息。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">从一开始的行号。</param>
    /// <param name="columnIndex">从一开始的列号。</param>
    /// <param name="propertyName">目标属性名称。</param>
    /// <param name="columnKey">稳定列键。</param>
    /// <param name="header">实际表头。</param>
    /// <param name="rawValue">原始单元格值。</param>
    public ExcelImportError(ExcelImportErrorCode code, string message, string sheetName, int rowIndex,
        int columnIndex, string propertyName, string columnKey = null, string header = null, object rawValue = null)
    {
        Code = code;
        Message = message;
        SheetName = sheetName;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
        ColumnKey = columnKey ?? propertyName;
        Header = header;
        RawValue = rawValue;
    }

    /// <summary>
    /// 获取错误代码。
    /// </summary>
    public ExcelImportErrorCode Code { get; }

    /// <summary>
    /// 获取错误消息。
    /// </summary>
    public string Message { get; }

    /// <summary>
    /// 获取工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取从一开始的行号。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// 获取从一开始的列号。
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// 获取目标属性名称。
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// 获取稳定列键；未配置独立列键时回退为属性名称。
    /// </summary>
    public string ColumnKey { get; }

    /// <summary>
    /// 获取实际表头文本。
    /// </summary>
    public string Header { get; }

    /// <summary>
    /// 获取未转换的单元格值。
    /// </summary>
    public object RawValue { get; }
}
