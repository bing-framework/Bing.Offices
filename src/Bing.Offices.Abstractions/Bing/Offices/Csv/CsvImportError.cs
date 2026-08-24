namespace Bing.Offices.Csv;

/// <summary>
/// CSV 导入错误。
/// </summary>
public sealed class CsvImportError
{
    /// <summary>
    /// 初始化一个<see cref="CsvImportError"/>类型的实例。
    /// </summary>
    /// <param name="message">错误信息。</param>
    /// <param name="rowIndex">从一开始的行号。</param>
    /// <param name="columnIndex">从一开始的列号。</param>
    /// <param name="propertyName">属性名称。</param>
    /// <param name="firstRowNumber">重复值首次出现的行号。</param>
    public CsvImportError(string message, int rowIndex, int columnIndex, string propertyName,
        int? firstRowNumber = null)
    {
        Message = message ?? throw new ArgumentNullException(nameof(message));
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
        FirstRowNumber = firstRowNumber;
    }

    /// <summary>
    /// 获取错误信息。
    /// </summary>
    public string Message { get; }

    /// <summary>
    /// 获取从一开始的行号。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// 获取从一开始的列号。
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// 获取属性名称。
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// 获取重复值首次出现的行号；非重复错误为 null。
    /// </summary>
    public int? FirstRowNumber { get; }
}
