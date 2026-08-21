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
    public CsvImportError(string message, int rowIndex, int columnIndex, string propertyName)
    {
        Message = message ?? throw new ArgumentNullException(nameof(message));
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
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
}
