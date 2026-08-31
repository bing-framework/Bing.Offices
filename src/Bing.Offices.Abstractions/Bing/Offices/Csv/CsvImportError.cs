namespace Bing.Offices.Csv;

/// <summary>CSV 导入错误分类。</summary>
public enum CsvImportErrorCode
{
    /// <summary>输入格式无效。</summary>
    InvalidInput,
    /// <summary>表头或列结构无效。</summary>
    InvalidHeader,
    /// <summary>值转换失败。</summary>
    ValueConversion,
    /// <summary>业务校验失败。</summary>
    Validation,
    /// <summary>超过资源限制。</summary>
    ResourceLimit
}

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
    /// <param name="code">错误分类。</param>
    public CsvImportError(string message, int rowIndex, int columnIndex, string propertyName,
        int? firstRowNumber = null, CsvImportErrorCode code = CsvImportErrorCode.InvalidInput)
    {
        Message = message ?? throw new ArgumentNullException(nameof(message));
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        PropertyName = propertyName;
        FirstRowNumber = firstRowNumber;
        Code = code;
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

    /// <summary>获取错误分类。</summary>
    public CsvImportErrorCode Code { get; }
}
