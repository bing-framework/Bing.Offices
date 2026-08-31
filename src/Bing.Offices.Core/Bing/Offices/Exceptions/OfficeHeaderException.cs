namespace Bing.Offices.Exceptions;

/// <summary>
/// 表示导入工作表表头缺少必需列或无法匹配。
/// </summary>
[Serializable]
public class OfficeHeaderException : OfficeException
{
    /// <summary>
    /// 表头对应的零基行索引。
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 相关列对应的零基列索引。
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 使用错误消息初始化 <see cref="OfficeHeaderException"/> 实例。
    /// </summary>
    /// <param name="message">描述表头匹配失败的异常消息。</param>
    public OfficeHeaderException(string message) : base(message) { }

    /// <summary>
    /// 使用错误消息和内部异常初始化 <see cref="OfficeHeaderException"/> 实例。
    /// </summary>
    /// <param name="message">描述表头匹配失败的异常消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    public OfficeHeaderException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// 使用错误消息和行索引初始化 <see cref="OfficeHeaderException"/> 实例。
    /// </summary>
    /// <param name="message">描述表头匹配失败的异常消息。</param>
    /// <param name="rowIndex">表头对应的零基行索引。</param>
    public OfficeHeaderException(string message, int rowIndex) : base(message)
    {
        RowIndex = rowIndex;
    }

    /// <summary>
    /// 使用错误消息、行索引和列索引初始化 <see cref="OfficeHeaderException"/> 实例。
    /// </summary>
    /// <param name="message">描述表头匹配失败的异常消息。</param>
    /// <param name="rowIndex">表头对应的零基行索引。</param>
    /// <param name="columnIndex">相关列对应的零基列索引。</param>
    public OfficeHeaderException(string message, int rowIndex, int columnIndex) : base(message)
    {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }
}
