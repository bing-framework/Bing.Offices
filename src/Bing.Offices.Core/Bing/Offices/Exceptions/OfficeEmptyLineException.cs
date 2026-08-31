namespace Bing.Offices.Exceptions;

/// <summary>
/// 表示导入数据中出现不允许的空行。
/// </summary>
[Serializable]
public class OfficeEmptyLineException : OfficeException
{
    /// <summary>
    /// 空行对应的零基行索引。
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 使用错误消息初始化 <see cref="OfficeEmptyLineException"/> 实例。
    /// </summary>
    /// <param name="message">描述空行错误的异常消息。</param>
    public OfficeEmptyLineException(string message) : base(message) { }

    /// <summary>
    /// 使用错误消息和内部异常初始化 <see cref="OfficeEmptyLineException"/> 实例。
    /// </summary>
    /// <param name="message">描述空行错误的异常消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    public OfficeEmptyLineException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// 使用错误消息和行索引初始化 <see cref="OfficeEmptyLineException"/> 实例。
    /// </summary>
    /// <param name="message">描述空行错误的异常消息。</param>
    /// <param name="rowIndex">空行对应的零基行索引。</param>
    public OfficeEmptyLineException(string message, int rowIndex) : base(message)
    {
        RowIndex = rowIndex;
    }
         
}