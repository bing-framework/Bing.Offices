using System;

namespace Bing.Offices.Exceptions;

/// <summary>
/// Office空行异常
/// </summary>
[Serializable]
public class OfficeEmptyLineException : OfficeException
{
    /// <summary>
    /// 行索引
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 初始化一个<see cref="OfficeEmptyLineException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    public OfficeEmptyLineException(string message) : base(message) { }

    /// <summary>
    /// 初始化一个<see cref="OfficeEmptyLineException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    /// <param name="innerException">错误来源</param>
    public OfficeEmptyLineException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// 初始化一个<see cref="OfficeEmptyLineException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    /// <param name="rowIndex">行索引</param>
    public OfficeEmptyLineException(string message, int rowIndex) : base(message)
    {
        RowIndex = rowIndex;
    }
         
}