using System;

namespace Bing.Offices.Exceptions;

/// <summary>
/// Office表头缺列异常
/// </summary>
[Serializable]
public class OfficeHeaderException : OfficeException
{
    /// <summary>
    /// 行索引
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 列索引
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 初始化一个<see cref="OfficeHeaderException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    public OfficeHeaderException(string message) : base(message) { }

    /// <summary>
    /// 初始化一个<see cref="OfficeHeaderException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    /// <param name="innerException">错误来源</param>
    public OfficeHeaderException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// 初始化一个<see cref="OfficeHeaderException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    /// <param name="rowIndex">行索引</param>
    public OfficeHeaderException(string message, int rowIndex) : base(message)
    {
        RowIndex = rowIndex;
    }

    /// <summary>
    /// 初始化一个<see cref="OfficeHeaderException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    /// <param name="rowIndex">行索引</param>
    /// <param name="columnIndex">列索引</param>
    public OfficeHeaderException(string message, int rowIndex, int columnIndex) : base(message)
    {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }
}