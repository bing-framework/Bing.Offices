using System;

namespace Bing.Offices.Exceptions;

/// <summary>
/// Office数据转换异常
/// </summary>
[Serializable]
public class OfficeDataConvertException : OfficeException
{
    /// <summary>
    /// 原始类型
    /// </summary>
    public Type PrimitiveType { get; set; }

    /// <summary>
    /// 目标类型
    /// </summary>
    public Type TargetType { get; set; }

    /// <summary>
    /// 行索引
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 列索引
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 值
    /// </summary>
    public object Value { get; set; }

    /// <summary>
    /// 初始化一个<see cref="OfficeDataConvertException"/>类型的实例
    /// </summary>
    /// <param name="message">序列化信息</param>
    /// <param name="innerException">错误来源</param>
    public OfficeDataConvertException(string message, Exception innerException) : base(message, innerException) { }
}