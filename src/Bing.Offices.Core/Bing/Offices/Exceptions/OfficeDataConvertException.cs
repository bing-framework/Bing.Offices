namespace Bing.Offices.Exceptions;

/// <summary>
/// 表示 Office 单元格或字段值无法转换为目标实体类型的异常。
/// </summary>
[Serializable]
public class OfficeDataConvertException : OfficeException
{
    /// <summary>
    /// 原始单元格值对应的运行时类型。
    /// </summary>
    public Type PrimitiveType { get; set; }

    /// <summary>
    /// 转换操作要求的目标类型。
    /// </summary>
    public Type TargetType { get; set; }

    /// <summary>
    /// 发生转换错误的零基行索引。
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 发生转换错误的零基列索引。
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 发生转换错误的属性或列名称。
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 转换失败的原始值。
    /// </summary>
    public object Value { get; set; }

    /// <summary>
    /// 使用转换错误消息和内部异常初始化 <see cref="OfficeDataConvertException"/> 实例。
    /// </summary>
    /// <param name="message">描述转换失败原因的异常消息。</param>
    /// <param name="innerException">导致转换失败的内部异常。</param>
    public OfficeDataConvertException(string message, Exception innerException) : base(message, innerException) { }
}