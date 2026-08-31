using System.Runtime.Serialization;

namespace Bing.Offices.Exceptions;

/// <summary>
/// 表示 Bing.Offices 处理 Office 文档时发生的业务异常。
/// </summary>
[Serializable]
public class OfficeException : Exception
{
    /// <summary>
    /// 初始化一个使用默认消息的 <see cref="OfficeException"/> 实例。
    /// </summary>
    public OfficeException() : base("Office异常") { }

    /// <summary>
    /// 使用格式化消息初始化 <see cref="OfficeException"/> 实例。
    /// </summary>
    /// <param name="msgFormat">消息格式字符串。</param>
    /// <param name="objects">用于填充消息格式的参数。</param>
    public OfficeException(string msgFormat, params object[] objects) : base(string.Format(msgFormat, objects)) { }

    /// <summary>
    /// 使用消息和内部异常初始化 <see cref="OfficeException"/> 实例。
    /// </summary>
    /// <param name="message">异常消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    public OfficeException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// 使用序列化数据恢复 <see cref="OfficeException"/> 实例。
    /// </summary>
    /// <param name="info">异常序列化数据。</param>
    /// <param name="context">序列化流上下文。</param>
    public OfficeException(SerializationInfo info, StreamingContext context) : base(info, context) { }
}
