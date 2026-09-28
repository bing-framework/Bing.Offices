namespace Bing.Offices.Imports;

/// <summary>
/// 指示图片资源超出导入限制。
/// </summary>
internal sealed class ImageResourceLimitException : InvalidOperationException
{
    /// <summary>
    /// 初始化一个 <see cref="ImageResourceLimitException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述超出图片资源限制的消息。</param>
    internal ImageResourceLimitException(string message) : base(message)
    {
    }
}
