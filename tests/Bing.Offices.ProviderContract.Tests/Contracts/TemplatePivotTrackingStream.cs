using System.IO;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 记录测试期间是否释放了调用方拥有的内存流。
/// </summary>
internal sealed class TemplatePivotTrackingStream : MemoryStream
{
    /// <summary>
    /// 初始化一个 <see cref="TemplatePivotTrackingStream" /> 类型的实例。
    /// </summary>
    /// <param name="buffer">用于初始化流内容的字节。</param>
    public TemplatePivotTrackingStream(byte[] buffer) : base(buffer, writable: true)
    {
    }

    /// <summary>
    /// 获取流是否曾被释放。
    /// </summary>
    public bool WasDisposed { get; private set; }

    /// <inheritdoc />
    /// <remarks>记录测试流释放状态后继续执行基类释放。</remarks>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
            WasDisposed = true;
        base.Dispose(disposing);
    }
}
