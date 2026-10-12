using System;
using System.IO;
using System.Threading.Tasks;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 保存前向导入使用的输入流及其所有权。
/// </summary>
internal sealed class NpoiForwardInput : IDisposable, IAsyncDisposable
{
    private readonly bool _ownsStream;

    /// <summary>
    /// 初始化一个 <see cref="NpoiForwardInput" /> 类型的实例。
    /// </summary>
    /// <param name="stream">前向导入输入流。</param>
    /// <param name="isHssf">输入是否为 OLE2 HSSF 工作簿。</param>
    /// <param name="ownsStream">是否由实例负责释放输入流。</param>
    internal NpoiForwardInput(Stream stream, bool isHssf, bool ownsStream)
    {
        Stream = stream ?? throw new ArgumentNullException(nameof(stream));
        IsHssf = isHssf;
        _ownsStream = ownsStream;
    }

    /// <summary>
    /// 获取输入流。
    /// </summary>
    internal Stream Stream { get; }

    /// <summary>
    /// 获取输入是否为 HSSF 工作簿。
    /// </summary>
    internal bool IsHssf { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_ownsStream)
            Stream.Dispose();
    }

    /// <inheritdoc />
    public ValueTask DisposeAsync()
    {
        if (_ownsStream)
            return Stream.DisposeAsync();
        return ValueTask.CompletedTask;
    }
}
