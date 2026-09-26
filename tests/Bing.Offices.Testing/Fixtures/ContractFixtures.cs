using System;
using System.IO;

namespace Bing.Offices.Testing.Fixtures;

/// <summary>
/// 为真实文件合同提供隔离的临时目录。
/// </summary>
public sealed class TestTempDirectory : IDisposable
{
    /// <summary>
    /// 初始化一个 <see cref="TestTempDirectory"/> 类型的实例。
    /// </summary>
    public TestTempDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "bing-offices-contract-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    /// <summary>
    /// 获取临时目录路径。
    /// </summary>
    public string Path { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        if (Directory.Exists(Path))
            Directory.Delete(Path, recursive: true);
    }
}

/// <summary>
/// 记录一次输出提交前后的目标文件状态。
/// </summary>
public sealed class TargetFileState
{
    /// <summary>
    /// 获取或初始化目标文件路径。
    /// </summary>
    public string Path { get; init; } = string.Empty;
    /// <summary>
    /// 获取或初始化提交前的字节内容。
    /// </summary>
    public byte[] Before { get; init; } = Array.Empty<byte>();
    /// <summary>
    /// 获取或初始化提交后的字节内容。
    /// </summary>
    public byte[] After { get; init; } = Array.Empty<byte>();
}
