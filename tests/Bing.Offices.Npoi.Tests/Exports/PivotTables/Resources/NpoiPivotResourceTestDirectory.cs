using System;
using System.IO;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Resources;

/// <summary>
/// 创建隔离的 NPOI 透视资源测试目录。
/// </summary>
internal sealed class NpoiPivotResourceTestDirectory : IDisposable
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotResourceTestDirectory" /> 类型的实例。
    /// </summary>
    public NpoiPivotResourceTestDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "bing-npoi-pivot-resource-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    /// <summary>
    /// 获取本测试独占的临时目录绝对路径。
    /// </summary>
    public string Path { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        if (Directory.Exists(Path))
            Directory.Delete(Path, recursive: true);
    }
}
