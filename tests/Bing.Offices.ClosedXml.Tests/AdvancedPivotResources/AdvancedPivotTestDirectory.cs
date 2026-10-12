using System;
using System.IO;

namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;

/// <summary>
/// 创建高级透视资源测试使用的隔离目录。
/// </summary>
internal sealed class AdvancedPivotTestDirectory : IDisposable
{
    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotTestDirectory" /> 类型的实例。
    /// </summary>
    public AdvancedPivotTestDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "bing-advanced-pivot-resource-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    /// <summary>
    /// 获取隔离目录的绝对路径。
    /// </summary>
    public string Path { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        if (Directory.Exists(Path))
            Directory.Delete(Path, recursive: true);
    }
}
