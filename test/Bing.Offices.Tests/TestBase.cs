using System.Collections.Generic;
using System.IO;
using Xunit.Abstractions;

namespace Bing.Offices.Tests;

/// <summary>
/// 测试基类
/// </summary>
public class TestBase
{
    /// <summary>
    /// 输出
    /// </summary>
    protected ITestOutputHelper Output { get; }

    /// <summary>
    /// 初始化
    /// </summary>
    public TestBase(ITestOutputHelper output) => Output = output;

    /// <summary>
    /// 获取根目录
    /// </summary>
    public string GetTestRootPath() => Directory.GetCurrentDirectory();

    /// <summary>
    /// 获取测试文件路径根
    /// </summary>
    public string GetTestFilePath(params string[] paths)
    {
        var rootPath = GetTestRootPath();
        var list = new List<string>
        {
            rootPath
        };
        list.AddRange(paths);
        var result = Path.Combine(list.ToArray());
        Output.WriteLine($"文件路径：{result}");
        return result;
    }

    /// <summary>
    /// 删除文件
    /// </summary>
    public void DeleteFile(string file)
    {
        if(File.Exists(file))
            File.Delete(file);
    }
}