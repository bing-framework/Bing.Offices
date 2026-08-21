using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using Xunit;
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

    /// <summary>
    /// 断言导出的工作簿可重新打开，且首个工作表包含预期表头。
    /// </summary>
    /// <param name="bytes">工作簿字节数组。</param>
    /// <param name="expectedHeader">预期表头文本。</param>
    /// <param name="headerRowIndex">表头行索引。</param>
    /// <param name="headerColumnIndex">表头列索引。</param>
    protected void AssertExportedWorkbook(byte[] bytes, string expectedHeader, int headerRowIndex = 0, int headerColumnIndex = 0)
    {
        Assert.NotNull(bytes);
        Assert.NotEmpty(bytes);
        using var stream = new MemoryStream(bytes);
        using var workbook = WorkbookFactory.Create(stream);
        Assert.True(workbook.NumberOfSheets > 0);

        var row = workbook.GetSheetAt(0).GetRow(headerRowIndex);
        Assert.NotNull(row);
        var cell = row.GetCell(headerColumnIndex);
        Assert.NotNull(cell);
        Assert.Equal(expectedHeader, cell.StringCellValue);
    }

    /// <summary>
    /// 创建测试临时文件路径。
    /// </summary>
    /// <param name="fileName">文件名。</param>
    protected string GetTemporaryFilePath(string fileName)
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, fileName);
    }

    /// <summary>
    /// 删除测试临时文件及其空目录。
    /// </summary>
    /// <param name="filePath">临时文件路径。</param>
    protected void DeleteTemporaryFile(string filePath)
    {
        DeleteFile(filePath);
        var directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory) && !Directory.EnumerateFileSystemEntries(directory).Any())
            Directory.Delete(directory);
    }
}
