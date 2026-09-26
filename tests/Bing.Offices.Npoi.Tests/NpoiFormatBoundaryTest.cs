using System;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.Npoi.Tests;

/// <summary>
/// NPOI 对新增格式枚举的显式兼容边界测试。
/// </summary>
public sealed class NpoiFormatBoundaryTest
{
    /// <summary>
    /// 验证创建工作簿时拒绝 XLSB 格式。
    /// </summary>
    [Fact]
    public void PrepareWorkbook_Xlsb_ShouldBeRejectedWithoutWorkbook()
    {
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsb));

        Assert.Equal("NPOI", exception.Provider);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证同步导出在写入前拒绝 XLSB 格式。
    /// </summary>
    [Fact]
    public void Export_Xlsb_ShouldBeRejectedBeforeWriting()
    {
        using var destination = new MemoryStream();
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().Export(CreateXlsbRequest(), destination));

        Assert.Equal("NPOI", exception.Provider);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证异步导出拒绝 XLSB 且不写入目标流。
    /// </summary>
    [Fact]
    public async Task ExportAsync_Xlsb_ShouldBeRejectedBeforeStaging()
    {
        using var destination = new MemoryStream();
        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().ExportAsync(CreateXlsbRequest(), destination));

        Assert.Equal("NPOI", exception.Provider);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证异步导出 XLSB 失败时保留目标文件内容。
    /// </summary>
    [Fact]
    public async Task ExportToFileAsync_Xlsb_ShouldLeaveExistingTargetUntouched()
    {
        var path = Path.Combine(Path.GetTempPath(), "bing-offices-xlsb-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.WriteAllText(path, "sentinel");
        try
        {
            await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                new NpoiExcelExporter().ExportToFileAsync(CreateXlsbRequest(), path));
            Assert.Equal("sentinel", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 创建请求导出 XLSB 的测试工作簿定义。
    /// </summary>
    /// <returns>使用 XLSB 格式的导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateXlsbRequest()
        => ExcelExport.Workbook(workbook => workbook
            .Format(ExcelFormat.Xlsb)
            .AddSheet("Data", new[] { new Row { Name = "Alice" } }));

    /// <summary>
    /// 用于导入或导出测试的数据行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置测试行名称。
        /// </summary>
        public string Name { get; set; }
    }
}
