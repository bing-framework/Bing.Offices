using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Tests.Models;
using Bing.Utils.Json;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Tests;

/// <summary>
/// 商品导入测试
/// </summary>
public class GoodsImportTest
{
    /// <summary>
    /// 输出
    /// </summary>
    protected ITestOutputHelper Output;

    /// <summary>
    /// Excel导入提供程序
    /// </summary>
    private readonly IExcelImportProvider _excelImportProvider;

    /// <summary>
    /// Excel导入服务
    /// </summary>
    private readonly IExcelImportService _excelImportService;

    /// <summary>
    /// Excel导出提供程序
    /// </summary>
    private readonly IExcelExportProvider _excelExportProvider;

    /// <summary>
    /// Excel导出服务
    /// </summary>
    private readonly IExcelExportService _excelExportService;

    /// <summary>
    /// 初始化一个<see cref="GoodsImportTest"/>类型的实例
    /// </summary>
    public GoodsImportTest(ITestOutputHelper output)
    {
        Output = output;
        _excelImportProvider = new ExcelImportProvider();
        _excelImportService = new ExcelImportService(_excelImportProvider);
        _excelExportProvider = new ExcelExportProvider();
        _excelExportService = new ExcelExportService(_excelExportProvider);
    }

    /// <summary>
    /// 测试 - 导入
    /// </summary>
    [Fact]
    public async Task Test_Import()
    {
        await ImportAsync("D:\\进销存商品导入模板 (26).xlsx");
        await ImportAsync("D:\\进销存商品导入模板 (29).xlsx");
        await ImportAsync("D:\\进销存商品导入模板 (32).xlsx");
        await ImportAsync("D:\\进销存商品导入模板 (27).xlsx");
        await ImportAsync("D:\\进销存商品导入模板 (26).xlsx");
        await ImportAsync("D:\\进销存商品导入模板 (29).xlsx");
        await ImportAsync("D:\\进销存商品导入模板 (32).xlsx");
    }

    /// <summary>
    /// 导入
    /// </summary>
    /// <param name="path">文件路径</param>
    private async Task ImportAsync(string path)
    {
        var workbook = await _excelImportService.ImportAsync<ImportGoods>(new ImportOptions
        {
            FileUrl = path
        });
        var result = workbook.GetResult<ImportGoods>();
        Output.WriteLine(result.Count().ToString());
        Output.WriteLine(result.ToJson());
    }
}
