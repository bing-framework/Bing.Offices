using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 MiniExcel 的 Provider 集成边界；跨 Provider 公共语义由 ProviderContract.Tests 覆盖。
/// </summary>
public sealed class MiniExcelProviderContractTest
{
    /// <summary>
    /// 验证不支持的功能会按约定在预检阶段失败且不产生部分输出。
    /// </summary>
    [Fact]
    public void UnsupportedFeatureContract_ShouldFailFastWithoutOutput()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new RegistrationRow() }, sheet =>
                sheet.SheetStyle(new Bing.Offices.Styles.ExcelCellStyle { Bold = true })));
        using var npoiOutput = new System.IO.MemoryStream();
        using var miniOutput = new System.IO.MemoryStream();

        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, npoiOutput);
        var exception = Assert.Throws<Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelExporter().Export(request, miniOutput));
        Assert.Equal(Bing.Offices.Exceptions.BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(Bing.Offices.Exceptions.BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("MiniExcel", exception.Provider);
        Assert.Equal(Bing.Offices.Exceptions.BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, miniOutput.Length);
        Assert.True(npoiOutput.Length > 0);
    }

    /// <summary>
    /// 验证旧版实体导入 Provider 未实现资源限制扩展时不会被隐式升级。
    /// </summary>
    [Fact]
    public void EntityResourceOptions_ShouldRequireOptInProviderExtension()
    {
        var layout = ExcelEntity.Layout<RegistrationRow>(builder =>
            builder.Cell("Data", "A1", row => row.Code));
        using var source = new MemoryStream(new byte[] { 1, 2, 3 });

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelImporter().ImportEntity(source, layout,
                new ExcelEntityImportOptions(new ExcelResourceLimits { MaxInputBytes = 3 })));

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("MiniExcel", exception.Provider);
        Assert.Contains("资源限制实体导入扩展接口", exception.Message);
    }

    /// <summary>
    /// 验证 Provider 注册顺序保留首个注册的实现。
    /// </summary>
    [Fact]
    public void ProviderRegistrationOrder_ShouldPreserveFirstRegistration()
    {
        using var npoiFirst = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        using var miniExcelFirst = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();

        Assert.IsType<Bing.Offices.Exports.NpoiExcelExporter>(
            npoiFirst.GetRequiredService<IExcelExporter>());
        Assert.IsType<Bing.Offices.Imports.NpoiExcelImporter>(
            npoiFirst.GetRequiredService<IExcelImporter>());
        Assert.IsType<MiniExcelExcelExporter>(
            miniExcelFirst.GetRequiredService<IExcelExporter>());
        Assert.IsType<Bing.Offices.Imports.MiniExcelExcelImporter>(
            miniExcelFirst.GetRequiredService<IExcelImporter>());
    }

    /// <summary>
    /// 验证首个注册的 Provider 可以完成公开异步往返流程。
    /// </summary>
    [Fact]
    public async Task ProviderRegistrationOrder_ShouldRunPublicRoundTripWithFirstProvider()
    {
        using var npoiFirst = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        using var miniExcelFirst = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();

        await AssertDiRoundTripAsync(npoiFirst, "NPOI");
        await AssertDiRoundTripAsync(miniExcelFirst, "MiniExcel");
    }

    /// <summary>
    /// 断言依赖注入注册的异步往返结果。
    /// </summary>
    /// <param name="provider">用于解析导入和导出服务的服务提供程序。</param>
    /// <param name="expectedProvider">预期的提供程序名称。</param>
    private static async Task AssertDiRoundTripAsync(IServiceProvider provider, string expectedProvider)
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration { PropertyName = nameof(RegistrationRow.Code), Title = "Code" },
                new ExcelColumnConfiguration { PropertyName = nameof(RegistrationRow.Count), Title = "Count" }
            }
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RegistrationRow { Code = expectedProvider, Count = 7 } },
            sheet => sheet.Mapping(mapping)));
        var importRequest = ExcelImport.Workbook<RegistrationWorkbook>(workbook =>
            workbook.Sheet<RegistrationRow>("Data", root => root.Rows,
                sheet => sheet.Mapping(mapping)));
        using var stream = new System.IO.MemoryStream();
        await provider.GetRequiredService<IExcelExporter>().ExportAsync(exportRequest, stream);
        stream.Position = 0;
        var result = await provider.GetRequiredService<IExcelImporter>().ImportAsync(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var row = Assert.Single(result.Workbook.Rows);
        Assert.Equal(expectedProvider, row.Code);
        Assert.Equal(7, row.Count);
    }

    /// <summary>
    /// DI 注册顺序测试使用的 Workbook。
    /// </summary>
    private sealed class RegistrationWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<RegistrationRow> Rows { get; } = new();
    }

    /// <summary>
    /// DI 注册顺序测试使用的一行数据。
    /// </summary>
    private sealed class RegistrationRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }
}
