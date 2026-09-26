using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Models;
using Bing.Offices.Testing.Requests;
using ClosedXML.Excel;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 Failure Workbook 的 Supported/Unsupported 公共结果和输出副作用。
/// </summary>
public sealed class FailureWorkbookContractTest
{
    /// <summary>
    /// 支持的 Provider 写出失败工作簿，不支持的 Provider 应预检失败且不写出。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task AnnotatedOriginal_ShouldMatchProviderProfile(string provider)
    {
        using var source = new MemoryStream();
        await ProviderDrivers.Get(provider).CreateExporter().ExportAsync(
            ContractRequests.ValidationExport(), source);
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        if (provider == "MiniExcel")
        {
            var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request));
            Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
            Assert.Equal(0, destination.Length);
            return;
        }

        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);
        Assert.False(result.IsSuccess);
        Assert.NotEmpty(result.Errors);
        Assert.True(destination.Length > 0);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var data = workbook.GetSheet("Data");
        Assert.Equal("Code", data.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Quantity", data.GetRow(0).GetCell(1).StringCellValue);
        Assert.NotNull(data.GetRow(1).GetCell(0).CellComment);
        var summary = workbook.GetSheet("_ImportErrors");
        Assert.Equal("Code", summary.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Data", summary.GetRow(1).GetCell(2).StringCellValue);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 仅错误行输出必须保留源表头、错误行和统一错误摘要，而非只产生任意字节。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("MiniExcel")]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task ErrorRowsOnly_ShouldContainHeaderErrorRowAndSummary(string provider)
    {
        using var source = new MemoryStream();
        await ProviderDrivers.Get(provider).CreateExporter().ExportAsync(
            ContractRequests.ValidationExport(), source);
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        if (provider == "MiniExcel")
        {
            var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request));
            Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
            Assert.Empty(destination.ToArray());
            return;
        }

        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        destination.Position = 0;
        using var workbook = new XLWorkbook(destination);
        var data = workbook.Worksheet("Data");
        Assert.Equal("Code", data.Cell(1, 1).GetString());
        Assert.Equal("Quantity", data.Cell(1, 2).GetString());
        Assert.Equal(2, data.LastRowUsed().RowNumber());
        Assert.Equal("Data", data.Cell(2, 3).GetString());
        Assert.Equal(2, data.Cell(2, 4).GetValue<int>());
        var summary = workbook.Worksheet("_ImportErrors");
        Assert.Equal("Code", summary.Cell(1, 1).GetString());
        Assert.Equal("Data", summary.Cell(2, 3).GetString());
        Assert.Equal(2, summary.Cell(2, 4).GetValue<int>());
    }

    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });
}
