using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Data;
using Bing.Offices.Testing.Requests;
using Bing.Offices.Testing.Models;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证公共输入资源限制在三个 Provider 上均 fail fast。
/// </summary>
public sealed class ResourceLimitContractTest
{
    /// <summary>
    /// 输入字节上限超出时应抛出结构化资源异常且不提交结果。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("MiniExcel")]
    [InlineData("ClosedXML")]
    public async Task MaxInputBytes_ShouldRejectBeforeMaterialization(string provider)
    {
        using var generated = new MemoryStream();
        await ProviderDrivers.Get(provider).CreateExporter().ExportAsync(
            ContractRequests.ScalarExport(), generated);
        generated.Position = 0;
        var request = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 1 })
            .Sheet<ScalarContractRow>("Data", root => root.Rows));

        using var source = new MemoryStream(generated.ToArray(), writable: false);
        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal("Import", exception.Operation.ToString());
    }

    /// <summary>
    /// 最大行数应返回资源错误且不提交超出预算的行。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task MaxRows_ShouldReturnStructuredResourceError(string provider)
    {
        using var source = await ExportAsync(provider, ContractRequests.ScalarExport());
        var request = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<ScalarContractRow>("Data", root => root.Rows));

        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.True(result.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit),
            string.Join("; ", result.Errors.Select(error => $"{error.Code}:{error.Message}:{error.RawValue}")));
    }

    /// <summary>
    /// 所有 Provider 的行预算必须跨 Sheet 共享，超限时不得返回前一 Sheet 的部分实体。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="synchronous">是否使用同步导入入口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("MiniExcel", false)]
    [InlineData("MiniExcel", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task MaxRowsAcrossSheets_ShouldNotExposePartialWorkbook(string provider, bool synchronous)
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("One", new[] { new ScalarContractRow { Code = "A", Count = 1 } })
            .AddSheet("Two", new[] { new ScalarContractRow { Code = "B", Count = 2 } }));
        using var source = await ExportAsync(provider, export);
        var request = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<ScalarContractRow>("One", root => root.Rows)
            .Sheet<ScalarContractRow>("Two", root => root.Rows));

        var importer = ProviderDrivers.Get(provider).CreateImporter();
        var result = synchronous
            ? importer.Import(source, request)
            : await importer.ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 所有 Provider 在两个 Sheet 的总行数恰好等于预算时应返回完整工作簿。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="synchronous">是否使用同步导入入口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("MiniExcel", false)]
    [InlineData("MiniExcel", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task MaxRowsAcrossSheets_ExactlyAtLimit_ShouldReturnCompleteWorkbook(
        string provider, bool synchronous)
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("One", new[] { new ScalarContractRow { Code = "A", Count = 1 } })
            .AddSheet("Two", new[] { new ScalarContractRow { Code = "B", Count = 2 } }));
        using var source = await ExportAsync(provider, export);
        var request = ExcelImport.Workbook<TwoSheetRowsWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 2 })
            .Sheet<ScalarContractRow>("One", root => root.OneRows)
            .Sheet<ScalarContractRow>("Two", root => root.TwoRows));

        var importer = ProviderDrivers.Get(provider).CreateImporter();
        var result = synchronous
            ? importer.Import(source, request)
            : await importer.ImportAsync(source, request);

        Assert.True(result.IsSuccess,
            string.Join("; ", result.Errors.Select(error => $"{error.Code}:{error.Message}:{error.RawValue}")));
        Assert.Empty(result.Errors);
        var first = Assert.Single(result.Workbook.OneRows);
        Assert.Equal("A", first.Code);
        Assert.Equal(1, first.Count);
        var second = Assert.Single(result.Workbook.TwoRows);
        Assert.Equal("B", second.Code);
        Assert.Equal(2, second.Count);
    }

    /// <summary>
    /// 工作表数量、列数量和物理单元格数量限制均应在 DOM 前结构化拒绝。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task WorkbookShapeLimits_ShouldRejectBeforeMaterialization(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        using var sheetsSource = await ExportAsync(provider, ExcelExport.Workbook(workbook => workbook
            .AddSheet("One", ContractData.ScalarRows())
            .AddSheet("Two", ContractData.ScalarRows())));
        var sheetsRequest = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<ScalarContractRow>("One", root => root.Rows));
        await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            driver.CreateImporter().ImportAsync(sheetsSource, sheetsRequest));

        using var cellsSource = await ExportAsync(provider, ContractRequests.ScalarExport());
        var columnsRequest = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 2 })
            .Sheet<ScalarContractRow>("Data", root => root.Rows));
        await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            driver.CreateImporter().ImportAsync(cellsSource, columnsRequest));

        using var cellsSource2 = await ExportAsync(provider, ContractRequests.ScalarExport());
        var cellsRequest = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxCells = 1 })
            .Sheet<ScalarContractRow>("Data", root => root.Rows));
        await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            driver.CreateImporter().ImportAsync(cellsSource2, cellsRequest));
    }

    /// <summary>
    /// 最大错误数应截断共享错误集合而不追加第 N+1 条错误。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task MaxErrors_ShouldTruncateStructuredErrors(string provider)
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", new[]
        {
            new ValidationContractRow { Code = string.Empty },
            new ValidationContractRow { Code = string.Empty }
        }));
        using var source = await ExportAsync(provider, export);
        var request = ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxErrors = 1 })
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Errors);
        Assert.True(result.ErrorsTruncated);
        Assert.Equal(1, result.MaxErrors);
    }

    /// <summary>
    /// 唯一值跟踪上限应返回资源错误并保留已收集的结构化结果。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task MaxTrackedUniqueValues_ShouldReturnStructuredResourceError(string provider)
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Codes", new[]
        {
            new UniqueContractRow { Code = "A" },
            new UniqueContractRow { Code = "B" }
        }));
        using var source = await ExportAsync(provider, export);
        var request = ExcelImport.Workbook<UniqueContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxTrackedUniqueValues = 1 })
            .Sheet<UniqueContractRow>("Codes", root => root.Rows));

        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(result.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit),
            string.Join("; ", result.Errors.Select(error => $"{error.Code}:{error.Message}:{error.RawValue}")));
        Assert.NotNull(result.Workbook);
    }

    /// <summary>
    /// 使用当前 Provider 生成真实 XLSX 输入并回到流起点。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="request">工作簿操作请求。</param>
    /// <returns>包含导出 XLSX 内容且位于起始位置的内存流。</returns>
    private static async Task<MemoryStream> ExportAsync(string provider, ExcelWorkbookExportRequest request)
    {
        var source = new MemoryStream();
        await ProviderDrivers.Get(provider).CreateExporter().ExportAsync(request, source);
        source.Position = 0;
        return source;
    }

    /// <summary>
    /// 为两个工作表提供互相独立的目标集合。
    /// </summary>
    private sealed class TwoSheetRowsWorkbook
    {
        /// <summary>
        /// 获取第一个工作表的导入数据。
        /// </summary>
        public List<ScalarContractRow> OneRows { get; } = new();
        /// <summary>
        /// 获取第二个工作表的导入数据。
        /// </summary>
        public List<ScalarContractRow> TwoRows { get; } = new();
    }

    /// <summary>
    /// 获取参与契约测试的 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });
}
