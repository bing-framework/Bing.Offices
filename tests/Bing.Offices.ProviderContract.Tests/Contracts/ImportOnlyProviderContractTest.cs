using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Models;
using Bing.Offices.Testing.Requests;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证只读 Provider 使用独立导入驱动器，不要求或伪造导出器。
/// </summary>
public sealed class ImportOnlyProviderContractTest
{
    /// <summary>
    /// 验证只读 Provider 的完整导入和分批导入。
    /// </summary>
    [Fact]
    public async Task ExcelDataReader_ShouldRunCompleteAndBatchImportContracts()
    {
        var driver = ProviderDrivers.GetImportOnly("ExcelDataReader");
        Assert.Equal("ExcelDataReader", driver.Name);
        Assert.Empty(((Bing.Offices.Providers.IExcelProviderCapabilityDescriptor)
            driver.CreateImporter()).WriteFormats);

        using var source = CreateSource();
        var result = await driver.CreateImporter().ImportAsync(source, ContractRequests.ScalarImport());

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "B" }, result.Workbook.Rows.Select(row => row.Code));

        source.Position = 0;
        var batches = new List<ExcelImportBatch<BatchScalarContractRow>>();
        var summary = driver.CreateBatchImporter().ImportBatches(source,
            new ExcelBatchImportRequest<BatchScalarContractRow>("Data") { BatchSize = 1 }, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(2, batches.SelectMany(batch => batch.Items).Count());
    }

    /// <summary>
    /// 创建只读 Provider 契约使用的标量工作簿。
    /// </summary>
    /// <returns>包含测试工作簿内容且位于起始位置的内存流。</returns>
    private static MemoryStream CreateSource()
    {
        var source = new MemoryStream();
        new NpoiExcelExporter().Export(ContractRequests.ScalarExport(), source);
        source.Position = 0;
        return source;
    }

    /// <summary>
    /// 用于分批导入标量数据的测试模型。
    /// </summary>
    private sealed class BatchScalarContractRow
    {
        /// <summary>
        /// 获取或设置测试行编码。
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 获取或设置测试数量。
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        /// 获取或设置测试金额。
        /// </summary>
        public decimal Amount { get; set; }
        /// <summary>
        /// 获取或设置测试启用状态。
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// 获取或设置测试分类。
        /// </summary>
        public ContractKind Kind { get; set; }
        /// <summary>
        /// 获取或设置测试日期。
        /// </summary>
        public System.DateTime Date { get; set; }
        /// <summary>
        /// 获取或设置可空的测试整数。
        /// </summary>
        public int? Optional { get; set; }
    }
}
