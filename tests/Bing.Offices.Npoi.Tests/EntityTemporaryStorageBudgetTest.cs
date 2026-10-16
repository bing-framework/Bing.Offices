using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI XLSX 与 XLS 前向实体临时存储预算合同。
/// </summary>
public sealed class EntityTemporaryStorageBudgetTest : EntityTemporaryStorageBudgetContractTest
{
    protected override string ProviderName => "NPOI";
    protected override ServiceProvider CreateProvider() =>
        new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();

    /// <summary>
    /// XLS 投影及逐工作表 XML 均计入预算，成功、资源失败路径都清理中间文件。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Hssf_ShouldChargeConvertedPackageAndSheetFiles(bool asynchronous)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Note");
        for (var index = 1; index <= 200; index++)
        {
            var row = sheet.CreateRow(index);
            row.CreateCell(0).SetCellValue($"C{index}");
            row.CreateCell(1).SetCellValue(new string('中', 100));
        }
        using var serialized = new MemoryStream();
        workbook.Write(serialized, leaveOpen: true);
        var bytes = serialized.ToArray();
        var before = TemporaryFiles();
        using var provider = CreateProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using (var source = new Input(bytes, seekable: false))
        {
            var result = await Import(importer, source, Options(10_000_000), asynchronous, _ => { });
            Assert.True(result.IsSuccess);
            Assert.Equal(200, result.SucceededRows);
            Assert.True(source.CanRead);
        }
        Assert.Equal(before, TemporaryFiles());
        using (var source = new Input(bytes, seekable: false))
        {
            var calls = 0;
            var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                Import(importer, source, Options(bytes.LongLength + 1000), asynchronous, _ => calls++));
            AssertResourceException(exception);
            Assert.Equal(0, calls);
            Assert.True(source.CanRead);
        }
        Assert.Equal(before, TemporaryFiles());
    }
}
