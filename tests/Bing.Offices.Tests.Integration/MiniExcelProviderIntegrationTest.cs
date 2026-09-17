using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// 使用真实文件路径验证 MiniExcel Provider 的包消费者边界。
/// </summary>
public sealed class MiniExcelProviderIntegrationTest
{
    [Fact]
    public async Task ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripRealXlsx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-miniexcel-{Guid.NewGuid():N}.xlsx");
        using var provider = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new IntegrationRow { Name = "真实文件", Count = 9 } }));
        var importRequest = ExcelImport.Workbook<IntegrationWorkbook>(workbook => workbook
            .Sheet<IntegrationRow>("Data", root => root.Rows));

        try
        {
            await provider.GetRequiredService<IExcelExporter>().ExportToFileAsync(request, path);
            var result = await ExcelStreamExtensions.ImportFromFileAsync<IntegrationWorkbook>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);

            var row = Assert.Single(result.Workbook.Rows);
            Assert.Empty(result.Errors);
            Assert.Equal("真实文件", row.Name);
            Assert.Equal(9, row.Count);
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    private sealed class IntegrationWorkbook
    {
        public List<IntegrationRow> Rows { get; } = new();
    }

    private sealed class IntegrationRow
    {
        public string Name { get; set; }
        public int Count { get; set; }
    }
}
