using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Tests.Models;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Tests
{
    /// <summary>
    /// Excel导出测试
    /// </summary>
    public class ExcelExportTest : TestBase
    {
        /// <summary>
        /// Excel导出提供程序
        /// </summary>
        private readonly IExcelExportProvider _excelExportProvider;

        /// <summary>
        /// Excel导出服务
        /// </summary>
        private readonly IExcelExportService _excelExportService;

        /// <summary>
        /// 初始化
        /// </summary>
        public ExcelExportTest(ITestOutputHelper output) : base(output)
        {
            _excelExportProvider = new ExcelExportProvider();
            _excelExportService = new ExcelExportService(_excelExportProvider);
        }

        [Fact(DisplayName = "数据注解导出测试")]
        public async Task Test_Export_DataAnnotations()
        {
            var filePath = GetTestFilePath($"{nameof(Test_Export_DataAnnotations)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataAnnotations>();
            var result = await _excelExportService.ExportAsync(new ExportOptions<ExportTestDataAnnotations>
            {
                Data = data
            });
            await File.WriteAllBytesAsync(filePath, result);
        }

    }
}
