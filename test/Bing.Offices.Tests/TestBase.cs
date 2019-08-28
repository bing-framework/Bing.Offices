using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Abstractions.Exports;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Utils.Extensions;
using Bing.Utils.Json;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Tests
{
    /// <summary>
    /// 测试基类
    /// </summary>
    public class TestBase
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
        /// 初始化一个<see cref="TestBase"/>类型的实例
        /// </summary>
        public TestBase(ITestOutputHelper output)
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
        public void Test_Import()
        {
            var result = _excelImportProvider.Convert<Barcode>("D:\\导入国标码_导入格式.xlsx");
            foreach (var sheet in result.Sheets)
            {
                foreach (var header in sheet.GetHeader())
                {
                    Output.WriteLine(header.Cells.Select(x => x.Value.ToString()).Join());
                }

                foreach (var body in sheet.GetBody())
                {
                    Output.WriteLine(body.Cells.Select(x => x.Value.ToString()).Join());
                }
            }
        }

        /// <summary>
        /// 测试 - 服务导入
        /// </summary>
        [Fact]
        public async Task Test_Import_1()
        {
            var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions()
            {
                FileUrl = "D:\\导入国标码_导入格式.xlsx",
            });
            var result = workbook.GetResult<Barcode>();
            Output.WriteLine(result.Count().ToString());
            Output.WriteLine(result.ToJson());
        }
        
        /// <summary>
        /// 测试 - 导出
        /// </summary>
        [Fact]
        public async Task Test_Export()
        {
            var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions()
            {
                FileUrl = "D:\\导入国标码_导入格式.xlsx",
            });
            var result = workbook.GetResult<Barcode>();

            var bytes = await _excelExportService.ExportAsync(new ExportOptions<Barcode>()
            {
                Data = result.ToList()
            });
            await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
        }
    }

    /// <summary>
    /// 条形码
    /// </summary>
    public class Barcode
    {
        [ColumnName("系统编号")]
        public string Id { get; set; }

        [ColumnName("国标码")]
        public string Code { get; set; }
    }
}
