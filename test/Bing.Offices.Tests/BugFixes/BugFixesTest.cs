using System;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Abstractions.Exports;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Tests.Models.Bugs;
using Bing.Utils.Json;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Tests.BugFixes
{
    /// <summary>
    /// Bug修复测试
    /// </summary>
    public class BugFixesTest
    {
        /// <summary>
        /// 输出
        /// </summary>
        protected ITestOutputHelper Output { get; }

        /// <summary>
        /// 导入提供程序
        /// </summary>
        protected IExcelImportProvider ImportProvider { get; }

        /// <summary>
        /// 导入服务
        /// </summary>
        protected IExcelImportService ImportService { get; }

        /// <summary>
        /// 导出提供程序
        /// </summary>
        protected IExcelExportProvider ExportProvider { get; }

        /// <summary>
        /// 导出服务
        /// </summary>
        protected IExcelExportService ExportService { get; }

        /// <summary>
        /// 当前目录
        /// </summary>
        protected string CurrentDir { get; }

        /// <summary>
        /// 初始化一个<see cref="BugFixesTest"/>类型的实例
        /// </summary>
        public BugFixesTest(ITestOutputHelper output)
        {
            Output = output;
            ImportProvider = new ExcelImportProvider();
            ImportService = new ExcelImportService(ImportProvider);
            ExportProvider = new ExcelExportProvider();
            ExportService = new ExcelExportService(ExportProvider);
            CurrentDir = Environment.CurrentDirectory;
        }

        /// <summary>
        /// Issue1 - decimal转换异常
        /// </summary>
        [Fact]
        public async Task Issue1_DecimalConvException()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue1.xlsx");
            var workbook = await ImportService.ImportAsync<Issue1>(new ImportOptions()
            {
                FileUrl = fileUrl
            });
            var result = workbook.GetResult<Issue1>();
            Output.WriteLine(result.ToJson());
        }
    }
}
