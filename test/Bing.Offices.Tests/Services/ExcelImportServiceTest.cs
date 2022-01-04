using Bing.Offices.Imports;
using Bing.Offices.Npoi.Imports;
using Xunit.Abstractions;

namespace Bing.Offices.Tests.Services
{
    /// <summary>
    /// Excel导入服务测试
    /// </summary>
    public class ExcelImportServiceTest
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
        /// 初始化一个<see cref="ExcelIETest"/>类型的实例
        /// </summary>
        public ExcelImportServiceTest(ITestOutputHelper output)
        {
            Output = output;
            ImportProvider = new ExcelImportProvider();
            ImportService = new ExcelImportService(ImportProvider);
        }
    }
}
