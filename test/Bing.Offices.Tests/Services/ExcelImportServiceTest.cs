using System;
using Bing.Offices.Abstractions.Imports;
using Xunit.Abstractions;

namespace Bing.Offices.Tests.Services
{
    /// <summary>
    /// Excel导入服务测试
    /// </summary>
    public class ExcelImportServiceTest : TestBase, IDisposable
    {
        /// <summary>
        /// Excel导入服务
        /// </summary>
        private readonly IExcelImportService _excelImportService;

        /// <summary>
        /// 初始化一个<see cref="TestBase"/>类型的实例
        /// </summary>
        public ExcelImportServiceTest(ITestOutputHelper output) : base(output)
        {
            _excelImportService = _excelImportService.Resolve();
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
        }
    }
}
