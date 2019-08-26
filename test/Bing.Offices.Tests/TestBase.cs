using System.Linq;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Attributes;
using Bing.Offices.Npoi.Imports;
using Bing.Utils.Extensions;
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
        private IExcelImportProvider _excelImportProvider;

        public TestBase(ITestOutputHelper output)
        {
            Output = output;
            _excelImportProvider = new ExcelImportProvider();
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
