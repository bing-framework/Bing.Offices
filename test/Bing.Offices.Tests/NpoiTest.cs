using System;
using Bing.Offices.Npoi.Extensions;
using Xunit;

namespace Bing.Offices.Tests
{
    public class NpoiTest
    {
        /// <summary>
        /// 测试 - 冻结首行
        /// </summary>
        [Fact]
        public void Test_Freeze_FirstRow()
        {
            var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            var sheet = workbook.CreateSheet("Test");
            sheet.CreateFreezePane(0, 1, 0, 1);
            var bytes = workbook.SaveToBuffer();
            Bing.IO.FileHelper.Write($"D:\\test_freeze_first_row_{DateTime.Now.Ticks}.xlsx", bytes);
        }

        /// <summary>
        /// 测试 - 冻结首列
        /// </summary>
        [Fact]
        public void Test_Freeze_FirstColumn()
        {
            var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            var sheet = workbook.CreateSheet("Test");
            sheet.CreateFreezePane(1, 0, 1, 0);
            var bytes = workbook.SaveToBuffer();
            Bing.IO.FileHelper.Write($"D:\\test_freeze_first_column_{DateTime.Now.Ticks}.xlsx", bytes);
        }

        /// <summary>
        /// 测试 - 冻结首行首列
        /// </summary>
        [Fact]
        public void Test_Freeze_FirstRowWithColumn()
        {
            var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            var sheet = workbook.CreateSheet("Test");
            sheet.CreateFreezePane(1, 1, 1, 1);
            var bytes = workbook.SaveToBuffer();
            Bing.IO.FileHelper.Write($"D:\\test_freeze_first_row_with_column_{DateTime.Now.Ticks}.xlsx", bytes);
        }

        /// <summary>
        /// 测试 - 冻结多个
        /// </summary>
        [Fact]
        public void Test_Freeze_Multiple()
        {
            var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            var sheet = workbook.CreateSheet("Test");
            sheet.CreateFreezePane(4, 3, 4, 3);
            var bytes = workbook.SaveToBuffer();
            Bing.IO.FileHelper.Write($"D:\\test_freeze_multiple_{DateTime.Now.Ticks}.xlsx", bytes);
        }
    }
}
