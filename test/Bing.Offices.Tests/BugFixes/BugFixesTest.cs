﻿using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Tests.Models.Bugs;
using Bing.Extensions;
using Bing.Offices.Tests.Models.Purchase;
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
                FileUrl = fileUrl,
            });
            Assert.Throws<OfficeDataConvertException>(() =>
            {
                try
                {
                    var result = workbook.GetResult<Issue1>();
                }
                catch (OfficeDataConvertException e)
                {
                    Assert.Equal(5,e.RowIndex);
                    Assert.Equal(1, e.ColumnIndex);
                    Assert.Equal("值",e.Name);
                    Output.WriteLine(e.ToJson());
                    throw;
                }
            });
        }

        /// <summary>
        /// Issue2 - 初始化单元格太多
        /// </summary>
        [Fact]
        public async Task Issue2_InitCelSoMuch()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue2.xlsx");

            await Assert.ThrowsAsync<OfficeHeaderException>(async () =>
            {
                var workbook = await ImportService.ImportAsync<Issue2>(new ImportOptions()
                {
                    FileUrl = fileUrl,
                    SheetIndex = 0,
                    DataRowIndex = 3,
                    HeaderRowIndex = 2,
                });
                var result = workbook.GetResult<Issue2>();
                Output.WriteLine(result.ToJson());
            });

        }

        /// <summary>
        /// Issue3 - 异常转换错误
        /// </summary>
        [Fact]
        public async Task Issue3_ExceptionConvertError()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue3.xlsx");
            var workbook = await ImportService.ImportAsync<Issue3>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 3,
                HeaderRowIndex = 2,
            });
            Assert.Throws<OfficeDataConvertException>(() =>
            {
                try
                {
                    var result = workbook.GetResult<Issue3>();
                }
                catch (OfficeDataConvertException e)
                {
                    Assert.Equal(5, e.RowIndex);
                    Assert.Equal(2, e.ColumnIndex);
                    Assert.Equal("调拨数量", e.Name);
                    Output.WriteLine(e.ToJson());
                    throw;
                }
            });
        }

        /// <summary>
        /// Issue4 - 校验失败返回真实单元格位置
        /// </summary>
        [Fact]
        public async Task Issue4_ValidateFail_ReturnRealLocation()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue4.xlsx");
            var workbook = await ImportService.ImportAsync<Issue4>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 3,
                HeaderRowIndex = 2,
            });
            var validateResult = workbook.Validate().ToList();
            Assert.True(validateResult.Any());
            Assert.Equal(5, validateResult[0].RowIndex);
            Output.WriteLine(validateResult.ToJson());
        }

        /// <summary>
        /// Issue5 - 校验失败返回真实单元格位置并不指定标题行索引
        /// </summary>
        [Fact]
        public async Task Issue5_ValidateFail_ReturnRealLocation_With_HeaderRowIndex()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue5.xlsx");
            var workbook = await ImportService.ImportAsync<Issue5>(new ImportOptions
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
            });
            var validateResult = workbook.Validate().ToList();
            Assert.True(validateResult.Any());
            Assert.Equal(4, validateResult[0].RowIndex);
            Output.WriteLine(validateResult.ToJson());
        }

        /// <summary>
        /// Issue6 - 导入忽略属性
        /// </summary>
        [Fact]
        public async Task Issue6_Import_IgnoreProperty()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue6.xlsx");
            var workbook = await ImportService.ImportAsync<Issue6>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
            });
            var result = workbook.GetResult<Issue6>();
            Output.WriteLine(result.ToJson());
        }

        /// <summary>
        /// Issue7 - 导入空行
        /// </summary>
        [Fact]
        public async Task Issue7_Import_EmptyLine()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue7.xlsx");
            await Assert.ThrowsAsync<OfficeEmptyLineException>(async() =>
            {
                try
                {
                    var workbook = await ImportService.ImportAsync<Issue7>(new ImportOptions()
                    {
                        FileUrl = fileUrl,
                        SheetIndex = 1,
                        DataRowIndex = 3,
                        HeaderRowIndex = 2,
                        EnabledEmptyLine = true
                    });
                    var result = workbook.GetResult<Issue7>();
                    Output.WriteLine(result.ToJson());
                }
                catch (OfficeEmptyLineException e)
                {
                    Output.WriteLine(e.ToJson());
                    throw;
                }
                
            });
        }

        /// <summary>
        /// Issue7 - 导入空行 -跳过
        /// </summary>
        [Fact]
        public async Task Issue7_Import_EmptyLine_Skip()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue7.xlsx");
            var workbook = await ImportService.ImportAsync<Issue7>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 1,
                DataRowIndex = 3,
                HeaderRowIndex = 2,
                EnabledEmptyLine = false
            });
            var result = workbook.GetResult<Issue7>();
            Output.WriteLine(result.ToJson());
        }

        /// <summary>
        /// Issue8 - 导入-时间转换
        /// </summary>
        [Fact]
        public async Task Issue8_Import_DateTimeParser()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue8.xlsx");
            var workbook = await ImportService.ImportAsync<Issue8>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
            });
            var result = workbook.GetResult<Issue8>();
            result.ForEach(x =>
            {
                Output.WriteLine(DateTime.Parse(x.CreateTime).ToString());
            });
            Output.WriteLine(result.ToJson());
        }

        /// <summary>
        /// Issue9 - 导入-清空左右两侧空格
        /// </summary>
        [Fact]
        public async Task Issue9_Import_Trim()
        {
            var fileUrl = Path.Combine(CurrentDir, "Resources/Bugs", "issue9.xlsx");
            var workbook = await ImportService.ImportAsync<Issue9>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
            });
            var result = workbook.GetResult<Issue9>();
            Output.WriteLine(result.ToJson());
        }


        #region Purchase  采购订单导入

        /// <summary>
        /// 采购订单导入模板_不含税单价  
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Purchase_Import_ExIncludeTax()
        {
            //预计到货日期 日期格式 有问题
            var fileUrl3= Path.Combine(CurrentDir, "Resources/Purchase", "采购订单导入模板_不含税单价3.xlsx");
            var workbook3 = await ImportService.ImportAsync<ImportPurchaseOrderExIncludeTax>(new ImportOptions()
            {
                FileUrl = fileUrl3,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
                EnabledEmptyLine = false,
                IgnoreEmptyLineAfterData = true,
            });
            var result3 = workbook3.GetResult<ImportPurchaseOrderExIncludeTax>();
            Output.WriteLine(result3.ToJson());

            //重复列名
            var fileUrl2 = Path.Combine(CurrentDir, "Resources/Purchase", "采购订单导入模板_不含税单价2.xlsx");
            var workbook2 = await ImportService.ImportAsync<ImportPurchaseOrderExIncludeTax>(new ImportOptions()
            {
                FileUrl = fileUrl2,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
                EnabledEmptyLine = false,
                IgnoreEmptyLineAfterData = true,
                HeaderColumnOnly = false,//重复列名报错
            });
            var result2 = workbook2.GetResult<ImportPurchaseOrderExIncludeTax>();
            Output.WriteLine(result2.ToJson());

            //不启用表头缓存
            var fileUrl = Path.Combine(CurrentDir, "Resources/Purchase", "采购订单导入模板_不含税单价.xlsx");
            var workbook = await ImportService.ImportAsync<ImportPurchaseOrderExIncludeTax>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
                EnabledEmptyLine = false,
                IgnoreEmptyLineAfterData = true,
                EnabledHeaderRowCache= false,
            });
            var result = workbook.GetResult<ImportPurchaseOrderExIncludeTax>();
            Output.WriteLine(result.ToJson());

            
        }

        /// <summary>
        /// 采购订单导入模板_含税单价  
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Purchase_Import_IncludeTax()
        {
            //预计到货日期 日期格式 有问题
            var fileUrl2 = Path.Combine(CurrentDir, "Resources/Purchase", "采购订单导入模板_含税单价2.xlsx");
            var workbook2 = await ImportService.ImportAsync<ImportPurchaseOrderIncludeTax>(new ImportOptions()
            {
                FileUrl = fileUrl2,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
                EnabledEmptyLine = false,
                IgnoreEmptyLineAfterData = true,
            });
            var result2 = workbook2.GetResult<ImportPurchaseOrderIncludeTax>();
            Output.WriteLine(result2.ToJson());

            var fileUrl = Path.Combine(CurrentDir, "Resources/Purchase", "采购订单导入模板_含税单价.xlsx");
            var workbook = await ImportService.ImportAsync<ImportPurchaseOrderIncludeTax>(new ImportOptions()
            {
                FileUrl = fileUrl,
                SheetIndex = 0,
                DataRowIndex = 1,
                HeaderRowIndex = 0,
                EnabledEmptyLine=false,
                IgnoreEmptyLineAfterData=true,
            });
            var result = workbook.GetResult<ImportPurchaseOrderIncludeTax>();
            Output.WriteLine(result.ToJson());
        }
        #endregion
    }
}
