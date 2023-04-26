using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Tests.Models;
using Bing.Extensions;
using Bing.Helpers;
using Bing.Offices.Metadata.Excels;
using Bing.Utils.IdGenerators.Core;
using Bing.Utils.Json;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Tests;

/// <summary>
/// 测试基类
/// </summary>
// ReSharper disable once InconsistentNaming
public class ExcelIETest : TestBase
{
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
    /// 初始化一个<see cref="ExcelIETest"/>类型的实例
    /// </summary>
    public ExcelIETest(ITestOutputHelper output) : base(output)
    {
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
        var result = _excelImportProvider.Convert<Barcode>(GetTestFilePath("Resources", "导入国标码_导入格式.xlsx"));
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
    /// 测试 - 导入 动态标题
    /// </summary>
    [Fact]
    public void Test_Import_DynamicTitle_1()
    {
        var result = _excelImportProvider.Convert<Barcode>(GetTestFilePath("Resources", "导入国标码_导入格式_动态标题.xlsx"));
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
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式.xlsx"),
        });
        var result = workbook.GetResult<Barcode>().ToList();
        Output.WriteLine(result.Count.ToString());
        Output.WriteLine(result.ToJson());
        Assert.Equal(3, result.Count);
    }

    /// <summary>
    /// 测试 - 导入 动态标题
    /// </summary>
    [Fact]
    public async Task Test_Import_DynamicTitle_2()
    {
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式_动态标题.xlsx")
        });
        var result = workbook.GetResult<Barcode>();
        Output.WriteLine(result.Count().ToString());
        Output.WriteLine(result.ToJson());
    }

    /// <summary>
    /// 测试 - 导入 动态标题
    /// </summary>
    [Fact]
    public async Task Test_Import_DynamicTitle_More_1()
    {
        var workbook1 = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式_动态标题.xlsx")
        });
        var result1 = workbook1.GetResult<Barcode>();
        Output.WriteLine(result1.Count().ToString());
        Output.WriteLine(result1.ToJson());
        var workbook2 = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式_动态标题1.xlsx")
        });
        var result2 = workbook2.GetResult<Barcode>();
        Output.WriteLine(result2.Count().ToString());
        Output.WriteLine(result2.ToJson());
    }

    /// <summary>
    /// 测试 - 服务导入
    /// </summary>
    [Fact]
    public async Task Test_Import_Validate_1()
    {
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式.xlsx"),
        });
        var validateResult = workbook.Validate();
        if (validateResult.Any())
        {
            Output.WriteLine(validateResult.ToJson());
            return;
        }

        var result = workbook.GetResult<Barcode>();
        Output.WriteLine(result.Count().ToString());
        Output.WriteLine(result.ToJson());
    }

    /// <summary>
    /// 测试 - 导入 可空值
    /// </summary>
    /// <returns></returns>
    [Fact]
    public async Task Test_Import_NullableValue_1()
    {
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式.xlsx"),
        });
        var validateResult = workbook.Validate();
        if (validateResult.Any())
        {
            Output.WriteLine(validateResult.ToJson());
            return;
        }

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
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式.xlsx"),
        });
        var result = workbook.GetResult<Barcode>();

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Barcode>
        {
            Data = result.ToList()
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 动态标题
    /// </summary>
    [Fact]
    public async Task Test_Export_DynamicTitle_1()
    {
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式_动态标题1.xlsx")
        });
        var result = workbook.GetResult<Barcode>();

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Barcode>
        {
            Data = result.ToList(),
            DynamicColumns = new List<string> { "创建人", "更新时间", "更新人", "备注" }
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 多个
    /// </summary>
    [Fact]
    public async Task Test_Export_Multiple()
    {
        var workbook = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式_动态标题1.xlsx")
        });
        var result = workbook.GetResult<Barcode>();

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Barcode>
        {
            Data = result.ToList(),
            DynamicColumns = new List<string> { "创建人", "更新时间", "更新人", "备注" }
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_动态标题_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);

        var workbook1 = await _excelImportService.ImportAsync<Barcode>(new ImportOptions
        {
            FileUrl = GetTestFilePath("Resources", "导入国标码_导入格式.xlsx"),
        });
        var result1 = workbook1.GetResult<Barcode>();

        var bytes1 = await _excelExportService.ExportAsync(new ExportOptions<Barcode>
        {
            Data = result1.ToList()
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes1);
    }

    /// <summary>
    /// 测试 - 导出 忽略属性
    /// </summary>
    [Fact]
    public async Task Test_Export_IgnoreProperty()
    {
        var data = new List<ExportOrder>();
        for (int i = 0; i < 10000; i++)
        {
            DateTime? currentTime = null;
            if (i % 2 == 0)
                currentTime = DateTime.Now;
            data.Add(new ExportOrder
            {
                Id = $"A{i}",
                Name = $"测试名称+++++{i}",
                Index = i + 1,
                CreateTime = currentTime,
                IgnoreProperty = $"忽略属性+++++{i}",
                NotMappedProperty = $"忽略映射属性+++++{i}"
            });
        }

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<ExportOrder>()
        {
            Data = data,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 格式化属性
    /// </summary>
    [Fact]
    public async Task Test_Export_FormatProperty()
    {
        var data = new List<Bing.Offices.Tests.Models.ExportFormat>();
        for (var i = 0; i < 10; i++)
        {
            data.Add(new Bing.Offices.Tests.Models.ExportFormat()
            {
                Id = $"A{i}",
                Name = $"测试名称+++++{i}",
                Index = i + 1,
                IgnoreProperty = $"忽略属性+++++{i}",
                NotMappedProperty = $"忽略映射属性+++++{i}",
                Quantity = i,
                Price = i * 0.42m,
                Money = i * (i * 0.42444m),
                CreateTime = DateTime.Now.AddMinutes(i)
            });
        }

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Bing.Offices.Tests.Models.ExportFormat>()
        {
            Data = data,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }


    /// <summary>
    /// 测试 - 导出 额外表头
    /// </summary>
    [Fact]
    public async Task Test_Export_HeaderRow()
    {
        var data = new List<Bing.Offices.Tests.Models.ExportFormat>();
        for (int i = 0; i < 10000; i++)
        {
            data.Add(new Bing.Offices.Tests.Models.ExportFormat()
            {
                Id = $"A{i}",
                Name = $"测试名称+++++{i}",
                Index = i + 1,
                IgnoreProperty = $"忽略属性+++++{i}",
                NotMappedProperty = $"忽略映射属性+++++{i}",
                Money = i * 1000,
                CreateTime = DateTime.Now.AddMinutes(i)
            });
        }

        var rows = new List<Bing.Offices.Metadata.Excels.Row>();
        var row = new Bing.Offices.Metadata.Excels.Row(0);
        row.Add(new Cell("开心", 2, 2, 1));
        row.Add(new Cell("懂得", 4, 2, 1));
        rows.Add(row);

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Bing.Offices.Tests.Models.ExportFormat>()
        {
            HeaderRow = rows.ToList<IRow>(),
            HeaderRowIndex = 1,
            DataRowStartIndex = 2,
            Data = data,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 自定义小数位数
    /// </summary>
    [Fact]
    public async Task Test_Export_Custom_Scale()
    {
        var data = new List<ExportScale>();
        for (var i = 0; i < 100; i++)
        {
            data.Add(new ExportScale
            {
                Id = SnowflakeIdGenerator.Current.Create().ToString(),
                Byte = Conv.ToByte(i),
                NullableByte = i % 2 == 0 ? null : Conv.ToByteOrNull(i),
                Short = Conv.ToShort(i),
                NullableShort = i % 2 == 0 ? null : Conv.ToShortOrNull(i),
                Int = i,
                NullableInt = i % 2 == 0 ? null : Conv.ToIntOrNull(i),
                Long = i,
                NullableLong = i % 2 == 0 ? null : Conv.ToLongOrNull(i),
                Float = i * 0.4222222f,
                NullableFloat = i % 2 == 0 ? null : Conv.ToFloatOrNull(i) * 0.4222222f,
                Double = i * 0.4222222,
                NullableDouble = i % 2 == 0 ? null : Conv.ToDoubleOrNull(i) * 0.4222222,
                Decimal = i * 0.4222222m,
                NullableDecimal = i % 2 == 0 ? null : Conv.ToDecimalOrNull(i) * 0.4222222m,
            });
        }

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<ExportScale>()
        {
            Data = data,
            HeaderRowIndex = 0,
            DataRowStartIndex = 1,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 报表
    /// </summary>
    [Fact]
    public async Task Test_Export_Report()
    {
        var list = new List<Report>();
        list.Add(new Report
        {
            WayName = "自配送",
            Extend = new Dictionary<string, object>
            {
                {"8/1", 17},
                {"8/2", 28},
                {"8/3", 15},
                {"8/4", 22},
                {"8/5", 32},
                {"8/6", 26},
                {"8/7", 25},
                {"8/8", 18},
                {"8/9", 18},
                {"8/10", 23},
                {"8/11", 21},
                {"8/12", 24},
                {"小计", 269},
            }
        });
        list.Add(new Report
        {
            WayName = "第三方（美团等）",
            Extend = new Dictionary<string, object>
            {
                {"8/1", 5},
                {"8/2", 0},
                {"8/3", 2},
                {"8/4", 6},
                {"8/5", 4},
                {"8/6", 1},
                {"8/7", 3},
                {"8/8", 6},
                {"8/9", 5},
                {"8/10", 2},
                {"8/11", 2},
                {"8/12", 1},
                {"小计", 37},
            }
        });
        list.Add(new Report
        {
            WayName = "日配送单量",
            Extend = new Dictionary<string, object>
            {
                {"8/1", 22},
                {"8/2", 28},
                {"8/3", 17},
                {"8/4", 28},
                {"8/5", 36},
                {"8/6", 27},
                {"8/7", 28},
                {"8/8", 24},
                {"8/9", 23},
                {"8/10", 25},
                {"8/11", 23},
                {"8/12", 25},
                {"小计", 306},
            }
        });
        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Report>()
        {
            Data = list,
            DynamicColumns = new List<string>() { "8/1", "8/2", "8/3", "8/4", "8/5", "8/6", "8/7", "8/8", "8/9", "8/10", "8/11", "8/12", "小计" }
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 额外表头 合并单元格多种情况
    /// </summary>
    [Fact]
    public async Task Test_Export_HeaderRow_OneColumnSpan_MultiRowSpan()
    {
        var data = new List<Bing.Offices.Tests.Models.ExportFormat>();
        for (int i = 0; i < 100; i++)
        {
            data.Add(new Bing.Offices.Tests.Models.ExportFormat()
            {
                Id = $"A{i}",
                Name = $"测试名称+++++{i}",
                Index = i + 1,
                IgnoreProperty = $"忽略属性+++++{i}",
                NotMappedProperty = $"忽略映射属性+++++{i}",
                Money = i * 1000,
                CreateTime = DateTime.Now.AddMinutes(i)
            });
        }

        var rows = new List<Bing.Offices.Metadata.Excels.Row>();
        var row = new Bing.Offices.Metadata.Excels.Row(2);
        row.Add(new Cell("单列多行", 2, 1, 2));
        row.Add(new Cell("单列单行", 4, 1, 1));
        rows.Add(row);

        var row2 = new Bing.Offices.Metadata.Excels.Row(5);
        row2.Add(new Cell("多列单行", 2, 2, 1));
        row2.Add(new Cell("多列多行", 4, 2, 2));
        rows.Add(row2);

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<Bing.Offices.Tests.Models.ExportFormat>()
        {
            HeaderRow = rows.ToList<IRow>(),
            HeaderRowIndex = 1,
            DataRowStartIndex = 2,
            Data = data,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 合并单元格
    /// </summary>
    [Fact]
    public async Task Test_Export_MergeRow()
    {
        var orders = new List<ExportOrderWithMerged>();
        for (var i = 0; i < 1000; i++)
        {
            var index = new Random(i).Next(i + 10, i + 13);
            var orderNumber = $"订单{index}";
            orders.Add(new ExportOrderWithMerged
            {
                Buyer = $"下单人{index}",
                Price = Math.Round(new Random(i).NextDouble(), 2),
                BuyQty = new Random(i).Next(1, 10),
                ProductName = $"商品{i}",
                OrderNumber = orderNumber,
                Index = orderNumber,
            });
        }

        var bytes = await _excelExportService.ExportAsync(new ExportOptions<ExportOrderWithMerged>
        {
            Data = orders,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }

    /// <summary>
    /// 测试 - 导出 值映射
    /// </summary>
    [Fact]
    public async Task Test_Export_ValueMapping()
    {
        var codes = new[] { "100", "200", "300", "400" };
        var entities = new List<ExportValueMapping>();
        for (int i = 0; i < 1000; i++)
        {
            var index = new Random().Next(0, 4);
            var code = codes[index];
            entities.Add(new ExportValueMapping
            {
                Index = i,
                Code = code
            });
        }
        var bytes = await _excelExportService.ExportAsync(new ExportOptions<ExportValueMapping>
        {
            Data = entities,
        });
        await File.WriteAllBytesAsync($"D:\\测试导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);
    }
}

/// <summary>
/// 条形码
/// </summary>
public class Barcode
{
    /// <summary>
    /// 系统
    /// </summary>
    [ColumnName("系统编号")]
    public string Id { get; set; }

    /// <summary>
    /// 国标码
    /// </summary>
    [ColumnName("国标码")]
    [Required]
    public string Code { get; set; }

    /// <summary>
    /// 创建时间
    /// </summary>
    [ColumnName("创建时间")]
    public string CreateDate { get; set; }

    /// <summary>
    /// 年龄
    /// </summary>
    [ColumnName("年龄")]
    public int? Age { get; set; }

    /// <summary>
    /// 扩展
    /// </summary>
    [DynamicColumn]
    public IDictionary<string, object> Extend { get; set; }
}

/// <summary>
/// 报表
/// </summary>
public class Report
{
    /// <summary>
    /// 配送方式
    /// </summary>
    [ColumnName("配送方式")]
    public string WayName { get; set; }

    /// <summary>
    /// 扩展
    /// </summary>
    [DynamicColumn]
    public IDictionary<string, object> Extend { get; set; }
}
