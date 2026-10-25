# CSV 与 Excel 快速开始

| 使用场景 | 添加包 | 服务注册 |
| --- | --- | --- |
| 只使用 CSV | `Bing.Offices.Core` | `AddBingOfficesCore()`，不依赖 NPOI |
| 需要 Excel | `Bing.Offices.Npoi` | `AddBingOfficesNpoi()`，依赖会带入 Core 和 Abstractions |

下列程序针对本仓库源码，或包含本次变更的本地构建包；这不表示新增 API 已进入当前 NuGet 已发布版本。

示例面向 .NET 8 Console 项目。每段程序都是独立入口，请分别创建项目，并用对应代码块替换该项目的默认 Program.cs。根据上表从包含这些 API 的源码或包源添加 Office 包；Console 项目还需显式添加依赖注入实现：

```shell
dotnet new console --framework net8.0
dotnet add package Microsoft.Extensions.DependencyInjection --version 8.0.0
```

两个示例都使用内存流创建自己的数据并完成导出、导入和结果检查，不依赖本机文件路径。

## CSV

```csharp
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Bing.Offices.Extensions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Microsoft.Extensions.DependencyInjection;

namespace BingOfficesQuickStart.Csv;

public sealed class OrderRow
{
    public string OrderNo { get; set; } = string.Empty;
    public string Customer { get; set; } = string.Empty;
}

public static class CsvQuickStartProgram
{
    public static async Task Main()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesCore();
        using var provider = services.BuildServiceProvider();
        if (provider.GetService<IExcelImporter>() is not null
            || provider.GetService<IExcelExporter>() is not null)
            throw new InvalidOperationException("Core 注册不应接入 Excel Provider。");
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var orders = new[]
        {
            new OrderRow { OrderNo = "ORD-123", Customer = "Alice" },
            new OrderRow { OrderNo = "ORD-124", Customer = "Bob" }
        };

        using var csv = new MemoryStream();
        await exporter.ExportAsync(orders, csv);
        csv.Position = 0;
        var result = await importer.ImportAsync<OrderRow>(csv);

        if (!result.IsComplete)
            Console.Error.WriteLine($"CSV 导入未完整完成：{result.StopReason}");
        if (!result.IsSuccess)
        {
            foreach (var error in result.Errors)
                Console.Error.WriteLine($"CSV 第 {error.RowIndex} 行，第 {error.ColumnIndex} 列：{error.Code} {error.Message}");
            throw new InvalidOperationException("CSV 导入未成功。");
        }
        if (result.StopReason != ImportStopReason.None || result.Items.Count != 2)
            throw new InvalidOperationException("CSV 导入结果状态或行数不符合预期。");

        var actual = result.Items.Select(order => $"{order.OrderNo}|{order.Customer}").ToArray();
        var expected = new[] { "ORD-123|Alice", "ORD-124|Bob" };
        if (!actual.SequenceEqual(expected))
            throw new InvalidOperationException("CSV 导入内容与导出数据不一致。");

        Console.WriteLine($"CSV 导入完成：{result.Items.Count} 行，{result.StopReason}");
    }
}
```

## Excel

```csharp
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace BingOfficesQuickStart.Excel;

public sealed class OrderRow
{
    public string OrderNo { get; set; } = string.Empty;
    public string Customer { get; set; } = string.Empty;
}

public static class ExcelQuickStartProgram
{
    public static async Task Main()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        using var provider = services.BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        var orders = new[]
        {
            new OrderRow { OrderNo = "ORD-123", Customer = "Alice" },
            new OrderRow { OrderNo = "ORD-124", Customer = "Bob" }
        };

        using var workbook = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Orders", orders));
        await exporter.ExportAsync(request, workbook);
        workbook.Position = 0;
        var result = await importer.ImportSheetAsync<OrderRow>(workbook);

        if (!result.IsComplete)
            Console.Error.WriteLine($"Excel 导入未完整完成：{result.StopReason}");
        if (!result.IsSuccess)
        {
            foreach (var error in result.Errors)
                Console.Error.WriteLine($"{error.SheetName}!R{error.RowIndex}C{error.ColumnIndex}：{error.Code} {error.Message}");
            throw new InvalidOperationException("Excel 导入未成功。");
        }
        if (result.StopReason != ImportStopReason.None || result.Items.Count != 2)
            throw new InvalidOperationException("Excel 导入结果状态或行数不符合预期。");
        if (!result.SourceRows.SequenceEqual(new[] { 1, 2 }))
            throw new InvalidOperationException("Excel 源行追溯信息不符合预期。");

        var actual = result.Items.Select(order => $"{order.OrderNo}|{order.Customer}").ToArray();
        var expected = new[] { "ORD-123|Alice", "ORD-124|Bob" };
        if (!actual.SequenceEqual(expected))
            throw new InvalidOperationException("Excel 导入内容与导出数据不一致。");

        Console.WriteLine($"Excel 导入完成：{result.Items.Count} 行，{result.StopReason}");
    }
}
```
