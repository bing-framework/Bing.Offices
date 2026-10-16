# 快速开始

## 功能与安装

Bing.Offices 通过请求模型和接口执行 Excel/CSV 导入导出。常规业务先选一个 Excel Provider；NPOI 支持 XLS/XLSX，ClosedXML、MiniExcel 支持各自能力范围内的 XLSX。只读与前向导出选型见 [Provider 矩阵](excel/09-providers.md)。

仓库 `version.props` 定义的版本为 `2.0.0`，不表示该版本已在公共 NuGet 源发布。应用应使用其包源中实际可获取且配套的版本；线上发布状态为**待实现/待确认（待确认）**。从源码构建参见[构建与验证](development/build-validation.md)。

NPOI 业务项目需要 `Bing.Offices.Npoi` 和依赖注入容器；Core、Abstractions 通过项目/包依赖传递。Provider 面向 `net6.0;net8.0`，Core/Abstractions 面向 `netstandard2.0`。这不表示所有运行时包都可在 netstandard2.0 应用中使用。

## Excel 往返示例

下例是可放入 .NET 控制台程序的完整示例。导入的集合属性需可写，行模型提供公共无参构造；默认表头来自属性名。

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

using var services = new ServiceCollection()
    .AddBingOfficesNpoi()
    .BuildServiceProvider();
var exporter = services.GetRequiredService<IExcelExporter>();
var importer = services.GetRequiredService<IExcelImporter>();
var rows = new[] { new OrderRow { Code = "PO-001", Amount = 12.5m } };
var exportRequest = ExcelExport.Workbook(book => book.AddSheet("订单", rows));
using var stream = new MemoryStream();
await exporter.ExportAsync(exportRequest, stream);
stream.Position = 0;
var importRequest = ExcelImport.Workbook<OrderWorkbook>(book =>
    book.Sheet("订单", model => model.Orders));
var result = await importer.ImportAsync(stream, importRequest);
if (!result.IsSuccess)
    throw new InvalidOperationException("导入失败，请检查 result.Errors 中的位置和错误码。");
Console.WriteLine(result.Workbook.Orders[0].Code);

public sealed class OrderRow
{
    public string Code { get; set; }
    public decimal Amount { get; set; }
}

public sealed class OrderWorkbook
{
    public List<OrderRow> Orders { get; set; } = new();
}
```

文件输出可调用 `exporter.ExportToFile(request, path)` 或异步版本，以复用同目录临时文件提交。调用方直接提供的流由调用方释放，读取刚生成的流前需恢复 `Position`；对不可定位输入不能假定可重复读取。

## 核心实现与配置

| 环节 | 关键文件或目录 |
| --- | --- |
| 请求构建、配置快照和格式定义 | `src/Bing.Offices.Abstractions/Bing/Offices/{Exports,Imports}` |
| 属性/Profile/文档/请求配置编译为映射计划 | `src/Bing.Offices.Core/Bing/Offices/Mappings` |
| 默认 CSV、配置加载器和文件提交器注册 | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactoryProvider.cs` |
| NPOI 注册 | `src/Bing.Offices.Npoi/Bing/Offices/Extensions/Extensions.Service.cs` |
| NPOI 实际读写 | `src/Bing.Offices.Npoi/Bing/Offices/{Imports,Exports}` |

通过 [Profile](excel/mapping-profile.md)、[JSON/XML v2](excel/mapping-json-xml.md) 配置列名、别名、校验和转换。优先级为约定 < Attribute < Profile < 文档 < 请求。注册 NPOI 等数据 Provider 时同时提供默认 CSV 服务；只有 CSV 需求时也可用 `ExcelMappingPlanFactoryProvider.RegisterDefault(services)` 注册 Core 服务，详细示例见 [CSV](csv/README.md)。

## 限制与注意事项

- 同一公共接口采用 `TryAdd` 注册；连续注册不同 Provider 不会按请求自动选择。需要按操作选型时显式构造相应实现或使用 [完整/前向导出策略](excel/09-providers.md)。
- 异步 API 的外围 IO 使用异步读写，DOM 构建、映射、引擎序列化仍有同步阶段；详见[异步 IO](excel/async-io.md)。
- `ExcelResourceLimits` 默认限制输入字节，并提供行数、错误、唯一值、关系键及 XLSX ZIP/XML 等上限。各入口的检查范围以 Provider 实现为准；上限不等于进程内存配额。
- 每个 Provider 只承诺其能力矩阵中的功能。不能因 Abstractions 存在某类型就假定所有引擎支持。

下一步可阅读[工作簿报表](excel/workbook-reports.md)、[实体布局](excel/entity-layout.md)、[批量导入](excel/batch-import.md)。
