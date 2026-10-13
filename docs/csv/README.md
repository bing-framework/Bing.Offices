# CSV 导入导出

CSV 能力由 `Bing.Offices.Core` 实现，复用与 Excel 相同的 Mapping Plan、值转换、校验、异常和文件提交基础设施。它提供同步与真实异步流 IO，不依赖具体 Excel 引擎。

## 注册与入口

任一官方数据 Provider（NPOI、ClosedXML、MiniExcel、ExcelDataReader）或 SpreadCheetah 流式 Provider 的注册扩展都会注册默认 Core 服务，因此可同时解析 CSV 接口。只提供商业扩展接口的 Aspose.Cells 不注册 CSV 服务：

```csharp
using Bing.Offices.Csv;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

var services = new ServiceCollection();
services.AddBingOfficesNpoi();
using var provider = services.BuildServiceProvider();

var exporter = provider.GetRequiredService<ICsvExporter>();
var importer = provider.GetRequiredService<ICsvImporter>();
```

CSV 引擎本身位于 Core；选择 NPOI、ClosedXML、MiniExcel、ExcelDataReader 或 SpreadCheetah 作为注册入口，不会改变 CSV 文件格式。多次注册时默认服务采用 `TryAdd`，已注册的自定义实现保持优先。仅使用 Aspose.Cells 时，宿主需要另行注册 Core 默认服务或自己的 `ICsvImporter`/`ICsvExporter` 实现。

## 基本用法

```csharp
using System.Text;
using Bing.Offices.Csv;

var rows = new[]
{
    new OrderRow { Number = "A-001", Amount = 12.5m }
};

var exportOptions = new CsvExportOptions<OrderRow>
{
    Encoding = new UTF8Encoding(false),
    IncludeHeader = true,
    FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
};

await exporter.ExportToFileAsync(rows, "orders.csv", exportOptions);

await using var source = File.OpenRead("orders.csv");
var result = await importer.ImportAsync<OrderRow>(source,
    new CsvImportOptions<OrderRow>
    {
        Encoding = new UTF8Encoding(false, true),
        MaxInputBytes = 10 * 1024 * 1024,
        MaxRows = 100_000,
        MaxErrors = 100
    });

public sealed class OrderRow
{
    public string Number { get; set; }
    public decimal Amount { get; set; }
}
```

`CsvImportResult<T>.Items` 包含成功物化的记录，`Errors` 包含从一开始计数的行列位置、属性名和错误代码。达到 `MaxErrors` 等限制时，`IsTruncated` 标识结果已提前停止。

## 核心实现

| 职责 | 关键位置 |
| --- | --- |
| 公共接口与选项 | `src/Bing.Offices.Abstractions/Bing/Offices/Csv` |
| 同步/异步导入导出 | `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityImporter.cs`、`CsvEntityExporter.cs` |
| 字段解析与写入 | `CsvRecordReader.cs`、`CsvRecordWriter.cs` 及异步对应实现 |
| 表头、属性与动态列绑定 | `CsvHeaderBinder.cs`、`CsvPropertyBinding.cs`、`CsvDynamicTypeResolver.cs` |
| 字节和文件便利入口 | `src/Bing.Offices.Core/Bing/Offices/Extensions/CsvStreamExtensions.cs` |
| 默认依赖注入 | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactoryProvider.cs` |

导出器按 Mapping Plan 决定列顺序、标题、转换和动态列；导入器先绑定表头，再执行转换、值映射、内置/命名校验与唯一值跟踪。请求级 `MappingDocument` 优先作为规范化文档，`MappingConfiguration` 用于本次请求的方向化覆盖。

## 选项与配置

### 导出

- `IncludeHeader` 控制表头，默认 `true`。
- `Delimiter`、`Quote`、`NewLine`、`Culture` 和 `Encoding` 控制 CSV 方言；默认 UTF-8 无 BOM、逗号、双引号和 CRLF。
- `FormulaInjectionPolicy` 默认 `Escape`，对可能被表格软件解释为公式的字段添加单引号；只有明确接受风险时才使用 `None`。
- `DynamicColumns` 声明需要导出的动态列名。

### 导入

- `HasHeader` 与 `RequireExpectedHeaders` 控制表头绑定；无表头时按列计划顺序绑定。
- `MaxInputBytes`、`MaxRows`、`MaxErrors`、`MaxFieldLength`、`MaxColumns` 和 `MaxTrackedUniqueValues` 提供请求级资源限制。
- `UniqueComparison` 控制唯一值比较，默认忽略大小写的 Ordinal 比较。
- `Encoding` 默认 UTF-8 无 BOM；需要拒绝无效字节时应传入启用严格解码的 `UTF8Encoding`。

## IO、错误与生命周期

- `Import`/`ImportAsync` 和 `Export`/`ExportAsync` 均保持调用方 Stream 打开。
- 异步入口使用 Reader/Writer 和 Stream 的异步 IO，并传递 `CancellationToken`；不通过 `Task.Run` 伪装异步。
- 直接写入调用方 Stream 不提供回滚保证，失败时可能已有部分内容。
- `ExportToFile*` 通过 `IFileExportCommitter` 在同目录暂存并提交；失败或取消时保护已有目标文件。
- 记录级转换和校验错误进入 `CsvImportResult<T>.Errors`；配置、IO 和公共运行边界错误使用 Bing.Offices 结构化异常，并通知已注册的 `IBingOfficesExceptionObserver`。

## 已知限制与注意事项

- CSV 是单表、顺序记录格式，不支持 Excel 的 Sheet、模板、样式、合并、图片、公式计算、关系和 Entity Layout。
- `CsvFormulaInjectionPolicy.Escape` 只降低常见电子表格公式注入风险，不是通用内容清洗或授权机制。
- 唯一值校验会占用内存；不可信或大规模输入应设置 `MaxTrackedUniqueValues` 及其他资源限制。
- 自定义分隔符、引用符、换行、区域性和编码必须在导入导出两端保持一致。
- 当前没有独立的 CSV Provider 包；应用通过官方 Provider 注册 Core 默认服务，或自行注册 `ICsvImporter`/`ICsvExporter` 实现。

更多异步与文件提交语义见 [Async IO 合同](../excel/async-io.md) 和 [自定义文件提交服务](../excel/file-committers.md)。
