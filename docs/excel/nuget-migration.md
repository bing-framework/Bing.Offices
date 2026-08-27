# NuGet Migration

当前 major 为 `2.x`。三个包身份和依赖方向保持不变：

`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`

迁移建议：

1. 新代码使用四种方向明确的 Profile 契约，按方向配置 Builder。
2. JSON/XML 使用 v2 `ExcelMappingDocument`；旧平铺 v1 必须通过 `MigrateV1Json`/`MigrateV1Xml` 显式选择方向。
3. 使用 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelMaxValue`、`ExcelRange`、`ExcelMaxLength`、`ExcelUnique`。
4. 继续通过 `AddNpoi(): void` 注册 NPOI；Profile Registry 使用独立的显式或程序集扫描扩展。
5. 只依赖 provider-neutral 请求、结果和转换器接口，不引用 NPOI 类型。
6. 调用方提供的输入、输出和配置流仍由调用方拥有，库不会关闭它们。

当前 2.x 保留既有公开入口以维持源码和二进制兼容，包括 `ExcelMapping.For<T>()`、旧的映射重载、布尔语义配置方法和 `ICellValueConverter` 兼容层。它们不应作为新代码的首选入口；迁移到方向明确的映射配置和 `IExcelValueConverter` 时，应在现有测试通过后逐项验证。

下一 major 的 breaking table 尚未在本任务中获得批准，因此本任务不删除或重命名上述 API，也不执行包发布。`dotnet pack` 和本地 package consumer 验证仅用于确认当前包的可消费性。

## 候选 API 迁移对照

下表是待批准的 next-major 候选项，不代表当前 2.x 已删除或已承诺删除：

| 当前 2.x 入口 | 新代码建议 | 当前兼容策略 | 批准状态 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>()` | 按导入/导出方向使用对应 Builder | 保留 | 待批准 |
| `Mapping(configuration)` / `Mapping(document)` | 使用具名的 MappingConfiguration/MappingDocument 入口 | 保留 | 待批准 |
| `HeaderMatch` | `RequireExpectedHeaders` | 保留 | 待批准 |
| `MaxColumnCount` | `MaxReadColumns` | 保留 | 待批准 |
| `EnabledEmptyLine` | `ReportEmptyRows` | 保留 | 待批准 |
| `IgnoreEmptyLineAfterData` | `StopAtFirstEmptyRow` | 保留 | 待批准 |
| `AddNavigationSheet` | `AddSheet(name, parents.SelectMany(...))` | 保留 | 待批准 |
| `ExcelSetting.Default` | request/DI 范围内的显式设置 | 保留 | 待批准 |
| `ICellValueConverter` | `IExcelValueConverter` | 保留兼容层 | 待批准 |

迁移示例：

```csharp
// 2.x 兼容入口
var mappingBuilder = ExcelMapping.For<OrderRow>();
mappingBuilder.Property(row => row.Code).HasTitle("订单号");
var mapping = mappingBuilder.Build();

// 新代码优先使用方向明确的 Profile/Mapping 配置
services.AddMappingProfile<OrderProfile>();
var request = ExcelImport.Workbook<OrderWorkbook>(builder =>
	builder.Sheet("订单", workbook => workbook.Rows));
```

在 breaking table 获得批准前，旧入口继续纳入包消费者兼容测试；批准后必须同步 obsolete shim、删除项 negative baseline、迁移示例和版本号。
