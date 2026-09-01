# NuGet Migration

当前工作树包版本为 `2.0.0`，三个包身份和依赖方向保持不变：

`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`

迁移建议：

1. 新代码使用四种方向明确的 Profile 契约，按方向配置 Builder。
2. JSON/XML 使用 v2 `ExcelMappingDocument`；旧平铺 v1 必须通过 `MigrateV1Json`/`MigrateV1Xml` 显式选择方向。
3. 使用 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelMaxValue`、`ExcelRange`、`ExcelMaxLength`、`ExcelUnique`。
4. 继续通过 `AddNpoi(IServiceCollection): void` 注册 NPOI；Profile Registry 使用独立的显式或程序集扫描扩展。
5. 只依赖 provider-neutral 请求、结果和转换器接口，不引用 NPOI 类型。
6. 调用方提供的输入、输出和配置流仍由调用方拥有，库不会关闭它们。

当前 `2.0.0` 工作树移除了 `ExcelSetting.Default` 静态可变入口；Workbook metadata 改为请求级快照。CSV 新增输入字节、行、错误、字段和列限制，超限结果使用 `CsvImportErrorCode.ResourceLimit` 并设置 `CsvImportResult.IsTruncated`。

除 `ExcelSetting.Default` 外，其它 breaking table 尚未在本任务中获得批准，因此本任务不删除或重命名其它 API，也不执行包发布。`dotnet pack` 和本地 package consumer 验证仅用于确认当前 `2.0.0` 包的可消费性。

## 候选 API 迁移对照

下表列出仍保持兼容的旧版入口和候选替代项；除 `ExcelSetting.Default` 外，不代表本次 3.0.0 已删除或承诺删除：

| 当前 2.x 入口 | 新代码建议 | 当前兼容策略 | 批准状态 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>()` | 按导入/导出方向使用对应 Builder | 保留 | 待批准 |
| `Mapping(configuration)` / `Mapping(document)` | 使用具名的 MappingConfiguration/MappingDocument 入口 | 保留 | 待批准 |
| `HeaderMatch` | `RequireExpectedHeaders` | 保留 | 待批准 |
| `MaxColumnCount` | `MaxReadColumns` | 保留 | 待批准 |
| `EnabledEmptyLine` | `ReportEmptyRows` | 保留 | 待批准 |
| `IgnoreEmptyLineAfterData` | `StopAtFirstEmptyRow` | 保留 | 待批准 |
| `AddNavigationSheet` | `AddSheet(name, parents.SelectMany(...))` | 保留 | 待批准 |
| `ExcelSetting.Default` | `ExcelWorkbookMetadataOptions` request 范围内的显式设置 | 已移除 | 本任务批准 |
| `ICellValueConverter` | `IExcelValueConverter` | 保留兼容层 | 待批准 |

迁移示例：

```csharp
// 仍兼容的旧版入口
var mappingBuilder = ExcelMapping.For<OrderRow>();
mappingBuilder.Property(row => row.Code).HasTitle("订单号");
var mapping = mappingBuilder.Build();

// 新代码优先使用方向明确的 Profile/Mapping 配置
services.AddMappingProfile<OrderProfile>();
var request = ExcelImport.Workbook<OrderWorkbook>(builder =>
	builder.Sheet("订单", workbook => workbook.Rows));
```

`ExcelSetting.Default` 已不再提供静态可变实例。请在每个 Workbook 请求上调用 `Metadata(...)`；不配置时使用请求级默认值。其余候选入口仍保持兼容，只有在单独批准对应 breaking table 后才可删除或重命名。
