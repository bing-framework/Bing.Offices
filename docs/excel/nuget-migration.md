# NuGet Migration

当前工作树包版本为 `2.0.0`，三个包身份和依赖方向保持不变：

`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`

迁移建议：

1. 新代码使用四种方向明确的 Profile 契约，按方向配置 Builder。
2. JSON/XML 使用 v2 `ExcelMappingDocument`；旧平铺 v1 必须通过 `MigrateV1Json`/`MigrateV1Xml` 显式选择方向。
3. 使用 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelMaxValue`、`ExcelRange`、`ExcelMaxLength`、`ExcelUnique`。
4. 通过 `AddBingOfficesNpoi(IServiceCollection): IServiceCollection` 注册 NPOI，并支持链式注册；Profile Registry 使用独立的显式或程序集扫描扩展。
5. 只依赖 provider-neutral 请求、结果和转换器接口，不引用 NPOI 类型。
6. 调用方提供的输入、输出和配置流仍由调用方拥有，库不会关闭它们。

当前 `2.0.0` 工作树已移除旧的 `ExcelSetting`/`SheetSetting` 配置类型；Workbook metadata 改为请求级快照。CSV 新增输入字节、行、错误、字段和列限制，超限结果使用 `CsvImportErrorCode.ResourceLimit` 并设置 `CsvImportResult.IsTruncated`。

失败工作簿输出限制已从 `MaxBytes` 更名为 `MaxSerializedBytes`。该限制只约束失败工作簿序列化输出，不代表原始 Workbook、解压内容、实体或失败输出 DOM 的内存上限。

旧的 `OfficeException` 异常层级、`ExcelSetting`/`SheetSetting` 已从发布 API 移除；CSV 实体实现类不再作为发布 API，调用方应通过 `ICsvImporter`/`ICsvExporter` 或 DI 使用。其余 breaking table 尚未在本任务中获得批准，因此本任务不再扩展删除范围。`dotnet pack` 和本地 package consumer 验证仅用于确认当前 `2.0.0` 包的可消费性。

## 候选 API 迁移对照

下表列出仍保持兼容的旧版入口和候选替代项；已标记移除的类型不应继续在新代码中引用：

| 当前 2.x 入口 | 新代码建议 | 当前兼容策略 | 批准状态 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>()` | 按导入/导出方向使用对应 Builder | 保留 | 已确认方向中立映射入口 |
| `Mapping(configuration)` / `Mapping(document)` | 使用具名的 `MappingConfiguration(...)` / `MappingDocument(...)` 入口 | 迁移中 | 当前 Major 仍使用 `Mapping(...)` |
| `HeaderMatch` | `RequireExpectedHeaders` | 删除 | Major 已迁移 |
| `MaxColumnCount` | `MaxReadColumns` | 删除 | Major 已迁移 |
| `EnabledEmptyLine` | `ReportEmptyRows` | 删除 | Major 已迁移 |
| `IgnoreEmptyLineAfterData` | `StopAtFirstEmptyRow` | 删除 | Major 已迁移 |
| `AddNavigationSheet` | `AddSheet(name, parents.SelectMany(...))` | 保留 | 待批准 |
| `ExcelSetting` / `SheetSetting` | `ExcelWorkbookMetadataOptions` 与 Sheet request | 已移除 | Round 6 已批准 |
| `OfficeException` 异常层级 | 标准参数/状态异常与结构化导入错误 | 已移除 | Round 6 已批准 |
| `CsvEntityImporter` / `CsvEntityExporter` | `ICsvImporter` / `ICsvExporter`，优先通过 DI 获取 | 已 internal 化 | Round 6 已批准 |
| `ICellValueConverter` | `IExcelValueConverter` | 已移除；请迁移到提供程序无关的双向转换器 | 本任务已批准 |

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

请在每个 Workbook 请求上调用 `Metadata(...)`；不配置时使用请求级默认值。CSV 调用方应依赖 `ICsvImporter`/`ICsvExporter`，不要直接构造 Core 内部实现。其余候选入口仍保持兼容，只有在单独批准对应 breaking table 后才可删除或重命名。
