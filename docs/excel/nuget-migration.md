# NuGet Migration

当前工作树包版本为 `2.0.0`，三个包身份和依赖方向保持不变：

`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`

迁移建议：

1. 新代码使用四种方向明确的 Profile 契约，按方向配置 Builder。
2. JSON/XML 只使用 v2 `ExcelMappingDocument`；v1 平铺配置应在升级前离线转换。
3. 使用 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelMaxValue`、`ExcelRange`、`ExcelMaxLength`、`ExcelUnique`。
4. 通过 `AddBingOfficesNpoi(IServiceCollection): IServiceCollection` 注册 NPOI，并支持链式注册；Profile Registry 使用独立的显式或程序集扫描扩展。
5. 只依赖 provider-neutral 请求、结果和转换器接口，不引用 NPOI 类型。
6. 导入源流、导出目标流和配置流仍由调用方拥有。模板流由 `UseTemplate` 的 `leaveOpen` 参数控制，默认 false；复用模板时显式设置 true 并管理流位置和释放。

当前 `2.0.0` 工作树已移除旧的 `ExcelSetting`/`SheetSetting` 配置类型；Workbook metadata 改为请求级快照。CSV 新增输入字节、行、错误、字段和列限制，超限结果使用 `CsvImportErrorCode.ResourceLimit` 并设置 `CsvImportResult.IsTruncated`。

本次导入绑定校验没有公共签名变化：Excel/CSV 中不同表头重复绑定同一固定属性或已声明动态 Key 时，现在会返回 `InvalidHeader`。CSV 无表头导入动态字典属性时，显式定义的动态列会按与导出一致的映射及动态列顺序展开；缺少定义时明确报错，避免静默丢弃动态值。

失败工作簿输出限制已从 `MaxBytes` 更名为 `MaxSerializedBytes`。该限制只约束失败工作簿序列化输出，不代表原始 Workbook、解压内容、实体或失败输出 DOM 的内存上限。

旧的 `OfficeException` 异常层级、`ExcelSetting`/`SheetSetting` 已从发布 API 移除；CSV 实体实现类不再作为发布 API，调用方应通过 `ICsvImporter`/`ICsvExporter` 或 DI 使用。本轮另经成员级批准移除四个重复的 `ExportToFile`/`ExportToFileAsync` 扩展声明，接口实例成员保持不变。`dotnet pack` 和本地 package consumer 验证确认当前 `2.0.0` 包的可消费性。

## API 迁移对照

### 本次易用性增量（兼容新增）

以下 API 已在当前源码实现，使用 NuGet 时需选择包含本次变更的包；不代表已向公共源发布。现有 Workbook 请求、接口、构造器和成员均保留，原调用无需迁移。

| 需求 | 新入口 | 保留的路径 |
| --- | --- | --- |
| 仅使用 CSV | `services.AddBingOfficesCore()`，位于 `Bing.Offices.Extensions` | 原 Provider SPI 仍可用；NPOI 注册也包含 Core 服务 |
| 单表 Excel 导入 | `importer.ImportSheet<T>(stream)` / `ImportSheetAsync<T>` | 原 `ExcelImport.Workbook<TWorkbook>` 和 `IExcelImporter.Import*` |
| 从路径/字节数组导入单表 | `ImportSheetFromFile*` / `ImportSheetFromBytes*` | 原 Workbook 文件/字节扩展 |
| 判断部分导入 | CSV、Workbook、单表结果的 `IsComplete` / `StopReason`；CSV 新增 `IsSuccess` | 旧 `ErrorsTruncated`、`IsTruncated` 不改义 |

单表入口默认选择物理索引 0；通过 `selector: ExcelSheetSelector.ByName("订单")` 选择表，通过 `configure` 设置映射和 `MaxReadColumns`，通过 `resourceLimits` 设置行数和错误上限。返回结果直接提供 `Items`、`Errors` 和零基 `SourceRows`。复杂多 Sheet、关系和失败工作簿配置继续使用 Workbook Request。

行为补正：Excel 的 MaxRows 恰好在 Sheet 边界耗尽而仍有请求 Sheet 时，现在返回 ResourceLimit，避免把未执行的后续 Sheet 静默当作成功；已导入数据不丢弃。错误与失败工作簿摘要使用已解析的实际表名，只有缺失 Sheet 才回退到选择器描述。完整示例见 [快速开始](quick-start.md)，结果字段语义见 [导入校验](import-validation.md)。

默认输入保护补正：此前省略 ResourceLimits 会漏掉输入复制阶段的 128 MiB 限制，现在 Workbook/单表、Sync/Async 均默认执行。依赖旧行为读取更大输入的调用方，应显式配置合适的 MaxInputBytes；确需关闭时设置 `new ExcelResourceLimits { MaxInputBytes = null }`，并自行管理资源预算。此修复不改变公共签名，也不承诺限制 DOM 峰值内存。

下表列出仍保持兼容的旧版入口和候选替代项；已标记移除的类型不应继续在新代码中引用：

| 当前 2.x 入口 | 新代码建议 | 当前兼容策略 | 批准状态 |
| --- | --- | --- | --- |
| `ExcelMapping.For<T>()` | `ImportMappingBuilder<T>` / `ExportMappingBuilder<T>` | 已移除 | 本次 RC Breaking Change |
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
| `CsvStreamExtensions.ExportToFile*` / `ExcelStreamExtensions.ExportToFile*` | `ICsvExporter.ExportToFile*` / `IExcelExporter.ExportToFile*` | 扩展声明已移除；接口实例成员保留 | FIX-003，`approvedBy=jian玄冰` |

迁移示例：

```csharp
// 使用方向明确的 Builder/Profile 配置
var importMapping = new ImportMappingBuilder<OrderRow>()
    .Property(row => row.Code).HasHeader("订单号").And()
    .Build();
services.AddMappingProfile<OrderProfile>();
var request = ExcelImport.Workbook<OrderWorkbook>(builder =>
	builder.Sheet("订单", workbook => workbook.Rows));
```

请在每个 Workbook 请求上调用 `Metadata(...)`；不配置时使用请求级默认值。CSV 调用方应依赖 `ICsvImporter`/`ICsvExporter`，不要直接构造 Core 内部实现。对于已移除的四个文件导出扩展，直接改为接口实例调用：`exporter.ExportToFile(...)` 或 `await exporter.ExportToFileAsync(...)`。其余未列入本次批准范围的入口保持兼容。
