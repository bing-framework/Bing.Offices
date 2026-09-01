# API Breaking Changes

状态：`BLOCKED`（不是代码失败，而是需用户/版本负责人批准后才能进行收敛）。

当前暂无本轮新增 breaking change。

已有工作树决策：`ExcelSetting.Default` 已移除，迁移至请求级 `ExcelWorkbookMetadataOptions` 和 `ExcelWorkbookExportBuilder.Metadata(...)`；版本保持 `2.0.0`。其余 API 收敛项在 API 分类账和用户版本决策前不得删除或重命名。

后续每项 API 变化必须记录 old/new、source/binary 影响、迁移示例、版本策略及 PackageConsumer 证据。

## 当前分类盘点

- 用户主路径：`ExcelImport.Workbook`、`ExcelExport.Workbook`、`IExcelImporter`、`IExcelExporter`、CSV 接口、请求/结果/metadata/资源限制类型。
- Provider SPI：`IExcelMappingPlanFactory`、`IExcelMappingPlan`、列/Sheet/Workbook plan 和 provider-neutral 样式/布局接口；需保持 provider-neutral。
- 合法扩展入口：映射 Profile、值转换器和校验规则接口在当前清单中属于 `User API`，因为 Docs/测试消费者直接注册或实现它们。
- 兼容层：`ExcelMapping.For<T>()`、旧属性/旧 `ICellValueConverter`、旧异常/辅助入口；当前仍被测试或文档消费者引用。
- 执行细节候选：`MappingConfigurationMerger`、具体 plan/type map、具体 validation、`UniqueTracker`、`SheetSetting`、`RegexConst` 等；当前公开且存在仓库引用，未获删除批准。

当前 `PublicApiContractTest` 仍将上述类型作为 baseline，证明本轮选择的是治理和可追溯，而不是未经批准的 API 收敛。Round 4 已将清单中的用户主路径、Provider SPI、兼容层和执行细节分类与本报告对齐；导出请求中的样式、布局和图表类型属于 `User API`，映射内部样式配置、`MappingProfileRegistry`、`UniqueTracker` 和具体校验实现属于 `Execution detail`。类型清单对实际导出类型执行 exact lookup，缺失或多余符号直接失败，并校验 Provider SPI 保持 `EditorBrowsable.Never`。

Round 4 的 `PublicApi_ExportedTypes_ShouldHaveGovernedClassification` 为每个导出类型的 public/protected/protected-internal 构造函数、属性、字段和方法生成稳定 member key，并将其绑定到类型分类、source/binary 影响和迁移策略；重复或缺少治理字段直接失败。这是自动生成的成员治理账，不把手工摘要误称为逐成员批准。四个独立静态快照见：[netcoreapp3.1](api-diff-netcoreapp3.1.md)、[net6.0](api-diff-net6.0.md)、[net7.0](api-diff-net7.0.md) 和 [net8.0](api-diff-net8.0.md)。

## 符号级分类账

分类值固定为：`User API`（业务调用和合法扩展入口）、`Provider SPI`（实现提供程序使用的隐藏契约）、`Compatibility`（历史兼容入口）、`Execution detail`（当前仍公开但不应新增依赖的实现细节）。本轮不改变现有 public/protected 可见性。

| 符号范围 | 分类 | 当前证据 | 当前动作 | Source/Binary 影响 | 版本/迁移策略 | 批准状态 |
| --- | --- | --- | --- | --- | --- | --- |
| `ExcelImport.Workbook`、`ExcelExport.Workbook`、`IExcelImporter`、`IExcelExporter`、`ICsvImporter`、`ICsvExporter` | `User API` | Stream-first 单元测试、Docs consumer、NPOI DI 注册和公开成员基线 | 保持公开；作为用户主路径 | 无 | 继续以请求对象和调用方流为主；不引入旧服务入口 | 已批准保留 |
| `ExcelWorkbookImportRequest*`、`ExcelSheetImportRequest`、`ExcelWorkbookImportResult*`、`ExcelSheetImportResult`、`ExcelImportError*`、`ExcelResourceLimits`、`ExcelImportFailure*` | `User API` | 资源、错误、Failure Workbook、取消和格式测试；公共导出基线 | 保持公开 | 无 | 新增字段采用默认兼容语义；限制和诊断不泄漏内部类型 | 已批准保留 |
| `ExcelWorkbookExportRequest`、`ExcelSheetExportRequest`、`ExcelWorkbookMetadataOptions`、`ExcelDynamicColumnDefinition`、样式/布局/图表请求类型 | `User API` | XLS/XLSX 导出、模板、metadata、样式、图表和 consumer 测试 | 保持公开 | 无 | metadata 通过请求级配置；模板流遵守 leave-open | 已批准保留 |
| `IExcelValueConverter`、`INamedExcelValueConverter`、`IExcelValidationRule`、`INamedExcelValidationRule`、`IImportMappingProfile*`、`IExportMappingProfile*`、`IMappingProfile*`、`IMappingProfileRegistry`、`IMappingProfileResolver` | `User API` | 转换器/校验器/Profile 注册测试、JSON/XML/Fluent 主链和 consumer 使用 | 保持公开；作为合法扩展点 | 无 | 新扩展优先 provider-neutral 契约；旧 `ICellValueConverter` 保留兼容 | 已批准保留 |
| `IExcelMappingPlan`、`IExcelMappingColumn`、`IExcelDynamicMappingColumn`、`IExcelMappingStyle`、`IExcelMappingLayout`、`IExcelMappingWorkbookPlan`、`IExcelMappingSheetPlan`、`IExcelMappingPlanFactory` | `Provider SPI` | 三个接口程序集边界、NPOI 无泄漏检查、`EditorBrowsable.Never` 检查和 mapping plan 测试 | 保持 public SPI 但隐藏 IntelliSense | 删除/改签名会影响 provider binary；本轮无变化 | Provider 实现通过这些 provider-neutral 类型协作；未来 breaking 需版本升级和迁移 | 已批准保留 |
| `ExcelMapping.For<T>`、`ExcelMappingBuilder*`、`ExcelColumnMappingBuilder*`、旧属性 `RequiredAttribute`/`RegexAttribute`/`DuplicationAttribute` 等、`ICellValueConverter` | `Compatibility` | 旧测试、迁移文档、legacy converter 和历史模型仍引用；旧属性已有 `Obsolete` | 保持公开和现有行为 | 删除或改名会产生 source/binary breaking | 继续引导到 `ExcelRequiredAttribute`、`ExcelRegexAttribute`、`ExcelUniqueAttribute` 和 `IExcelValueConverter` | 未批准删除 |
| `OfficeException`、旧数据/表头/空行异常、legacy stream helpers | `Compatibility` | 旧异常测试、公共成员快照和历史扩展路径 | 保持现状；不新增依赖 | 删除会影响 catch、编译和二进制绑定 | 新错误优先结构化结果；兼容异常只在既有入口保留 | 未批准删除 |
| `MappingConfigurationMerger`、`ExcelMappingDocumentFactory`、具体 loader、`ExcelTypeMap*`、`ExcelPropertyMap`、`ExcelValueMap*`、`ExcelValidationBindingFactory`、具体 validation rule、`UniqueTracker`、`SheetSetting`、`RegexConst` | `Execution detail` | 现有导出类型基线、仓库测试引用和内部计划编译路径；部分类型标记 `EditorBrowsable.Never` | 本轮不删除、不重命名、不改为 internal；禁止新增文档主路径依赖 | 任一可见性或签名收敛都可能是 source/binary breaking | 后续先补 consumer/多 TFM API diff，再由版本负责人批准 breaking；新代码只依赖上层契约 | 未批准 breaking |
| `ExcelMappingPlanFactory`、`ExcelMappingPlanFactoryProvider`、`MappingProfileRegistry` | `Execution detail`（注册/计划编译实现） | cache 命中、淘汰、快照并发和 Profile 注册测试；当前仍由测试/部分仓库代码直接使用 | 保持现有公开符号；不承诺为用户主 API | 改签名会影响直接调用方和 provider binary | 引导调用方使用 `IExcelMappingPlanFactory`/Profile 契约；正式收敛需迁移示例和版本策略 | 未批准 breaking |
| `Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddNpoi(IServiceCollection)` | `User API` | NPOI 程序集精确成员基线、DI 生命周期测试、Docs migration 文档 | 保持唯一公开注册入口 | 无 | 继续作为 package consumer 的 NPOI 注册入口 | 已批准保留 |

## 多 TFM 门禁

发布程序集目标为 Abstractions/Core `netstandard2.0`，NPOI 由工程配置构建 `netcoreapp3.1;net6.0;net7.0;net8.0`。四个 shipped NPOI TFM 均有独立的静态 API 输出，使用相同的 canonical member line 和 SHA-256 口径；当前 Npoi 输出 hash 一致，但 Abstractions/Core 当前 hash 与批准基线不匹配。net6/net8 API contract 当前均为 `6/7`，唯一失败为批准 hash 断言；net7/netcoreapp3.1 仅完成静态生成，且本机缺少对应 runtime，不能作为运行时 PASS。测试使用 `NETCOREAPP3_1`、`NET6_0`、`NET7_0`、`NET8_0` 的显式分支，不使用通用 `#else` 代表 shipped TFM。未批准任何跨 TFM 差异，因此出现新增导出类型或成员签名变化时测试必须失败并进入 API 审批。

## 迁移决策

本轮没有批准删除、重命名、降可见性或版本升级；因此不存在可执行的 old/new breaking migration。已完成的非 breaking 文档迁移是 `AddNpoi(IServiceCollection): void` 和 metadata 请求级配置。任何未来 `Execution detail` 收敛必须同时提交 symbol ledger、source/binary 影响、示例迁移、版本策略、各 shipped TFM API diff 和 packed-package consumer 证据后再审批。当前 API 自动生成/比较入口已补齐，但当前 Abstractions/Core hash mismatch、runtime gate 失败、版本负责人批准和 Git 交付面仍是发布前置条件。
