# 弃用与删除候选清单

## P0-03 完整扫描口径

本节是本轮 P0-03 的可复核扫描基线。扫描使用 Python `pathlib.Path.read_text(encoding="utf-8")` 按 UTF-8 读取文本；范围为 `src/`、`tests/`、`benchmarks/`、`docs/` 和根 `README.md`。为避免把编译输出、历史产物或旧二进制误计入，排除所有 `bin/`、`obj/`、`artifacts/`、`output/` 目录；候选定义和引用以源码/测试/文档/基准文本为准，公开 API 完整类型集合以 `tests/Bing.Offices.Tests/PublicApiContractTest.cs` 的 `ApiTypeCategories` 为准。

| 范围 | 文件数 | 说明 |
| --- | ---: | --- |
| `src/` | 222 | 生产 C# 文件 |
| `tests/` | 81 | Unit、Integration、Docs、ResourceProbe C# 文件及其测试配置 |
| `benchmarks/` | 62 | Benchmark C# 文件及其配置 |
| `docs/` + `README.md` | 14 | Markdown 文档和根 README |

### 特殊构造扫描结果

| 扫描项 | 命中 | 范围/结论 |
| --- | ---: | --- |
| `[Obsolete]` / `ObsoleteAttribute` | 初始 14；Round 5 最终 0 | 初始仅 `src/` 命中；Round 5 已删除 `ICellValueConverter`、DataTable/global CSV 旧状态和 6 个旧 validation attribute |
| `EditorBrowsableState.Never` | 71 | 仅生产源码；用于部分 request/plan/SPI 低层成员，不作为删除证明 |
| `EditorBrowsableState.Advanced` | 2 | `Resources.Designer.cs` 生成资源成员，不是 API 删除候选 |
| 生产 `NotImplementedException` | 3 | 均位于 `NpoiFailureWorkbookWriter.CopyRow` 的窄范围 NPOI row capability fallback |
| 生产 `.Result` | 0 | 未发现同步 Task 结果阻塞 |
| 生产 `.Wait()` | 0 | 未发现同步等待阻塞 |
| 生产 `Task.Run(...)` | 0 | 未引入伪异步包装；测试中的 3 处不计生产路径 |
| `TODO` | 0 | 当前扫描范围无 TODO 标记 |
| 生产 `InternalsVisibleTo` | 6 | 三个生产程序集各指向 `Bing.Offices.Tests` 和 `Bing.Offices.Tests.Integration`，无生产友元 |

`catch (Exception)` 命中 17 处，其中生产路径均有异常转换、清理诊断、取消过滤或边界记录语义；本项不按字符串命中直接判定缺陷，需结合调用上下文审查。Benchmark 的两个 `.Wait()` 仅用于启动闸门，测试的三个 `Task.Run` 用于并发/线程池行为验证，均不属于生产 API 候选。

### Public API 分类总账

`PublicApiContractTest` 当前对 183 个公开顶层类型建立唯一分类：

| 程序集 | User API | Provider SPI | Compatibility | Execution detail | 合计 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Bing.Offices.Abstractions` | 70 | 8 | 0 | 43 | 121 |
| `Bing.Offices.Core` | 0 | 0 | 10 | 51 | 61 |
| `Bing.Offices.Npoi` | 1 | 0 | 0 | 0 | 1 |
| **合计** | **71** | **8** | **10** | **94** | **183** |

上表的分类账以当前测试源码的字典条目为准；发布程序集的实际导出类型数量及成员快照仍由 API contract 和 snapshot artifact 另行核验。`Execution detail` 的 94 个类型按定义文件、生产/测试/文档/Benchmark 引用计数列入下方附录；计数格式为 `src/tests/docs/benchmarks`，不包含生成 XML 和排除目录。

## 逐符号引用与治理矩阵

### Obsolete、兼容层和删除候选

下表的引用计数格式为 `src/tests/docs/benchmarks`，并保留关键定义文件与实际生产桥接。`DI/反射` 已在同一全仓文本扫描中检查：除 `AddBingOfficesNpoi`、Profile 注册、默认 loader 和 plan factory provider 外，未发现候选通过字符串程序集扫描或反射创建形成未记录入口；API contract 的反射枚举列入测试引用计数。

| Candidate | 定义 | 引用计数 | 生产/测试/文档/Benchmark 关键证据 | 替代路径 | 删除风险与最终决定 |
| --- | --- | ---: | --- | --- | --- |
| `ICellValueConverter` | 已删除（原 `src/Bing.Offices.Abstractions/Bing/Offices/Conversions/ICellValueConverter.cs`） | 0/0/1/0 | Round 5 已移除 importer/materializer legacy text bridge、测试专用 converter 和接口文件；`docs/excel/nuget-migration.md` 仅保留迁移说明 | `IExcelValueConverter` | 已完成生产/测试/Docs/Benchmark/反射/API contract 负向扫描；**VERIFIED / 已删除** |
| `CsvHelper` DataTable 兼容类 | `src/Bing.Offices.Core/Bing/Offices/CsvHelper.cs:12` | 14/21/0/2 | 生产：类本身仍保留显式 DataTable API；测试：显式 delimiter/quote 专项；Benchmark：仍有 helper 场景 | `ICsvImporter`、`ICsvExporter` 与显式 delimiter/quote | 全局状态和旧隐式重载已删除；类本身是否继续收敛需后续 breaking approval。**PARTIAL / 授权子集已完成** |
| `CsvSeparatorCharacter` | 已删除（原 `CsvHelper.cs` 字段） | 0/0/0/0 | Round 5 删除全局可变 delimiter 状态及依赖它的旧重载；显式 delimiter 参数保留 | 含显式 delimiter 的 `CsvHelper` 重载 | **VERIFIED / 已删除**；避免跨请求共享可变格式状态 |
| `CsvQuoteCharacter` | 已删除（原 `CsvHelper.cs` 字段） | 0/0/0/0 | Round 5 删除全局可变 quote 状态及依赖它的旧重载；显式 quote 参数保留 | 含显式 quote 的 `CsvHelper` 重载 | **VERIFIED / 已删除** |
| `RequiredAttribute` | 已删除（原 `Attributes/Filters/RequiredAttribute.cs`） | 0/0/0/0 | 生产 validation binding/rule/type map、测试模型和 API contract 已迁移到 `ExcelRequiredAttribute` | `ExcelRequiredAttribute` | **VERIFIED / 已删除** |
| `RegexAttribute` | 已删除（原 `Attributes/Filters/RegexAttribute.cs`） | 0/0/0/0 | 生产 validation binding/rule/type map、Docs consumer 和 API contract 已迁移到 `ExcelRegexAttribute` | `ExcelRegexAttribute` | **VERIFIED / 已删除** |
| `RangeAttribute` | 已删除（原 `Attributes/Filters/RangeAttribute.cs`） | 0/0/0/0 | 生产 validation binding/rule、测试和 API contract 已迁移到 `ExcelRangeAttribute` | `ExcelRangeAttribute` | **VERIFIED / 已删除** |
| `MaxLengthAttribute` | 已删除（原 `Attributes/Filters/MaxLengthAttribute.cs`） | 0/0/0/0 | 生产 validation binding/rule/type map、测试和 API contract 已迁移到 `ExcelMaxLengthAttribute` | `ExcelMaxLengthAttribute` | **VERIFIED / 已删除** |
| `DateTimeAttribute` | 已删除（原 `Attributes/Filters/DateTimeAttribute.cs`） | 0/0/0/0 | 生产 validation binding/rule、测试和 API contract 已迁移到 `ExcelDateAttribute` | `ExcelDateAttribute` | **VERIFIED / 已删除** |
| `DuplicationAttribute` | 已删除（原 `Attributes/Filters/DuplicationAttribute.cs`） | 0/0/0/0 | 生产 validation binding/rule、NPOI column planning、测试和 API contract 已迁移到 `ExcelUniqueAttribute` | `ExcelUniqueAttribute` | **VERIFIED / 已删除** |
| `OfficeException` | `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeException.cs:9` | 17/5/0/0 | 生产：`ExcelTypeMapFactory` 非法配置、NPOI 合并跨 Sheet 错误；测试：`StreamPipelineTest`、API contract | 领域化配置/结构异常和 `ExcelImportError` | 基类仍有生产语义和派生类依赖；**BLOCKED / 先完成错误分类迁移** |
| `OfficeHeaderException` | `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeHeaderException.cs:7` | 9/2/0/0 | 生产定义和派生关系；当前无业务抛出点；测试/API contract 保留类型 | `ExcelImportError` header code | 可作为删除候选，但需先确认外部异常捕获和 API baseline；**CANDIDATE / 待批准** |
| `OfficeEmptyLineException` | `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeEmptyLineException.cs:7` | 7/2/0/0 | 仅定义/继承与 API contract；当前无生产抛出点 | `ExcelImportError` row/validation code | 可作为删除候选，但需迁移外部 catch 和 API contract；**CANDIDATE / 待批准** |
| `OfficeDataConvertException` | `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeDataConvertException.cs:7` | 3/2/0/0 | 仅定义/继承与 API contract；当前无生产抛出点 | `ExcelImportError` conversion code | 可作为删除候选，但需迁移外部 catch 和 API contract；**CANDIDATE / 待批准** |
| `ExcelSetting` | 已删除（原 `src/Bing.Offices.Abstractions/Bing/Offices/Settings/ExcelSetting.cs`） | 0/0/0/0 | Round 6 已删除旧设置类型，并同步 API contract 与迁移文档 | Workbook request/builder options | **VERIFIED / 已删除**；无生产调用，受控引用已完成迁移。 |
| `SheetSetting` | 已删除（原 `src/Bing.Offices.Abstractions/Bing/Offices/Settings/SheetSetting.cs`） | 0/0/0/0 | Round 6 已删除旧设置类型，并同步 API contract | Sheet request/builder options | **VERIFIED / 已删除**；无生产调用，受控引用已完成迁移。 |
| `ExcelMappingDocumentFactory` request overload | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocumentFactory.cs:8` | 2/4/0/0 | 生产：`ExcelMappingPlanFactory`；测试：request configuration merge 和 API contract | 统一 Plan 编译入口或 internal clone | 当前语义已修复且有生产调用；**VERIFIED / 保留至 API 收敛决策** |
| `AddNpoi` | 已无源码定义 | 0/0/0/0 | 全仓受控扫描无旧名称残留；新入口由 NPOI 扩展提供 | `AddBingOfficesNpoi` | **VERIFIED / 已删除**；API exact baseline 和 package consumer 使用新入口 |
| `MaxBytes` | 已无源码定义 | 0/0/1/0 | 仅 `docs/excel/nuget-migration.md` 的迁移说明保留旧名称；生产/测试无残留 | `MaxSerializedBytes` | **VERIFIED / 已删除**；文档残留是迁移历史证据，不是兼容成员 |

### Provider SPI、execution detail 与合法 fallback

| Candidate group | 定义/范围 | 引用矩阵 | 替代路径与治理裁决 |
| --- | --- | --- | --- |
| `IExcelDynamicMappingColumn`、`IExcelMappingLayout`、`IExcelMappingColumn`、`IExcelMappingPlan`、`IExcelMappingPlanFactory`、`IExcelMappingSheetPlan`、`IExcelMappingStyle`、`IExcelMappingWorkbookPlan` | `src/Bing.Offices.Abstractions/Bing/Offices/Providers/`；8 个 Provider SPI | 8 个均已列入 `PublicApiContractTest`；`IExcelMappingPlan`/`IExcelMappingWorkbookPlan` 由 NPOI 实现和测试引用；其余为契约边界，未发现生产 IVT | 保留 public + `EditorBrowsable(Never)`，因为它们定义 Provider 计划/样式/动态列契约；必须维持 API contract 和 provider migration，不可伪装成普通 User API |
| `UniqueTracker` | `src/Bing.Offices.Abstractions/Bing/Offices/Providers/UniqueTracker.cs:11` | 8/6/0/2；生产 CSV/Excel/NPOI 三条主链；Benchmark `MappingValidationBenchmarks`、`Program` | 当前由多模块和 Benchmark 直接使用，保留为跨模块执行协作类型；后续需在 API 批准阶段选择正式 SPI 或迁入 Core internal，当前 **BLOCKED / 不删除** |
| 94 个 `Execution detail` 顶层类型 | 定义和计数见下方完整附录；包括 mapping merger/cloner、loader、plan/type/value map、CSV concrete、extensions、metadata、styles、validation rule 等 | 每项均有 `src/tests/docs/benchmarks` 计数；没有任何项目通过生产 IVT 获取其内部成员 | 当前不自动 internal 化。按职责和外部包消费逐项迁移；统一前置条件为：替代 User API/SPI、源码/二进制 breaking approval、API negative scan、Release build/test/pack/consumer。**BLOCKED / 待逐项批准** |
| `NpoiFailureWorkbookWriter` 三处 `NotImplementedException` catch | `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs:271,278,285` | 生产 3；测试通过 Failure Workbook 专项；无 docs/Benchmark 引用 | NPOI 2.7.4 HSSF 的 `IRow.Hidden`/`IRow.Collapsed` 明确不支持、`ZeroHeight` 支持；XSSF 三项均支持。catch 仅包围对应属性访问，不覆盖序列化/复制/主异常。**VERIFIED / 合法 capability fallback，不删除** |

## Execution detail 完整逐符号附录

下表覆盖 `ApiTypeCategories` 中全部 94 个 `Execution detail` 类型。引用计数格式为 `src/tests/docs/benchmarks`；`-` 表示类型名在源码中由泛型反射名表达，定义证据由所在构造文件和 API contract 反射提供。所有条目的统一裁决是“当前证据已完成扫描，但不等于已批准删除”。

| Symbol | 定义证据 | 引用计数 |
| --- | --- | ---: |
| `Bing.Offices.Abstractions:Bing.Offices.Attributes.DecoratorAttributeBase` | `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/DecoratorAttributeBase.cs:6` | 4/2/1/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Attributes.FilterAttributeBase` | `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/FilterAttributeBase.cs:6` | 37/15/1/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Attributes.BindFilterAttribute` | `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/BindFilterAttribute.cs:7` | 4/3/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelColumnConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelColumnConfiguration.cs:6` | 46/52/0/6 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingConfiguration.cs:6` | 81/65/0/15 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingConfigurationMerger` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingConfigurationMerger.cs:7` | 6/8/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDiagnostic` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDiagnostic.cs:6` | 16/2/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocument` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocument.cs:8` | 86/58/3/10 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocumentFactory` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocumentFactory.cs:8` | 2/4/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicColumnConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDynamicColumnConfiguration.cs:10` | 27/19/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicValidationConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDynamicValidationConfiguration.cs:6` | 23/11/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelDynamicColumnMergeMode` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelDynamicColumnMergeMode.cs:6` | 4/4/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingLayoutConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingLayoutConfiguration.cs:6` | 12/6/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingStyleConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingStyleConfiguration.cs:6` | 14/7/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValidationRuleMergeMode` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelValidationRuleMergeMode.cs:6` | 6/3/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingMergeMode` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelValueMappingMergeMode.cs:6` | 8/5/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingConfiguration` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelValueMappingConfiguration.cs:6` | 11/6/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportColumnMappingBuilder`2` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExportMappingBuilder.cs:45` | 10/5/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportMappingBuilder`1` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExportMappingBuilder.cs:10` | 9/5/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.FluentSetting`2` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/FluentSetting.cs:8` | 9/13/1/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportColumnMappingBuilder`2` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ImportMappingBuilder.cs:51` | 16/6/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportMappingBuilder`1` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ImportMappingBuilder.cs:10` | 9/7/0/3 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingDirection` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingDirection.cs:6` | 62/87/1/17 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingProfileRegistry` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingProfileRegistry.cs:6` | 5/4/0/1 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingSourceKind` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingSourceKind.cs:6` | 25/9/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelModelAliasRegistry` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelModelAliasRegistry.cs:10` | 15/4/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.ProfileDescriptor` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ProfileDescriptor.cs:8` | 22/6/0/1 |
| `Bing.Offices.Abstractions:Bing.Offices.Configurations.IExcelMappingConfigurationLoader` | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/IExcelMappingConfigurationLoader.cs:6` | 3/6/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.ExcelFormat` | `src/Bing.Offices.Abstractions/Bing/Offices/ExcelFormat.cs:8` | 12/58/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Providers.UniqueTracker` | `src/Bing.Offices.Abstractions/Bing/Offices/Providers/UniqueTracker.cs:11` | 8/6/0/2 |
| `Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelUnsupportedFeaturePolicy` | `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs:193` | 11/4/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Imports.ValidateMode` | `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ValidateMode.cs:6` | 21/9/1/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWhitespacePolicy` | `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs:20` | 33/6/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Settings.ExcelSetting` | `src/Bing.Offices.Abstractions/Bing/Offices/Settings/ExcelSetting.cs:6` | 1/2/5/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Settings.SheetSetting` | `src/Bing.Offices.Abstractions/Bing/Offices/Settings/SheetSetting.cs:6` | 1/2/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderLineStyle` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:43` | 9/7/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderStyle` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:81` | 6/7/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyle` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:157` | 25/27/0/1 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyleReset` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:97` | 2/10/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelColor` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:9` | 10/12/0/1 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelFillPattern` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:32` | 7/7/0/1 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelHorizontalAlignment` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:57` | 9/3/0/0 |
| `Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelVerticalAlignment` | `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs:70` | 7/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ColumnNameAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/ColumnNameAttribute.cs:7` | 4/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.DataFormatAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/DataFormatAttribute.cs:7` | 6/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.DecimalScaleAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/DecimalScaleAttribute.cs:7` | 4/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.DynamicColumnAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/DynamicColumnAttribute.cs:10` | 3/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelDateAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelDateAttribute.cs:9` | 9/3/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelIgnoreAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/ExcelIgnoreAttribute.cs:7` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxLengthAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelMaxLengthAttribute.cs:9` | 8/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxValueAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelMaxValueAttribute.cs:9` | 7/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelRangeAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelRangeAttribute.cs:9` | 9/3/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelRegexAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelRegexAttribute.cs:9` | 10/2/0/1 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelRequiredAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelRequiredAttribute.cs:9` | 6/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ExcelUniqueAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelUniqueAttribute.cs:9` | 8/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.HeaderAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Decorators/HeaderAttribute.cs:10` | 4/3/0/2 |
| `Bing.Offices.Core:Bing.Offices.Attributes.MergeColumnsAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Decorators/MergeColumnsAttribute.cs:10` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.ValueMappingAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/ValueMappingAttribute.cs:7` | 4/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Attributes.WrapTextAttribute` | `src/Bing.Offices.Core/Bing/Offices/Attributes/Decorators/WrapTextAttribute.cs:10` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Configurations.ExcelColumnMappingBuilder`2` | `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs:91` | 12/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Configurations.ExcelMapping` | `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs:9` | 1/10/2/0 |
| `Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingBuilder`1` | `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs:23` | 5/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingConfigurationLoader` | `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs:16` | 7/60/4/6 |
| `Bing.Offices.Core:Bing.Offices.Configurations.DefaultExcelMappingConfigurationLoader` | `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs:504` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Csv.CsvEntityExporter` | `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs:19` | 4/16/0/0 |
| `Bing.Offices.Core:Bing.Offices.Csv.CsvEntityImporter` | `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs:223` | 4/34/1/0 |
| `Bing.Offices.Core:Bing.Offices.CsvHelper` | `src/Bing.Offices.Core/Bing/Offices/CsvHelper.cs:12` | 14/21/0/2 |
| `Bing.Offices.Core:Bing.Offices.Extensions.ExpressionExtension` | `src/Bing.Offices.Core/Bing/Offices/Extensions/ExpressionExtension.cs:9` | 1/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions` | `src/Bing.Offices.Core/Bing/Offices/Extensions/CsvStreamExtensions.cs:9` | 1/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions` | `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelStreamExtensions.cs:10` | 1/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Extensions.PropertyInfoExtensions` | `src/Bing.Offices.Core/Bing/Offices/Extensions/PropertyInfoExtensions.cs:10` | 1/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Extensions.TypeExtensions` | `src/Bing.Offices.Core/Bing/Offices/Extensions/TypeExtensions.cs:10` | 1/23/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelPropertyMap` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelPropertyMap.cs:9` | 17/3/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactory` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs:21` | 5/19/1/10 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactoryProvider.cs:14` | 3/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelTypeMap`1` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelTypeMap.cs:7` | 12/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelTypeMapFactory` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelTypeMapFactory.cs:20` | 2/28/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelValidationBindingFactory` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelValidationBindingFactory.cs:9` | 1/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueMap`1` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelValueMap.cs:9` | 3/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueConverterBindingResolver` | `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelValueConverterBindingResolver.cs:13` | 3/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Metadata.MergedRegionInfo` | `src/Bing.Offices.Core/Bing/Offices/Metadata/MergedRegionInfo.cs:6` | 11/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Metadata.PictureInfo` | `src/Bing.Offices.Core/Bing/Offices/Metadata/PictureInfo.cs:6` | 23/3/0/0 |
| `Bing.Offices.Core:Bing.Offices.Metadata.PictureStyle` | `src/Bing.Offices.Core/Bing/Offices/Metadata/PictureStyle.cs:6` | 12/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.RegexConst` | `src/Bing.Offices.Core/RegexConst.cs:6` | 1/6/0/0 |
| `Bing.Offices.Core:Bing.Offices.Styles.Color` | `src/Bing.Offices.Core/Bing/Offices/Styles/Color.cs:6` | 64/10/0/1 |
| `Bing.Offices.Core:Bing.Offices.Validations.DateTimeExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:150` | 2/3/0/0 |
| `Bing.Offices.Core:Bing.Offices.Validations.DuplicationExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:188` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Validations.ExcelValidationRules` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:11` | 5/3/1/0 |
| `Bing.Offices.Core:Bing.Offices.Validations.MaxLengthExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:131` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Validations.MaxValueExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:113` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Validations.RangeExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:92` | 2/3/0/0 |
| `Bing.Offices.Core:Bing.Offices.Validations.RegexExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:45` | 2/4/0/1 |
| `Bing.Offices.Core:Bing.Offices.Validations.RequiredExcelValidationRule` | `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs:31` | 2/2/0/0 |
| `Bing.Offices.Core:Bing.Offices.Extensions.MappingProfileServiceCollectionExtensions` | `src/Bing.Offices.Core/Bing/Offices/Extensions/MappingProfileServiceCollectionExtensions.cs:14` | 1/2/0/0 |

## P0-03 最终验收结论

- **扫描：`VERIFIED`**。范围、编码、排除规则、特殊构造、IVT、API 分类和 94 个 execution detail 类型均已有可复核记录。
- **删除：按授权子集完成。** Round 5 已删除 `ICellValueConverter`、6 个旧 validation attributes、CSV 全局 separator/quote 和旧隐式 DataTable 重载；Office exceptions、Settings、public execution detail 等未授权候选仍保持 `BLOCKED/CANDIDATE`。
- **API 分层：`PARTIAL/BLOCKED`。** 分类账已完整，但 94 个 execution detail 和 10 个 compatibility 类型仍需后续逐项批准、迁移、API negative scan、build/test/pack/consumer 闭环。
- **发布：`No-Go`。** 正式 API baseline、缺失 TFM runtime、性能预算、Failure Workbook 双 DOM 资源边界和 package consumer 长路径限制仍独立阻断 RC。

## 逐符号证据与当前裁决

| Candidate | 定义证据 | 生产引用 | 测试/Docs/Consumer 引用 | 替代路径 | 当前裁决 |
| --- | --- | --- | --- | --- | --- |
| `ICellValueConverter` | 已删除（原 `Conversions/ICellValueConverter.cs`） | 当前生产无引用，legacy text bridge 已移除 | 测试专用 converter 已删除；迁移文档只保留替代路径 | `IExcelValueConverter` | `VERIFIED / 已删除`；Round 5 快照删除 2 个 Abstractions member lines |
| `CsvHelper` DataTable 兼容层与全局 `CsvSeparatorCharacter`/`CsvQuoteCharacter` | `src/Bing.Offices.Core/Bing/Offices/CsvHelper.cs`；类保留，旧全局状态和隐式重载已删除 | 实体管线继续使用 provider-neutral `CsvEntityPipeline`；显式 DataTable API 保留 | 显式 delimiter/quote 回归 `6/6`；旧字段/重载无当前源码引用 | `ICsvImporter`/`ICsvExporter` 与显式格式参数 | `PARTIAL / 授权子集已完成`；DataTable 显式 API 未删除，剩余类是否收敛需后续批准 |
| `Required/Regex/Range/MaxLength/DateTime/Duplication` attributes | 已删除（原 `Attributes/Filters/*` 六个文件） | validation binding/rule/type map 已只识别 `Excel*` attributes | 测试模型、CsvTest、Docs consumer 和 API contract 已迁移；package-only consumer 编译 `ExcelRequired` | `ExcelRequired`、`ExcelRegex`、`ExcelRange`、`ExcelMaxLength`、`ExcelDate`、`ExcelUnique` | `VERIFIED / 已删除`；Round 5 快照删除 28 个 Core member lines |
| `OfficeException`、`OfficeHeaderException`、`OfficeEmptyLineException`、`OfficeDataConvertException` | `src/Bing.Offices.Core/Bing/Offices/Exceptions/*` | `OfficeException` 仍用于类型映射非法配置和 NPOI 合并跨 Sheet 错误；其余三类当前无生产调用 | `StreamPipelineTest` 断言 `OfficeException`；API contract 仍列四类；无 Docs/consumer 使用证据 | 标准参数/配置异常和结构化 `ExcelImportError` | `PARTIAL`。四类不能整体删除；先将无调用的三个列为删除候选，`OfficeException` 需先完成错误分类迁移。 |
| `ExcelMappingDocumentFactory` request overload | `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocumentFactory.cs` | `ExcelMappingPlanFactory` 调用三参数 overload 时可传 request configuration；当前源码另有 request-facing builder/document 调用链 | `ReviewFixRegressionTest.MappingDocumentFactory_RequestConfiguration_ShouldMergeOnlySelectedDirection`；API contract 反射检查 | internal clone 或统一 Plan 编译入口 | `DONE / 语义已修复`。当前 request 配置已被合并；是否 internal 化留待 API 批准。 |
| `ExcelSetting` / `SheetSetting` | `src/Bing.Offices.Abstractions/Bing/Offices/Settings/*` | 当前源码无生产引用 | 当前任务和历史 docs 仅有候选说明；package consumer 未使用 | Workbook/Sheet Request Builder 与 metadata options | `CANDIDATE / 可删除但未执行`。无当前代码消费者；为遵守删除安全边界和正式 API 审批，暂不删除文件。 |
| `AddNpoi` | `src/Bing.Offices.Npoi/Extensions/Extensions.Service.cs` | 无旧名称生产调用 | 全仓调用已迁移；exact API、Integration、Docs、package consumer 使用新入口 | `AddBingOfficesNpoi`，返回 `IServiceCollection` | `VERIFIED / 已删除`。没有 wrapper/forwarder/obsolete 保留。正式 NPOI hash 待批准更新。 |
| 生产 `NotImplementedException` 路径 | `NpoiFailureWorkbookWriter.CopyRow` 的 row capability catches | 只捕获 NPOI 属性访问不支持；不是向调用方抛出占位异常 | 直接运行 NPOI 2.7.4 HSSF/XSSF provider 验证；生产搜索仅剩三处 | 保持 capability fallback，不扩展为 catch-all | `VERIFIED / 合法 capability fallback`。HSSF 的 `Hidden`/`Collapsed` 明确不支持，`ZeroHeight` 支持；XSSF 三项均支持。三个 catch 均只包围相应属性访问，主序列化异常仍由外层保留。 |
| public execution detail 类型 | `PublicApiContractTest` 分类账；Abstractions/Core 仍大量 exported types | 多个仍被 Core/Npoi 主链使用；部分由 Docs/consumer 直接使用（loader、CSV concrete） | API snapshot 与 package consumer 已确认实际可消费 API | 只保留 User API/Provider SPI；其余 internal 化并迁移调用者 | `PARTIAL / 待逐项批准`。不能以 `EditorBrowsable` 或 hash 测试替代治理。 |

## 已完成的 API/语义收敛

- `AddNpoi` 已迁移为 `AddBingOfficesNpoi`，返回同一 `IServiceCollection`；全仓生产、测试、Docs、Benchmark 和 package consumer 无旧入口残留。
- 失败工作簿 `MaxBytes` 已迁移为 `MaxSerializedBytes`；旧名称在生产、测试和公开文档中无残留。该属性只限制序列化输出，不限制 DOM/实体峰值。
- `ExcelMappingDocumentFactory` 请求配置已按选定方向合并，非目标方向保留独立快照；当前不再存在“接收参数但忽略”的行为。

## Round 5 删除后验收

- `ICellValueConverter`：已完成生产 bridge、测试专用实现、API contract 和 package consumer 迁移；Round 5 Abstractions 快照删除 2 个 member lines。
- 六个旧 validation attributes：已完成生产识别分支、测试模型、Docs consumer、API contract、XML 和 package consumer 迁移；Round 5 Core 快照删除对应类型、构造函数和属性。
- `CsvSeparatorCharacter`、`CsvQuoteCharacter` 及旧隐式 DataTable 重载：已删除；显式 delimiter/quote API 保留，并由 CSV 专项 `6/6` 验证。
- Round 5 负向扫描：`src/**/*.cs` 无 `[Obsolete]`/`ObsoleteAttribute`，生产/测试/Benchmark/README 无精确旧符号使用；任务历史报告和迁移表中的旧名称仅作为历史证据保留。
- Round 5 Release build、net6/net8 Unit、Integration、Docs、pack 和 package-only consumer 均已重新取证；正式 API hash mismatch 仍保留为 No-Go。

每个后续真正删除的符号仍必须完成：定义、编译引用、`nameof`/字符串/反射、DI/程序集扫描、测试/Benchmark、README/docs/sample、替代 API、源码删除、API contract、Release build/test/pack/consumer。未获本轮授权的 Office exceptions、Settings、UniqueTracker、public execution detail 和其它 compatibility 候选不得据此视为已删除。
