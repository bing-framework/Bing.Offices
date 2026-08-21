<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: bing-offices-excel-import-export-enhancement-v2
AI_EXECUTION_FINISHED_AT: 2026-08-21T13:50:03.5993942+08:00

# 实施执行报告

## 执行结论

本轮按已批准的 `plan.md` 持续实施，完成了 P0 样式/模板/结构错误修复、P1 策略契约基础、导出宽度和批注、导入选择/范围/规范化、图片/资源限制/失败工作簿基础，以及 Workbook Data Validation 首版支持。显式列表规则已针对 NPOI 重开后的 `Formula1`/转义引号表示修复，Unit、Integration、Release build、pack、Public API 顶层 approval 和 benchmark smoke 均已验证。

状态为 `PARTIAL`：计划中的完整不可变 Column Plan 拆分、ErrorRowsOnly 保真复制、完整关系/唯一性 API 收敛、完整成员级 Public API approval、Docs Consumer、Office/LibreOffice 互操作和格式基线治理仍未完成或不可验证。未修改 `plan.md`，未自动执行 Git 提交、推送、Tag、PR 或发布。

## 任务信息

- Task ID：`bing-offices-excel-import-export-enhancement-v2`
- 执行器：Copilot / `plan-executor`
- 执行开始：`2026-08-20T13:00:27.204Z`
- 执行结束：`2026-08-20T21:54:12.1702739+08:00`
- 工作树：执行前已有大量未提交改动，本轮保留并在其上继续实施；未执行 reset、clean 或回退陌生改动。

## 计划执行情况

| Task | 状态 | 证据/说明 |
| --- | --- | --- |
| P0-001 | COMPLETED | locked restore 初次因锁漂移失败；按计划使用 `--force-evaluate` 后重新 `--locked-mode` 成功。记录了构建、测试和现有 API 基线。 |
| P0-002 | COMPLETED | `NpoiStyleCache` 支持 base + overlay 逐属性合成；修复 Body/Dynamic-only 样式、模板 origin、图表列偏移和结构错误结果化。 |
| P1-001 | COMPLETED | 增加 Selector、Read Range、Whitespace、Width、Failure、Validation、Image、Comment、Resource Limit 等 provider-neutral 类型并接入 Workbook Request。 |
| P1-002 | PARTIAL | 固定/动态列的转换、校验、动态目标、重复校验已有接线；完整不可变 Column Plan 编译器和所有扩展点统一仍待后续。 |
| P2-001 | COMPLETED | Fixed/AutoFit/Adaptive 宽度基础与样式 overlay 已进入导出主链。 |
| P2-002 | PARTIAL | 自定义表头偏移和传统 Note 批注已接入；完整布局编译、冲突矩阵和全部批注场景仍待补齐。 |
| P3-001/P3-002 | COMPLETED | Sheet Selector、Read Columns、名称比较、Header/Body Whitespace 已接入并通过现有回归。 |
| P4-001 | PARTIAL | 固定/动态基础唯一性、忽略和规范化链路已接入；完整 Request API、首行来源和泛型关系 Key 尚未完全收敛。 |
| P5-001/P5-002 | PARTIAL | 结构化错误、输入资源限制和 AnnotatedOriginal 基础已实现；完整 Error Model、统计、ErrorRowsOnly 保真复制仍待完成。 |
| P5-003 | NOT_DONE | 未完成 ErrorRowsOnly 的原始 CellType/公式/图片/合并复制契约。 |
| P6-001 | PARTIAL | Workbook Rules 首版支持并已验证显式列表；整数/小数/文本长度等实现存在，完整支持矩阵、Unsupported 策略、HSSF/组合覆盖仍待补齐。 |
| P7-001 | PARTIAL | 图片索引、byte[] 基础绑定和输入资源上限已接入；完整 multiplicity、像素/格式限制和失败文件图片复制仍待补齐。 |
| P8-001/P8-002 | PARTIAL | 顶层 Public API approval 已更新，NPOI 兼容层和大文件职责拆分尚未完成。 |
| P9-001 | NOT_DONE | 仓库没有 Docs Consumer 项目，本轮未新增。 |

## 已完成事项

- 样式缓存从整对象替换改为基础样式加 overlay 的逐属性合成，保留 NumberFormat、模板字体、填充和边框等未覆盖属性。
- 修复仅配置 BodyStyle 或 DynamicStyle 时被提前跳过的问题。
- 统一模板 Header、Merge、Chart 的列 origin，覆盖非 A1 模板回归场景。
- 导出动态列 Converter 接线；导入动态 Target、Converter、Validator 和 Duplication 基础接线。
- 导入按名称/下标选择、名称比较、指定读取列范围、Header/Body 空白策略已接入主链。
- 导出 Fixed/AutoFit/Adaptive 列宽基础、传统 Header Note 和批注冲突策略已接入。
- 图片锚点索引和 `ExcelImageData` provider-neutral DTO 已接入；图片锚点行不会再因空文本被跳过。
- 输入资源大小限制和 AnnotatedOriginal 失败工作簿基础输出已接入，保留原值并生成 `_ImportErrors`。
- Workbook Data Validation 首版支持 `LIST`、部分数值/文本长度规则；显式列表会兼容 NPOI 重开后 `ExplicitListValues`、`Formula1`、XML/反斜杠转义引号。
- Public API 顶层 approval 新增 `ExcelColumnWidthOptions`、`ExcelComment`、`ExcelImageData`。

## 部分/未完成事项

- 统一 Column Plan 仍是目标状态，Importer/Exporter 内仍存在较大的编排类，尚未按 Planning/Reading/Writing/Validation/Failures/Images 完整拆分。
- ErrorRowsOnly 尚未完成原始行保真复制；失败工作簿资源、取消和图片/合并保真覆盖不足。
- 完整 Unique/Ignore/Relation 泛型 Key 契约、动态列等价性和错误来源坐标仍需补测/收敛。
- Public API approval 当前仍以公开顶层类型为主，尚未升级到全成员签名快照；计划中列出的 legacy API internal 化/删除尚未全部执行。
- Docs Consumer 项目不存在；真实 Excel/LibreOffice 互操作没有可验证环境和产物，因此不宣称通过。
- `dotnet format --verify-no-changes --no-restore` 未通过。当前日志包含历史导入排序 `IMPORTS` 和 `IDE1006` 命名规则问题，涉及旧文件及部分本轮触及文件；为避免无关大范围格式化，本轮未自动改写这些文件。

## 修改文件

本轮直接修改的收口文件：

- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`：显式列表 Data Validation 的 NPOI Formula1/转义兼容。
- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`：保留结构化断言并验证显式列表行为。
- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`：补充三个 provider-neutral 顶层类型 approval。

本轮执行前已存在并继续保留的计划相关改动还包括 Abstractions 的 Import/Export policy、Comment、Image DTO，以及 NPOI Exporter/StyleCache 等文件；这些改动未被回退或覆盖。

## API/数据/配置变化

- 新增/接入的公共策略包括列宽、批注、Sheet Selector、读取范围、名称比较、空白规范化、失败工作簿、Workbook Validation、图片多重性和资源限制。
- 新增公共 DTO 不引用 `NPOI.*`；NPOI 具体 Workbook/Constraint/Picture 类型仍位于适配器实现路径。
- Data Validation 的显式常量列表比较只接受解析后的精确字符串，不会把不支持的公式规则静默当作通过。
- 未执行数据库、网络、生产基础设施或破坏性数据操作。

## 测试结果

- Unit net8：`113/113` 通过。
- Unit net6：`113/113` 通过。
- Integration net8：`10/10` 通过。
- Integration net6：`10/10` 通过。
- Workbook 显式列表定向测试：`1/1` 通过。
- 最后一次严格 `Assert.Single` 断言复核：`1/1` 通过。
- Public API 顶层 approval：通过。
- 编辑器静态诊断：`src`、`tests`、`benchmarks` 无错误。

关键回归方法：

- `ExcelP0RegressionTest.Import_WorkbookExplicitListValidation_ShouldReportInvalidCell`
- `ExcelP0RegressionTest.Import_AnnotatedFailureWorkbook_ShouldAnnotateOriginal`
- `ExcelP0RegressionTest.Import_ImageAnchoredCell_ShouldBindBytes`
- `ExcelWorkbookRequestTest.Export_TemplateRegion_ShouldPreserveTemplateAndKeepInputOpenWhenRequested`
- `ExcelWorkbookRequestTest.Export_RequestStyle_ShouldReuseStyleAndWriteCustomXlsxColor`
- `PublicApiContractTest.PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline`

## Build/Typecheck/Lint/Format

- `dotnet build .\Bing.Offices.sln -c Release --no-restore`：通过；存在 netcoreapp3.1 包支持警告、legacy `ICellValueConverter` 警告和 XML 参数注释警告。
- `dotnet pack .\Bing.Offices.sln -c Release --no-build --no-restore`：通过。
- `dotnet format .\Bing.Offices.sln --verify-no-changes --no-restore`：未通过，见“部分/未完成事项”；未执行自动格式化。
- `git diff --check`：通过；Git 仅报告工作树 CRLF/LF 提示，无空白错误。
- Benchmark smoke：通过并清理产物。`StreamPipelineBenchmarks` 观察值包括 Import InProcess `11.154 ms`、Export InProcess `8.269 ms`、ShortRun Import `11.268 ms`、ShortRun Export `13.049 ms`；仅作为当前环境基线，不作为新功能性能阈值。

## 计划偏差

- 计划要求的完整 API Breaking cut 和大文件拆分未在本轮强行完成，因为当前工作树已包含大量重叠未提交改动，继续删除/移动会扩大不可逆 review 面并可能覆盖用户进行中的工作。
- 计划定义的 `ErrorRowsOnly`、Docs Consumer、真实 Office/LibreOffice 互操作需要额外实现或外部工具/样本，当前环境不足以给出通过证据。
- Data Validation 发现 NPOI 2.7.4 重开显式列表的表示与创建时不同，按根因增加了兼容解析和直接回归测试；未修改计划中的支持矩阵声明。

## 基线问题

- locked restore 初次因锁文件漂移失败；使用一次 `--force-evaluate` 后重新 locked restore 成功。
- 工作树在执行开始前已存在大量未提交改动；本轮未尝试获得 clean workspace，也未回退陌生文件。
- `dotnet format` 的导入排序/命名规则错误是当前仓库格式基线问题，不能据此声称格式验收通过。

## 已知问题

- 构建仍有 `ICellValueConverter` 过时警告和 XML 参数注释警告，未在本轮扩大范围清理。
- netcoreapp3.1 依赖存在包不支持警告；构建通过，但旧运行时不作为本轮运行时兼容性证明。
- Workbook Validation 对 Named Range、跨 Sheet、相对/自定义公式等能力仍按计划保持未承诺状态。

## 风险与回归关注点

- NPOI 重写 XLSX 高级部件可能影响批注、图片、验证、图表和模板保真；当前没有 Office/LibreOffice 人工互操作证据。
- 失败工作簿仍需补充 ErrorRowsOnly 的原始 CellType、公式、合并、验证和图片复制测试。
- 公共 API 仍存在历史类型面，成员签名 approval 未完成，后续 Breaking 收敛需要单独 review。

## Reviewer 注意事项

- 优先检查 `NpoiExcelImporter.ValidateWorkbookValue` 对 NPOI 显式列表字符串的规范化是否只覆盖常量列表，避免误把公式列表当作常量列表。
- 检查所有失败工作簿输出是否遵守调用方流所有权、大小限制和取消语义。
- 检查固定列和动态列在 Converter、Validator、Unique、Ignore、Image、Comment、Width 上是否真正共享同一执行计划，而非存在旁路逻辑。
- 独立 Review 应继续把本报告标记为 `PARTIAL`，不要把顶层 API approval 或已有组合测试解释为完整成员级 API/互操作验收。

## Git 状态

- 仅执行了 `git status`、`git diff`、`git diff --check` 等只读 Git 命令。
- 未执行 `git add`、`git commit`、`git push`、`git reset`、`git clean`、Tag、PR 或发布。

## Review 修复记录

本轮依据独立 Review 的 `FIX-001` 至 `FIX-006` 执行，未修改 `review.md` 或 `plan.md`，未执行提交、推送或发布。

### FIX-001

- 严重程度：HIGH；处理要求：MUST_FIX；状态：COMPLETED
- 根因：`MaxRows`、`MaxErrors`、`MaxImageBytes` 只存在于配置对象，未进入逐行、错误收集和图片索引主链。
- 修复：导入执行选项接入 Workbook 资源限制；逐行达到 `MaxRows` 时生成 `ResourceLimit`；达到 `MaxErrors` 时停止后续收集并生成截断错误；图片索引同时执行单图/累计字节和数量限制；输入、图片和失败输出继续检查取消及流所有权。
- 相关文件：`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportErrorCode.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`src/Bing.Offices.Npoi/Imports/ExcelImportExecutionOptions.cs`、`tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`。
- 验证：`Import_WorkbookResourceLimits_ShouldStopRowsAndErrors` 通过；Unit net8 `114/114` 通过。

### FIX-002

- 严重程度：HIGH；处理要求：MUST_FIX；状态：COMPLETED
- 根因：NPOI Importer/Exporter、Helper 和低层扩展处于 public，消费者可直接依赖 NPOI 类型。
- 修复：NPOI 实现、Helper、Workbook/Sheet/Cell 等低层扩展改为 `internal`；通过 `AddNpoi()` 注册 `IExcelImporter`/`IExcelExporter`；测试使用 `InternalsVisibleTo`，不扩大生产 API；Public API approval 删除实现类型。
- 相关文件：`src/Bing.Offices.Npoi/AssemblyInfo.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`、`src/Bing.Offices.Npoi/ExcelHelper.cs`、`src/Bing.Offices.Npoi/Extensions/*`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`。
- 验证：Public API approval 通过；Integration net8 `10/10` 通过；新增 provider-neutral 成员签名扫描。

### FIX-003

- 严重程度：HIGH；处理要求：MUST_FIX；状态：COMPLETED
- 根因：`ErrorRowsOnly` 只在原 Workbook 上删除非错误行，未剔除无错误 Sheet，也未输出来源和聚合错误字段。
- 修复：按 Sheet 和源行去重，仅保留有错误的 Sheet/失败行；直接复用原始 NPOI 行和单元格对象，保留原 CellType、公式、样式和格式；追加 `__SourceRow`、`__Errors` 字段并聚合同一行错误；继续执行输出大小和取消检查。
- 相关文件：`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`。
- 验证：现有 Annotated/Failure 回归通过；Unit net8 `114/114`、Integration net8 `10/10` 通过。

### FIX-004

- 严重程度：HIGH；处理要求：MUST_FIX；状态：COMPLETED
- 根因：原生 Data Validation 查询逐 Cell 遍历全部规则，Unsupported Policy 未接线，日期/时间和图片多重性缺少执行路径。
- 修复：新增 Workbook 级 `UnsupportedFeaturePolicy`；支持日期/时间规则；不支持规则按 `Report/Fail` 分流；新增 `ValidationRangeIndex`，按 Sheet 预索引坐标；动态图片列支持 First/All/Fail，图片绑定支持单值和集合目标；图片数量、单图和累计字节限制执行。
- 相关文件：`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelWorkbookImportRequest.cs`、`src/Bing.Offices.Npoi/Imports/ValidationRangeIndex.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`。
- 验证：P0 Validation/Image 回归通过；Unit net8 `114/114`、Integration net8 `10/10` 通过。

### FIX-005

- 严重程度：HIGH；处理要求：MUST_FIX；状态：PARTIAL
- 根因：固定列和动态列的列元数据仍由 Importer/Exporter 各自维护，缺少共享不可变计划。
- 修复：新增 internal immutable `ExcelColumnPlan`，Importer 的 fixed/dynamic `ImportColumn` 统一继承并共享标题、属性、动态定义和物理位置元数据；保留现有请求级映射隔离和动态转换/校验入口。
- 相关文件：`src/Bing.Offices.Npoi/ExcelColumnPlan.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`。
- 未完成：Exporter 尚未完全改为消费同一计划，完整 Reader/Writer/Layout/Validation/Unique/Image 单管线和直接 internal Plan 测试仍需下一轮；因此本 FIX 不宣称完全闭合。
- 验证：编译通过；Unit net8 `114/114`、Integration net8 `10/10` 通过。

### FIX-006

- 严重程度：HIGH；处理要求：MUST_FIX；状态：COMPLETED
- 根因：旧 `MaxColumnLength`、`HeaderMappings`、legacy `ExcelImportResult<T>` 和公开 `Sheet.Items` 与 Workbook Request 结果面并存；关系 Key 固定为 string。
- 修复：删除生产 `MaxColumnLength()`/`HeaderMappings()`；删除公开 legacy 结果类型和 `ExcelSheetImportResult.Items`；测试兼容适配器仅在测试程序集内将旧 HeaderMappings 转译为 Mapping Overlay；`HasMany<TParent,TChild,TKey>` 支持 `IEqualityComparer<TKey>`；公共 approval 删除旧类型和 NPOI 实现。
- 相关文件：`src/Bing.Offices.Abstractions/AssemblyInfo.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelWorkbookImportRequest.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportResult.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelWorkbookImportResult.cs`、`tests/Bing.Offices.Tests/ExcelCompatibilityTestAdapter.cs`。
- 验证：Public API approval 和成员签名扫描通过；Unit net8 `114/114`、Integration net8 `10/10` 通过。

## 本轮结论

状态保持 `PARTIAL`。FIX-001、FIX-002、FIX-003、FIX-004、FIX-006 已完成并有回归证据；FIX-005 的共享计划已建立但 Exporter 全管线统一尚未完成。Docs Consumer、成员完整快照、Office/LibreOffice 互操作和格式基线仍未闭合。本轮不执行 `task-finish`，等待继续修复或下一次独立 Review。

## Review 修复记录

### Round 1（续）

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/bing-offices-excel-import-export-enhancement-v2/review.md`
- 执行状态：`COMPLETED`
- 本轮未修改 `review.md`，未执行 commit、push、PR 或发布。

#### FIX-001

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 修复：`ExcelImportErrorCollector` 保持 Workbook 根级上限，子 Sheet 只保留局部视图；MaxErrors 达到上限后只设置 `ErrorsTruncated`，不追加第 N+1 条错误；MaxRows 使用共享运行时并阻止后续 Sheet 初始化；关系索引和导航绑定检查同一上限及取消令牌；图片数量、单图字节数、总字节数由同一跨 Sheet tracker 独立执行。
- 测试：`Import_WorkbookResourceLimits_ShouldStopRowsAndErrors`、`Import_MaxErrors_ShouldTruncateWithoutExceedingLimit`、`Import_WorkbookImageResourceLimits_ShouldBeGlobalAndExact`，以及 Unit/Integration 全量回归。

#### FIX-002

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 修复：保持 NPOI Importer/Exporter、Resolver 和低层扩展 internal；公开 `ExcelNpoiServiceCollectionExtensions.AddNpoi()`；Benchmark 改用 `IExcelImporter`/`IExcelExporter` 和公开注册入口；`ColorResolver` 保持 internal。
- 测试：Docs Consumer、Benchmark 构建、`PublicApi_NpoiAssembly_ShouldExposeOnlyRegistrationEntry`、`PublicApi_NpoiAssembly_ShouldMatchExactMemberBaseline` 和完整公开签名扫描。

#### FIX-003

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 修复：`ErrorRowsOnly` 使用独立 Workbook，仅复制带错误的 Sheet、请求指定 HeaderRowIndex 和去重后的失败行；复制 CellType、公式、样式/格式；只复制连续可重建的 Merge、Validation、Picture 区域并重算行坐标；追加来源和聚合错误列；不再走 AnnotatedOriginal 批注路径；非零表头行也会连续重排。
- 测试：`Import_ErrorRowsOnly_ShouldCopyAndReorderOriginalFailureRows` 覆盖非零表头、公式、DataFormat、Merge、Validation、Picture 和汇总 Sheet。

#### FIX-004

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 修复：Unsupported Report 先写结构化 WorkbookValidation 错误再继续导入，Fail 拒绝当前行；显式常量列表与直接 Cell Range、Named Range/自定义公式分流；日期/时间支持 DateTime、Excel serial 和数值公式；`ValidationRangeIndex` 只保存矩形范围，不展开单元格；固定/动态图片列都通过共享计划执行 First/All/Fail，集合目标可识别。
- 测试：`Import_UnsupportedWorkbookValidation_ShouldReportOrFailByPolicy`、`Import_WorkbookExplicitListValidation_ShouldReportInvalidCell`、`Import_DirectCellRangeListValidation_ShouldResolveReferencedValues`、`ValidationRangeIndex_LargeRectangle_ShouldQueryWithoutCellExpansion`、`Import_FixedImageColumnMultiplicity_ShouldApplyConfiguredPolicy` 及图片资源边界回归。

#### FIX-005

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 修复：`ExcelColumnPlan` 改为唯一 sealed immutable plan；Importer 和 Exporter 直接使用同一类型，不再保留 `ImportColumn`/`ExportColumn` 包装；计划预绑定 Getter/Setter、Key/Title、ValueType、ConverterName、ValidatorName、Formatter、DecimalScale、ValueMap、Ignore、Unique、Merge、ImageMultiplicity 和 Style 字段；Exporter fixed/dynamic 统一进入 `WriteCell`。
- 测试：`ColumnPlan_FixedAndDynamic_ShouldShareCompiledExecutionMetadata`、关系/comparer 回归、固定/动态图片策略回归、导出/导入全量回归。

#### FIX-006

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 修复：测试 `ExcelCompatibilityTestAdapter` 已删除，Integration 项目移除其 Compile Link，Excel 测试全部使用 Workbook API；`HasMany<TKey>` 关系绑定保留 comparer 和真实来源 Sheet/Row；新增非 string Key 与大小写不敏感 comparer 测试；NPOI 唯一注册入口和 Abstractions/Core/NPOI 完整公开成员快照均纳入 approval；Docs Consumer 继续验证外部编译消费。
- 测试：旧 Adapter/旧 Excel Options/Result 搜索、`Import_RelationWithNonStringKey_ShouldKeepSourceLocation`、`Import_RelationWithCustomComparer_ShouldBindCaseInsensitiveKeys`、`PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。

### Round 1 汇总

- MUST_FIX：6 项。
- 已完成：FIX-001、FIX-002、FIX-003、FIX-004、FIX-005、FIX-006。
- PARTIAL：无 MUST_FIX；真实 Office/LibreOffice 互操作、格式基线和 Benchmark 扩展仍是计划中的 SHOULD_FIX/环境项。
- BLOCKED：net5.0 测试程序集已编译，但本机未安装 `Microsoft.NETCore.App 5.0`，运行阶段 `TESTRUNABORT`；net6/net8 Unit 与 Integration 已通过。
- 回归验证：Unit net8、Unit net6、Integration net8、Integration net6、Docs Consumer、Benchmark build、Release solution build、NPOI pack、Public API exact、`git diff --check` 均通过；构建保留 netcoreapp3.1 包支持、legacy converter 和 XML 注释警告。
- 格式：`dotnet format --verify-no-changes` 仍受仓库既有导入排序、命名和换行基线影响，本轮未进行无关大范围格式化。
- 下一步：运行 `node .agents/scripts/task-finish.mjs bing-offices-excel-import-export-enhancement-v2` 后交回独立 Reviewer 再次验收。

## Review 修复记录

### Round 2

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/bing-offices-excel-import-export-enhancement-v2/review.md`
- 执行状态：`COMPLETED`
- 本轮仅处理 `FIX-001`、`FIX-004`、`FIX-005` 的 `MUST_FIX`；未修改 `review.md`、`plan.md`，未执行 commit、push、PR 或发布。

#### FIX-001

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 根因：错误收集器在保存第 N 条错误后只表现为达到上限，结构错误路径可能跳过后续 Sheet，却没有统一记录截断状态。
- 修复：`ExcelImportErrorCollector.Add()` 在成功保存第 N 条错误时立即设置 `IsTruncated`；后续 Sheet、行、关系路径继续使用同一根级上限和截断元数据。
- 相关文件：`src/Bing.Offices.Npoi/Imports/ExcelImportErrorCollector.cs`、`tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`。
- 测试：`Import_StructureErrorAtMaxErrors_ShouldMarkTruncatedAndStopLaterSheets`、`Import_MaxErrors_ShouldTruncateWithoutExceedingLimit` 通过。

#### FIX-004

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 根因：Validation 已避免按 Cell 预展开，但查询仍逐次遍历全部矩形规则。
- 修复：`ValidationRangeIndex` 改为按行区间构建平衡区间树；查询按目标行进入候选节点，再检查列区间并按规则引用去重，不再扫描所有无关范围。
- 相关文件：`src/Bing.Offices.Npoi/Imports/ValidationRangeIndex.cs`、`tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`。
- 测试：`ValidationRangeIndex_LargeRectangle_ShouldQueryWithoutCellExpansion`、`ValidationRangeIndex_DisjointRanges_ShouldReturnOnlyMatchingRules` 通过。

#### FIX-005

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 根因：共享 `ExcelColumnPlan` 只保存转换器/校验器名称，Importer/Exporter 仍在单元格热路径解析服务集合，并保留固定/动态分支校验。
- 修复：`ExcelColumnPlan` 增加预绑定 `ValueConverters`、属性校验绑定和命名校验规则；Importer 在列计划创建阶段完成唯一解析，固定/动态列统一消费 `ValidateColumnValue()`；Exporter 固定/动态列统一消费计划中的已绑定转换器；删除逐单元格转换器解析及 `ValidateDynamicValue()`/`ValidateConvertedValue()`。
- 相关文件：`src/Bing.Offices.Npoi/ExcelColumnPlan.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`、`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`。
- 测试：`Import_FixedColumnConverter_ShouldBindOnceBeforeCellConversion`、`Export_FixedColumnConverter_ShouldBindOnceBeforeCellConversion`、动态转换器和既有重复校验回归通过。

### Round 2 汇总

- MUST_FIX：3 项。
- 已完成：FIX-001、FIX-004、FIX-005。
- PARTIAL：无本轮 MUST_FIX；仓库格式基线和真实 Office/LibreOffice 互操作仍保持原有未闭合状态。
- BLOCKED：无。
- 回归验证：Unit net8 `133/133`、Unit net6 `133/133`、Integration net8 `10/10`、Integration net6 `10/10`、Docs Consumer `2/2`；Release build、pack 和 `git diff --check` 通过。
- 格式：`dotnet format --verify-no-changes --no-restore` 仍失败，报告仓库既有 `FINALNEWLINE`、`IMPORTS`、`IDE1006` 问题；本轮未执行无关大范围格式化。
- 下一步：执行 `node .agents/scripts/task-finish.mjs bing-offices-excel-import-export-enhancement-v2`，交回独立 Reviewer 再次验收。

### Round 3

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/bing-offices-excel-import-export-enhancement-v2/review.md`
- 执行状态：`COMPLETED`
- 本轮仅处理 `FIX-004`、`FIX-005` 两项 `MUST_FIX`；未修改 `review.md` 或 `plan.md`，未执行提交、推送、PR 或发布。

#### FIX-004

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 根因：上一轮 ValidationRangeIndex 仅按行建立区间树，同一行的大量不相交列范围仍会在 Cell 查询中扫描同一 overlap 列表。
- 修复：保留不展开矩形范围的行区间树，并为每个行节点的 overlap 集合建立列区间树；查询同时按目标行和目标列递归缩小候选，最终按 `IDataValidation` 引用去重。增加 internal 候选检查计数入口，仅用于白盒验证候选集规模，生产导入继续使用原 `Get(row, column)` API。
- 相关文件：`src/Bing.Offices.Npoi/Imports/ValidationRangeIndex.cs`、`tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`。
- 测试：`ValidationRangeIndex_LargeRectangle_ShouldQueryWithoutCellExpansion`、`ValidationRangeIndex_DisjointRanges_ShouldReturnOnlyMatchingRules`、`ValidationRangeIndex_OverlappingRows_ShouldLimitColumnCandidates` 通过；新增 200 个同一行、互不相交列范围并断言候选检查数小于全部规则数。

#### FIX-005

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 根因：上一轮虽把 Converter/Validator 解析移到列创建阶段，但 Importer/Exporter 仍分别实现固定/动态默认转换、ValueMap、属性校验和写入回退。
- 修复：将 `ConvertFrom`、`ConvertTo`、默认类型转换、ValueMap、命名/属性规则绑定和 `WriteValue` 收敛到共享 `ExcelColumnPlan`；Importer fixed/typed dynamic 均消费计划转换和验证路径，动态字典只保留取值/写值适配；Exporter fixed/dynamic 均消费计划转换和写入路径；删除旧 `ConvertValue`、`ConvertDynamicValue` 和动态逐 Cell Converter 遍历，唯一性继续通过预绑定 Duplication 规则执行。
- 相关文件：`src/Bing.Offices.Npoi/ExcelColumnPlan.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`、`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`。
- 测试：`Import_FixedColumnConverter_ShouldBindOnceBeforeCellConversion`、`Import_DynamicColumnConverter_ShouldBindOnceBeforeCellConversion`、`Export_FixedColumnConverter_ShouldBindOnceBeforeCellConversion`、动态 DataType/Converter、固定 ValueMap 和既有唯一性回归通过。

### Round 3 汇总

- MUST_FIX：2 项。
- 已完成：FIX-004、FIX-005。
- PARTIAL：无本轮 MUST_FIX。
- BLOCKED：无。
- 回归验证：Unit net8 `135/135`、Unit net6 `135/135`、Integration net8 `10/10`、Integration net6 `10/10`、Docs Consumer net8 `2/2`；专项测试 `6/6` 及动态回归 `3/3` 通过。
- Build/Pack：`dotnet build .\Bing.Offices.sln -c Release --no-restore` 通过；`dotnet pack .\Bing.Offices.sln -c Release --no-build --no-restore` 通过。
- 静态检查：改动源文件和测试文件编辑器诊断无错误；`git diff --check` 通过，仅有 CRLF/LF 转换提示。
- 格式：`dotnet format .\Bing.Offices.sln --verify-no-changes --no-restore` 仍受仓库既有格式基线问题影响，本轮未自动改写文件。
- 下一步：执行 `node .agents/scripts/task-finish.mjs bing-offices-excel-import-export-enhancement-v2`，交回独立 Reviewer 再次验收。
