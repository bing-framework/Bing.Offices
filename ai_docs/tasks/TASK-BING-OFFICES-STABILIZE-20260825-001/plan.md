# Bing.Offices 稳定化、API 收敛与发布准备实施计划

- Task ID: `TASK-BING-OFFICES-STABILIZE-20260825-001`
- 计划状态: `IN_PROGRESS`
- 创建日期: 2026-08-25
- 本轮角色: `plan-writer`，仅完成现状审计与计划写入，未修改生产代码、测试、配置、版本或包。
- 目标: 以 Workbook Request 为唯一主路径，修复 Mapping/导入正确性缺陷，收敛公开 API，拆分过重实现，补齐可重复测试和性能证据，并达到可审查的发布准备状态。

## 1. 输入、约束与冲突处理

### 1.1 有效输入

- 用户任务描述要求覆盖 `Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi`、Unit/Integration/Docs Tests、Benchmarks 与 docs。
- 根目录 `AGENTS.md` 要求 UTF-8 文件处理、保留用户现有改动、避免破坏性 Git 操作；其 SQL 专项规则不适用于本任务。
- `.github/copilot-instructions.md` 的作用域是 `framework/`，与本仓库的 `src/`、`tests/` 不重叠；其中 Unit/Integration 分层、xUnit/AAA、中文测试说明与最小改动原则仍作为参考。
- `.github/skills/chinese-comments/SKILL.md` 适用于后续 C# 修改：改动范围内类型、成员、字段、枚举须有准确中文 XML 注释；实现/override 优先 `inheritdoc`。
- 用户引用的 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260825.md` 不存在，不能作为当前事实来源。计划以当前源码、测试和 `ai_docs/excel/implementation-progress.md` 为准。

### 1.2 模式冲突

用户原始任务要求“生成计划后立即持续实施”，但当前 Agent 为 `plan-writer` 模式，唯一允许写入目标是本文件，明确禁止进入实施阶段。因此本计划在此停止；执行阶段须通过 `/execute-plan`、`/run-plan` 或对应 `plan-executor` 角色继续。

### 1.3 工作区保护

- 本轮未能取得 Git 状态，因为可用工具未包含终端执行能力；实施 Agent 的第一步必须运行 `git status --short` 并将既有改动与本任务改动隔离。
- 不执行 `git reset --hard`、强制 checkout、清理仓库、提交、推送、tag、PR、NuGet publish 或版本号修改。
- 后续任务状态只可使用 `TODO`、`IN_PROGRESS`、`DONE`、`BLOCKED`；任一时刻最多一个顶层任务为 `IN_PROGRESS`。

## 2. 仓库事实与当前完成度

### 2.1 技术结构

| 范围 | 当前事实 |
| --- | --- |
| 解决方案 | `Bing.Offices.sln` 含 Abstractions、Core、Npoi、Unit、Integration、Docs Tests、ProfileFixtures、Benchmarks。 |
| 产品目标框架 | Abstractions/Core 为 `netstandard2.0`；Npoi 经 `framework.props` 为 `net8.0;net7.0;net6.0;netcoreapp3.1`。 |
| 测试 | Unit 为 `tests/Bing.Offices.Tests`（`net8.0;net7.0;net6.0;net5.0;netcoreapp3.1`）；Integration 为 `net6.0;net8.0`；另有本地包消费者文档测试。 |
| 依赖 | Core 使用 `Bing.Utils`、`CsvHelper 33.1.0` 和 DI abstractions；Npoi 使用由属性管理的 NPOI 版本。 |
| 基准 | `benchmarks/Bing.Offices.Benchmarks` 为 `net8.0`，使用 BenchmarkDotNet 0.14.0 和 `MemoryDiagnoser`。 |
| 文档 | `docs/excel` 已有请求、Profile、JSON/XML、导入校验、动态列和迁移文档；`ai_docs/excel` 有内部设计与历史执行进度。 |

### 2.2 已实现且有源码/测试证据的能力

- Workbook Request 已是导入/导出公开主契约；NPOI 实现类为 internal，仅公开 DI 扩展。`PublicApiContractTest` 还检查 NPOI 类型不泄漏和生产 IVT 仅面向测试。
- 映射支持 Attribute、Profile、JSON/XML Document 与请求配置；Profile 有 import-only、export-only、同模型和异模型四种契约，并有注册、重复冲突与并发读取测试。
- 映射 Patch 合并单测覆盖 `ClearDynamicColumns`、列值映射/别名清除、`ResetStyle`、`ResetLayout` 和部分 Attribute/Profile 覆盖。
- NPOI 导入已有失败工作簿、非零表头、图片、Workbook 原生校验、资源上限、非 seekable 输入、取消、关系与结构化错误测试。
- 模板区域、样式、图表、动态列、CSV、JSON/XML 安全限制已有实现和测试入口。
- 默认直接构造 `NpoiExcelImporter` 与 `ExcelMappingPlanFactory` 都调用 `ExcelValidationRules.CreateDefault()`，其中包含 `MaxValueExcelValidationRule`。

### 2.3 已确认缺口与风险

| ID | 优先级 | 当前证据 | 结论 |
| --- | --- | --- | --- |
| P0-VAL-01 | P0 | `NpoiExcelImporter.ImportSheet` 仅对 `workbookValidations` 按 `ValidationMode` 分支；`ValidateRawValues` 与 `TryCreateItem` 无条件执行。 | `Disabled`、`WorkbookRules` 仍会运行 Attribute/配置校验和 unique journal，违反四模式语义。 |
| P0-VAL-02 | P0 | `ExcelNpoiServiceCollectionExtensions.AddNpoi()` 手工注册 Required/Regex/Range/MaxLength/Date/Unique，遗漏 `MaxValueExcelValidationRule`；直接构造则使用完整 `CreateDefault()`。 | DI 与直接构造规则集合漂移；现有 DI 集成测试未覆盖 `[ExcelMaxValue]` 和完整默认规则等价性。 |
| P0-MAP-01 | P0 | `ExcelSheetImportBuilder.Build()`、`ExcelSheetExportBuilder.Build()`、`ExcelMappingDocumentFactory.Create()`、`ExcelMappingPlanFactory.ResolveDocument()`、NPOI provider 路径均出现 `Merge`/规范化调用。 | 合并边界不唯一，Patch tombstone 可能在中间合并后丢失；必须重建全调用链并以端到端测试定责。 |
| P0-MAP-02 | P0 | `ExcelMappingDocumentFactory.CloneOrEmpty()` 将缺失方向转换为 convention 空配置；`ExcelMappingDocument.Import/Export` 默认实例化为空配置。`FromJson/FromXml` 兼容 facade 固定返回 `.Import`。 | 缺失方向可静默回退；方向不安全 facade 仍是公开 API，需作为 breaking change 收敛。 |
| P0-FAIL-01 | P0 | 已有 `ErrorRowsOnly` 单 Sheet + 非零 HeaderRowIndex 回归测试，但需审查 writer 是否以真实解析后的 Sheet identity 关联 request，并补多 Sheet/ByIndex 反例。 | 基础行为已有证据，跨 Sheet identity 和失败产物规模/ownership 尚未满足目标。 |
| P0-TPL-01 | P0 | 导出正文单元格使用 `GetCell() ?? CreateCell()`；但表头循环仍使用 `header.CreateCell(...)`。现有测试证明部分模板格式保留。 | 需定义 overwrite policy，检查表头、固定列、custom headers 和公式/批注的完整保真，而非仅验证正文路径。 |
| P1-API-01 | P1 | `ExcelMapping.For<T>()`、`FromJson/FromXml`、`MappingConfigurationMerger`、多个编译实现类型及 NPOI metadata 仍在 `PublicApiContractTest` 基线。 | 公开面仍过宽，必须先建立成员级清单和迁移策略后集中破坏性收敛。 |
| P1-API-02 | P1 | `CsvStreamExtensions.ExportToBytesAsync` 使用 `Task.FromResult`，虽已 Obsolete。 | 伪异步 API 未删除，不符合“无伪异步”完成定义。 |
| P1-PROFILE-01 | P1 | Profile 批量注册扩展仍在 Npoi；接口 `IMappingProfileRegistry` 同时暴露 `Register` 和读取；Profile 名默认 `FullName`。扫描直接调用 `assembly.GetTypes()`。 | Core DI 归属、只读 resolver、稳定 alias、`ReflectionTypeLoadException` 容错未完成。 |
| P1-PERF-01 | P1 | 导入无条件复制 source 到 `MemoryStream`；失败输出基于 byte[]；Plan cache key 每次 JSON 序列化；基准含主动 90KB LOH payload 探针。 | 有已知大对象复制与不代表产品路径的基准负载，尚无可比优化基线。 |
| P1-MAINT-01 | P1 | `NpoiExcelImporter`、`NpoiExcelExporter`、`ExcelMappingConfigurationLoader`、`CsvEntityPipeline` 职责密集；`Extensions.Service.cs` 命名不清。 | 需要先以 internal collaborator 拆分，避免公开 API 重构与行为修复耦合。 |

### 2.4 测试和发布基线可信度

`ai_docs/excel/implementation-progress.md` 记录过一次 build/test/pack 成功，但也记录 locked restore 的 `NU1004`，且该记录日期早于当前任务。当前计划不得把其结果视为本轮验证；实施开始时必须重新运行并将原始输出写入 `progress.md` 和 `verification.md`。

当前 Unit/Integration 有较多回归测试，但不足以证明：四种 ValidationMode 的交叉矩阵、所有默认规则 DI/直接构造一致性、Profile/Document 缺失方向的失败语义、所有 Patch 字段经 Builder 到 Provider 的保留、真实模板/Office 互操作、资源限制分层和完整成员 API baseline。

### 2.5 完成度判断

功能主链约为 70% 至 75%：Workbook Request、多 Sheet、动态列、模板、图表、CSV 与配置加载均已落地。正确性和 API 发布成熟度较低：P0 校验模式、DI 规则漂移、方向缺失 fallback 和多次合并仍直接影响行为。测试结构存在但矩阵不完整；性能基准存在但未构成产品路径的决策证据。原始输入给出的 68% 总体估计可作为历史参照，不能作为当前验收结论。

## 3. 目标不变量和实施顺序

### 3.1 必须保持的不变量

1. 配置优先级固定为 `Attribute < Profile < JSON/XML Document < Request Fluent`；仅 Mapping Plan compiler 可完成最终合并和归一化。
2. `Clear`、`Reset`、`Remove` 为 patch/tombstone，直到最终 compile 前不得降格为普通空值。
3. import-only/export-only Profile 或 Document 请求相反方向必须失败；只有显式 `UseConventionFallback` 可回退，默认关闭。
4. JSON 与 XML 必须使用同一 schema/validation/优先级规则；禁止 DTD 和外部实体，限制大小、深度和集合规模。
5. 只保留测试、集成测试、benchmark 的 IVT；不可用生产 IVT 规避程序集边界。
6. source、destination、template 的 ownership、keep-open、取消语义和可重复执行性必须在 API docs 说明并测试。
7. NPOI DOM 管道不得宣称 streaming 或 0 GC；性能变更必须有 benchmark 或 allocation/GC 证据。

### 3.2 阶段依赖

`0 基线 -> 1 P0 正确性 -> 2 Mapping 语义 -> 3 API 治理 -> 4 内部重构 -> 5 测试/集成扩展 -> 6 性能 -> 7 文档/发布验证`。

Phase 1/2 的端到端测试必须先于 API 删除和文件拆分，以锁定外部行为。Phase 3 的 breaking change 合并必须先完成 API 清单和兼容文档。性能优化只在正确性回归已绿后进行。

## 4. 分阶段实施任务

## Phase 0: 基线、状态和可追溯性

### TASK-0.1 [P0] 初始化状态文件并重新采集真实基线

- 状态: `TODO`
- 目标: 按任务要求创建/更新 `progress.md`、`decisions.md`、`verification.md`、`api-governance.md`、`performance-baseline.md`；仅在完成或真实外部阻断时创建 `final-review.md`。
- 已确认文件: `Bing.Offices.sln`、`common.props`、`framework.props`、`common.tests.props`、各产品/测试/基准 csproj、`ai_docs/excel/implementation-progress.md`。
- 步骤:
  1. 执行 `git status --short`，将既有用户改动逐项记录，不修改无关文件。
  2. 列出 solution 项目、TFM、PackageReference、lock file 状态和 docs test 的包消费前置条件。
  3. 搜索并分类 `TODO`、`FIXME`、`HACK`、`NotImplementedException`、`.Result`、`.Wait()`、`Task.FromResult`、`InternalsVisibleTo`、`Obsolete`，并将结果按产品/测试/基准分类。
  4. 在未改代码前按顺序运行下列真实命令并保存原始结果：
     - `dotnet restore Bing.Offices.sln --locked-mode`
     - 若 locked mode 因 lock 不一致失败，记录失败后执行 `dotnet restore Bing.Offices.sln --force-evaluate`，不得静默接受 lock 变更。
     - `dotnet build Bing.Offices.sln -c Release --no-restore`
     - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore`
     - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore`
     - `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore`
     - `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore`
     - `dotnet pack Bing.Offices.sln -c Release --no-build --no-restore`
     - `dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -- --filter *MappingValidationBenchmarks* --job Dry`
  5. 建立“生产符号 -> 测试方法”的 P0 可追溯表。
- 验收: 所有命令均有实际结果、耗时和失败原文；未知状态标记 `BLOCKED`；未把历史文档结果当本轮证据。
- 风险: 老目标框架 SDK/NuGet 兼容、锁文件不一致、Office 软件缺失。外部阻断必须记录解除步骤。

## Phase 1: P0 导入校验和失败工作簿正确性

### TASK-1.1 [P0] 统一默认校验规则注册

- 状态: `TODO`
- 依赖: TASK-0.1。
- 已确认文件: `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`、`src/Bing.Offices.Npoi/Extensions/Extensions.Service.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、Integration tests。
- 修改范围:
  - 将 `AddNpoi()` 的内置规则注册委托给唯一的默认规则定义，或在 Core 提供能创建 ServiceDescriptor 的唯一注册入口。
  - 不改变调用方自定义规则的可替换性、生命周期和枚举顺序。
  - 重命名 `Extensions.Service.cs` 为职责明确的文件仅在引用和 API 基线同步完成后进行。
- 用例矩阵:
  - Given DI `AddNpoi`，When 解析 `IExcelValidationRule`，Then 精确包含 Required/Regex/Range/MaxValue/MaxLength/Date/Unique 且无重复。
  - Given 直接构造和 DI importer，When 导入含每一种 Attribute 的相同 workbook，Then 错误码、行列和成功项一致。
  - Given `[ExcelMaxValue(10)]` 与 `11`，When DI 导入真实 XLSX，Then 返回 `MaxValue` 错误而不是成功。
  - Given 调用方预注册替代规则，When `AddNpoi`，Then 既有注册语义按既有 `TryAdd` 约定保持。
- 测试: Unit 覆盖 descriptor 集合；Integration 通过 `ServiceProvider` + NPOI XLSX 覆盖所有默认规则。测试使用中文目的说明和 AAA。
- 验收: 不存在两套手工默认规则清单；MaxValue 与其他内置规则 DI/直接构造完全等价。

### TASK-1.2 [P0] 修复 `ExcelImportValidationMode` 全链路语义

- 状态: `TODO`
- 依赖: TASK-1.1。
- 已确认文件: `ExcelImportPolicies.cs`、`ExcelImport.cs`、`ExcelImportExecutionOptions.cs`、`NpoiExcelImporter.cs`、`ExcelP0RegressionTest.cs`。
- 修改范围:
  1. 建立内部明确谓词：Configured Rules 仅在 `ConfiguredRules`、`ConfiguredAndWorkbook` 启用；Workbook Rules 仅在 `WorkbookRules`、`ConfiguredAndWorkbook` 启用。
  2. 对 raw validation、转换后 validation、named validation、UniqueTracker begin/reserve/commit/rollback 分别应用 configured 谓词；不可只跳过某一方法。
  3. Workbook validation 保持当前顺序先于 configured validation，或经测试明确为相反顺序；同单元格双失败时固定错误顺序和 `ValidateMode` 短路边界。
  4. 修正 `Disabled` 枚举中文注释为“禁用配置和 Workbook 规则”，同步 docs。
- 用例矩阵:
  - Given 一个单元格同时违反 `[ExcelMaxValue]` 和原生 list validation，When 分别运行四模式，Then `Disabled=0`、`ConfiguredRules=配置错误`、`WorkbookRules=Workbook错误`、`ConfiguredAndWorkbook=两个来源且顺序固定`。
  - Given 重复值和转换失败，When `Disabled` 或 `WorkbookRules`，Then 不创建/提交 unique journal 且不产生 configured validation 错误。
  - Given Workbook 原生规则和 `StopOnFirstFailure`/`Continue`，Then 每种模式的行保留、错误数与短路边界固定。
  - Given Unsupported Workbook rule，When `Report` 与 `Fail`，Then 仅 Workbook-enabled 模式触发既有策略。
- 验收: 四模式从 raw、实体、unique 到 native validation 均一致；无通过删除断言或吞异常获得的绿测。

### TASK-1.3 [P0] 失败工作簿按解析 Sheet identity 输出

- 状态: `TODO`
- 依赖: TASK-1.2。
- 已确认文件: `NpoiExcelImporter.cs`、`ExcelP0RegressionTest.cs`。
- 步骤:
  1. 跟踪错误收集、`ResolveSheetIndex`、`WriteFailureWorkbook` 和 Sheet request 的关联键；将按索引解析的实际 sheet name/index 持久化为执行上下文，而非二次按用户 selector 猜测。
  2. 修复 `ErrorRowsOnly` 的 ByIndex + 非零 HeaderRowIndex 场景，保持复制后的表头行和源行元数据一致。
  3. 覆盖连续重排行、合并区域、style/comment/formula/Data Validation/picture、多 Sheet 同名误匹配防护。
  4. 评审 `ExcelImportFailureOptions`：保留调用方 Destination ownership；若新增 artifact/temporary-file/callback，必须有容量阈值、清理责任和 breaking/API 评审，禁止默认返回无界 `byte[]`。
- 验收: 真实 importer 的多 Sheet ByIndex 测试通过，错误不会串 Sheet；失败工作簿重新打开后断言结构与内容。

### TASK-1.4 [P0] 定义并锁定模板 overwrite/style 语义

- 状态: `TODO`
- 依赖: TASK-0.1。
- 已确认文件: `NpoiExcelExporter.cs`、`ExcelWorkbookRequestTest.cs`、NPOI style/color extensions。
- 修改范围:
  - 在公开 options/request 上定义每种写入区域的策略：值覆盖是否保留 cell style/comment/formula，显式样式是否覆盖模板样式，完全替换是否可用。
  - 将无条件 `CreateCell()` 路径改为策略一致的读取/创建实现，特别是 header、custom header、固定列和模板区域。
  - 检查 XSSF fill 前景/背景、任意 ARGB border/font/fill color、`General`/`Bottom`/`None` 显式 reset 的可表达性。
  - 让固定列的宽度/样式能力与动态列保持同级，但不重复 provider 合并。
- 用例矩阵: XSSF 与 HSSF 能力差异须显式测试或 `NotSupportedException`；模板 cell 的 value/style/comment/formula 在每个 overwrite policy 下逐字段断言；标题与多行合并 headers 单独断言。
- 验收: 文档仅承诺已实现的“区域原点写入”能力，不声称行块复制、公式引用重写或命名区域扩张；每条支持能力有回归测试。

## Phase 2: Mapping 单一合并边界与方向语义

### TASK-2.1 [P0] 建立 Mapping layer collector 和单次 Plan Compile

- 状态: `TODO`
- 依赖: TASK-1.1。
- 已确认文件: `ExcelSheetImportBuilder.cs`（位于 `ExcelImport.cs`）、`ExcelSheetExportBuilder.cs`、`ExcelMappingDocumentFactory.cs`、`ExcelMappingPlanFactory.cs`、`NpoiExcelImporter.cs`、`NpoiExcelExporter.cs`、`CsvEntityPipeline.cs`。
- 步骤:
  1. 画出 Builder -> request -> document factory -> plan factory -> NPOI/CSV 的真实调用图，并标记每次 clone/merge/serialization/cache key。
  2. Builder 只存储 Attribute/Profile/Document/Request 各层快照；不得 merge document/request。
  3. Document factory 仅做方向选择、不可变快照和结构校验；不得 merge 或用空 configuration 伪造缺失方向。
  4. `ExcelMappingPlanFactory` 成为唯一合并点，严格按 Attribute -> Profile -> Document -> Request 生成一次 plan；cache key 以未合并层的稳定规范化表达或 revision 组成。
  5. NPOI/CSV 只消费已编译 plan，删除 provider 内再归一化路径。
- 用例矩阵:
  - Given Attribute/Profile/JSON/Request 配置，When import/export via NPOI and CSV，Then 最终 title/converter/value mapping/style/layout 为最高层值。
  - Given `ResetStyle`、`ResetLayout`、`ClearDynamicColumns`、`DynamicColumnKeysToRemove`、`ClearValidationRules`、`ValidationRuleNamesToRemove`，When 通过真实 exporter/importer，Then 移除操作不会被低层恢复。
  - Given 高层清除 Attribute validation，When 执行 importer，Then 被移除规则不产生错误；反向未移除场景仍产生错误。
  - Given 缓存 hit/miss 和同租户重复渲染，Then plan identity/结果稳定且不会二次 merge。
- 风险: 合并职责迁移会影响 JSON/XML、Profile、CSV 和 NPOI；必须使用 P0 回归群与完整 API diff。
- 验收: `MappingConfigurationMerger.Merge` 的生产调用仅存在于 Plan compiler（及有明确局部规范化理由的 loader migration）；所有 Patch 字段都有真实端到端覆盖。

### TASK-2.2 [P0] 收紧缺失方向和显式 convention fallback

- 状态: `TODO`
- 依赖: TASK-2.1。
- 已确认文件: `ExcelMappingDocument.cs`、`ExcelMappingDocumentFactory.cs`、`ExcelMappingConfigurationLoader.cs`、`ExcelMappingPlanFactory.cs`、Profile contracts/registry/tests。
- API 方案:
  - `ExcelMappingDocument.Import`/`Export` 改为可表示“未提供”的方向，而不是默认空配置；或引入明确方向存在位，避免 null/empty 混淆。
  - 在请求或 mapping options 上提供显式 `UseConventionFallback`，默认 `false`。
  - `IImportMappingProfile<T>` 请求 export、`IExportMappingProfile<T>` 请求 import、Document 缺少目标方向均抛出包含 profile/document、direction、model 的明确 `InvalidOperationException`。
  - `FromJson`/`FromXml`/file overload 固定 `.Import` 的 API 设为 breaking removal 或改名为 `FromJsonImportConfiguration`、`FromXmlImportConfiguration`；主 API 只返回 `ExcelMappingDocument`。
- 测试: JSON、XML、Profile、Request 四来源分别测试缺失方向失败与显式 fallback 成功；测试双向不同模型不可误用；docs consumer 使用新推荐 API。
- 验收: 不存在默认 silent convention fallback；方向错误可诊断；JSON/XML 保留且共享 validation。

### TASK-2.3 [P1] Profile 注册治理与扫描容错

- 状态: `TODO`
- 依赖: TASK-2.2。
- 已确认文件: `IMappingProfileRegistry.cs`、`MappingProfileRegistry.cs`、`MappingProfileServiceCollectionExtensions.cs`、`ProfileDescriptorFactory.cs`、`MappingProfileRegistryTest.cs`、`docs/excel/mapping-profile.md`。
- 修改范围:
  1. 将 Profile DI 扩展和 factory 从 Npoi 移至 Core，以便 Core-only consumer 无 provider 依赖即可注册解析 Profile。
  2. 拆分只读 `IMappingProfileResolver` 与仅 bootstrap/DI 使用的 mutable registry；普通业务 consumer 只能读取。
  3. 为 Profile 引入显式稳定 alias（接口成员、attribute 或 registration overload 三选一，需以契约清晰和多 TFM 兼容为准）；类型 FullName 只作为兼容期 fallback，不能作为文档推荐。
  4. 扫描 `Assembly.GetTypes()` 时捕获 `ReflectionTypeLoadException`，保留可加载类型，并将 `LoaderExceptions` 聚合为可观测诊断/异常信息。
  5. 保持 `(alias, direction, model)` 冲突确定性和并发读取不可变快照。
- 验收: Core-only 注册可用，Npoi 不再拥有 Profile 注册 API；alias 冲突、显式接口实现、扫描部分加载失败、并发解析均有测试。

## Phase 3: API 收敛和 Breaking Change 治理

### TASK-3.1 [P0] 建立成员级公开 API 清单与迁移决策

- 状态: `TODO`
- 依赖: TASK-2.2。
- 已确认文件: `PublicApiContractTest.cs`、公开类型搜索结果、README/docs。
- 步骤:
  1. 导出 Abstractions/Core/Npoi 的 public types 和 public/protected member signatures，替换当前只覆盖顶层类型且部分精确成员的基线。
  2. 对每个候选类型记录消费方引用：仓库生产、测试、docs consumer、NuGet public compatibility。
  3. 在 `api-governance.md` 标记保留、合并、删除、重命名、internal/private、`EditorBrowsable(Never)` provider SPI，并说明原因、影响和迁移。
  4. 对 breaking API 指定下一个 major 或 pre-release；本任务不修改版本号、不发包。
- 首批候选:
  - `ExcelMapping.For<T>()` 与旧 Fluent builder；
  - fixed-import `FromJson`/`FromXml` facade；
  - `CsvStreamExtensions.ExportToBytesAsync`；
  - obsolete CSV helper overload 和旧 validation attributes；
  - `SheetSetting`、`ExcelValueMap<T>`、专用 exceptions 的实际引用；
  - `ExcelTypeMapFactory`、`ExcelTypeMap<T>`、`ExcelPropertyMap`、`ExcelValidationBindingFactory`、`ExcelValueConverterBindingResolver`；
  - reflection/expression/PropertyInfo extensions；
  - `MergedRegionInfo`、`PictureInfo`、`PictureStyle`。
- 验收: 每个公开类型有决策和引用证据；不根据“当前仓库无引用”单独判断 NuGet API 可删除。

### TASK-3.2 [P0] 落地批准的公开 API 收敛

- 状态: `TODO`
- 依赖: TASK-3.1。
- 步骤:
  1. 删除伪异步 `ExportToBytesAsync`，不以 `ValueTask` 或同步包装替代；docs 改为 Stream-first / `ExportToBytes`。
  2. 删除或定向重命名方向不安全 loader facade，更新所有生产、测试、docs consumer 调用。
  3. 将仅编译细节降为 internal/private；若 Npoi 需要跨程序集访问，则用最小的 provider-neutral SPI 而非 production IVT。
  4. 将真正 provider SPI 保持 public + `EditorBrowsable(Never)`，中文 docs 标注“provider 实现扩展点，普通用户不应直接调用”。
  5. 对 obsolete API 执行已批准的 major/pre-release 删除，不无限保留兼容包装器。
  6. 更新精确 API baseline、README 及迁移文档；确保 public signature 无 NPOI 类型。
- 验收: 一条文档化推荐路径；无 `Task.FromResult` 伪异步；无生产 IVT；API baseline 与迁移表同步通过。

## Phase 4: 内部职责拆分和目录治理

### TASK-4.1 [P1] 拆分 importer/exporter 内部协作者

- 状态: `TODO`
- 依赖: TASK-1.3、TASK-1.4、TASK-3.2。
- 候选文件: `NpoiExcelImporter.cs`、`NpoiExcelExporter.cs`、`ValidationRangeIndex.cs`、NPOI extensions。
- 步骤:
  1. 先从 importer 提取 internal concrete collaborator：source reader、sheet selector、header binder、row materializer、configured validation pipeline、workbook validation pipeline、image index/binder、relation binder、failure workbook writer。
  2. 从 exporter 提取：workbook target/template reader、sheet/header writer、cell value writer、style resolver/cache、column width planner、merge planner、chart renderer。
  3. 构造器只接收必需依赖；仅在出现第二实现、provider contract 或测试替身时引入 interface。
  4. 一个主要 public 类型一个文件；internal 按职责分目录；修正 `ConditionalFormattin` 和 `Extensions.Service.cs` 等命名时同步 namespace、项目包含项和 docs。
- 测试: 为承担真实逻辑的 internal collaborator 增加直接 Unit 测试，公共 API 集成测试锁定行为。
- 验收: 导入/导出主类不再同时拥有所有算法；无循环依赖、无生产 IVT、无“策略链”过度抽象。

### TASK-4.2 [P1] 拆分 Mapping/serialization 与 CSV 管线

- 状态: `TODO`
- 依赖: TASK-2.1、TASK-2.3、TASK-3.2。
- 候选文件: `ExcelMappingConfigurationLoader.cs`、`ExcelMappingPlanFactory.cs`、`CsvEntityPipeline.cs`。
- 步骤: 分离 mapping layer collector、plan compiler、Profile discovery/registry/resolver、JSON parser、XML parser、v1 migration、shared document validator；CSV 分离 header binding、record reader/writer、converter/validation execution。
- 验收: JSON/XML 共享 schema validation；缓存、parser security 和 CSV 行级行为有直接测试；不改变已锁定公开行为。

## Phase 5: 测试体系补齐

### TASK-5.1 [P0] 建立 P0 行为矩阵和测试可追溯表

- 状态: `TODO`
- 依赖: Phase 1 和 Phase 2 对应修改。
- 测试项目: `tests/Bing.Offices.Tests` 为主；真实 DI/NPOI、临时文件和非 seekable stream 进入 `tests/Bing.Offices.Tests.Integration`；消费者 API/文档示例保持在 `Bing.Offices.Docs.Tests`。
- 必测集合:
  - DI MaxValue 与全部默认规则一致性；
  - ValidationMode 四模式交叉矩阵；
  - 缺失方向和 explicit convention fallback；
  - Attribute/Profile/JSON/XML/Request 优先级；
  - Builder -> Plan -> NPOI/CSV 的 Clear/Reset/Remove；
  - ByIndex + HeaderRowIndex + ErrorRowsOnly；
  - 模板 style/comment/formula/value；
  - XSSF background/border/reset style；
  - `ReflectionTypeLoadException` 扫描。
- 约束: 测试方法英文 `Method_State_Expected`，每个测试中文 XML 目的说明，AAA；只 mock 时间/IO/外部依赖，不 mock 被测实现细节；不使用 Skip 掩盖失败。
- 验收: `api-governance.md` 和 `verification.md` 含最终生产符号到测试方法映射。

### TASK-5.2 [P1] 扩展边界、并发和安全集成测试

- 状态: `TODO`
- 依赖: TASK-5.1。
- 覆盖: 四种 Profile、alias 冲突、并发读取；Required/Regex timeout/Date/Max/Range/MaxLength/Unique 的 nullable/empty/culture/boundary；Unique rollback/限制；Workbook native validation 多类型和隐藏列表 sheet；取消和 ownership；多 Sheet relations；HSSF/XSSF 图片；多租户 cache 隔离；损坏、空和不支持的流；JSON/XML v1-v2、未知字段、深度/大小、DTD/实体攻击。
- Office 互操作: Excel/LibreOffice/WPS 验证仅在有可控安装环境时执行；没有环境时标记 `BLOCKED`，保留 NPOI reopen 测试作为已验证范围。
- 验收: Integration 不依赖公网、sleep 或生产服务；外部条件经环境变量或显式前置条件门控。

## Phase 6: 性能与 Benchmark 治理

### TASK-6.1 [P1] 修复基准有效性并记录基线

- 状态: `TODO`
- 依赖: TASK-2.1、TASK-4.2、TASK-5.1。
- 已确认文件: `Program.cs`、`MappingValidationBenchmarks.cs`、`StreamPipelineBenchmarks.cs`。
- 步骤:
  1. 保留 `MemoryDiagnoser`；修复 `MultiRulePlanBuild` 缺少真实 named validators 的问题，确保基准测构建而非异常。
  2. 将 `DynamicPlanBuild` 分为 cold、cache hit、cache miss；`AssemblyScanRegistration` 应包含 provider build/profile registration/resolution。
  3. 删除或隔离 `MappingValidationBenchmarks` 和 `ResourceProbe` 中主动分配 90KB byte[] 的 LOH 负载，使产品路径与资源压力测试分开。
  4. 统一 `LohSizeBytes`/`LohRetainedBytes` 输出名称；峰值 working set 保持独立子进程采集。
  5. 记录 Mean/Median（BenchmarkDotNet 支持时）、Allocated、Gen0/1/2、LOH、managed peak、working set、输出大小、输入规模和开关到 `performance-baseline.md`。
- 验收: smoke 不因配置抛异常；资源 probe 不把主动假负载误称为产品 LOH。

### TASK-6.2 [P1] 以证据优化高影响路径

- 状态: `TODO`
- 依赖: TASK-6.1。
- 优先顺序:
  1. source -> MemoryStream -> NPOI DOM、NPOI DOM -> MemoryStream -> destination、failure workbook `ToArray()` 的整块复制；
  2. 无 relation 时 `sourceLocations` 的无意义全量保留；
  3. Plan cache key JSON+UTF8+SHA256+Base64、profile clone、registry lock；
  4. validation range 临时集合、`UniqueTracker.PendingCount()`、CSV 每行配置/Writer、header style cache、动态空 dictionary、`DynamicInvoke`、无界 compiled Regex、`RemoveAll` 中间数组。
- 约束: 不因“0 GC”引入无依据的 Span/ArrayPool/ValueTask/对象池；81,920 byte 缓冲区不默认判定为 LOH；超大文件 streaming/SXSSF 必须作为受限能力另行设计，明确模板/图表/图片/AutoFit 兼容性。
- 验收: 每项优化有前后相同 workload 数据和回归阈值，正确性测试无回退。

## Phase 7: 文档、发布准备和最终审查

### TASK-7.1 [P1] 同步用户文档、XML 注释和示例执行

- 状态: `TODO`
- 依赖: 已批准的 API/行为修改和 TASK-5.1。
- 现有 docs: `docs/excel/README.md`、`mapping-profile.md`、`mapping-json-xml.md`、`import-validation.md`、`dynamic-columns.md`、`nuget-migration.md`。
- 文档扩展目标: `docs/excel/{getting-started,import,export,dynamic-columns,styles-and-layout,templates,multi-sheet-and-relations,validation,images,charts,failure-workbook,performance-and-limits}.md`，`docs/mapping/{precedence,profiles,json,xml,migration}.md`，`docs/csv/{import,export}.md`，`docs/api/compatibility-and-breaking-changes.md`。若不新增单独文件，应明确将每个主题映射到现有文件并确保索引可发现。
- 必须准确说明: 单一推荐调用路径、四种 Profile、JSON/XML、优先级和 Patch、四类 validation mode、模板 overwrite、错误工作簿、stream ownership、CancellationToken、资源限制和大文件边界。
- 示例: Docs tests 使用本地 NuGet 包消费者编译执行；每个 API breaking change 有前后迁移对照。
- 注释: 严格遵守 chinese-comments skill，public API 说明参数、返回、异常、ownership、线程安全和取消；不要为无关遗留代码批量添加注释。
- 验收: docs consumer 编译/运行通过，README 不同时推广多条旧路径，示例与最终 API 一致。

### TASK-7.2 [P0] 清洁验证、包审计和发布结论

- 状态: `TODO`
- 依赖: TASK-7.1、所有 P0/P1 任务。
- 步骤:
  1. 重新执行 TASK-0.1 所列 restore/build/unit/integration/pack/benchmark 命令；在可支持的 TFM 完整覆盖目标矩阵。
  2. 执行 Docs Tests、本地包安装消费、API diff；检查 NuGet dependency、license、symbols、SourceLink、README、package metadata、锁文件。
  3. 在 `final-review.md` 写出 release/no-release 结论、真实风险、BLOCKED 项与解除条件；不得生成“应该通过”。
  4. 建议版本策略：包含 API 删除/方向语义变更的发行使用下一 major 或明确 pre-release；保留兼容期时以 obsolete + migration 文档管理，不静默进入 patch。
- 验收: P0/P1 清零或有证据充分的外部 `BLOCKED`；无 production IVT、无伪异步、mapping end-to-end 通过；未执行任何发布操作。

## 5. 文件影响清单

### 已确认将修改的候选文件

- `src/Bing.Offices.Npoi/Extensions/Extensions.Service.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocument.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocumentFactory.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/IMappingProfileRegistry.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingProfileRegistry.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/CsvStreamExtensions.cs`
- `src/Bing.Offices.Npoi/Extensions/MappingProfileServiceCollectionExtensions.cs`
- `tests/Bing.Offices.Tests/{ExcelP0RegressionTest,MappingConfigurationPatchTest,MappingProfileRegistryTest,PublicApiContractTest,StreamPipelineTest}.cs`
- `tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs`
- `tests/Bing.Offices.Docs.Tests/*`
- `benchmarks/Bing.Offices.Benchmarks/{Program,MappingValidationBenchmarks,StreamPipelineBenchmarks}.cs`
- `docs/excel/*`、`docs/mapping/*`、`docs/csv/*`、`docs/api/*`

### 仅在任务证据确认后修改的候选文件

- `ExcelTypeMapFactory.cs`、`ExcelTypeMap.cs`、`ExcelPropertyMap.cs`、`ExcelValidationBindingFactory.cs`、`ExcelValueConverterBindingResolver.cs`、`ExcelValueMap.cs`。
- `MergedRegionInfo.cs`、`PictureInfo.cs`、`PictureStyle.cs` 及相应 NPOI consumers。
- `CsvEntityPipeline.cs` 与 `ExcelMappingConfigurationLoader.cs` 拆分后的新 internal 文件。
- 项目文件仅在新 Core DI 扩展、测试访问或 docs test 编译需要时最小修改；不得为方便实现引入生产 IVT。

## 6. 实施完成定义

任务仅在以下全部成立时可标记 `DONE`：

1. Phase 1 至 Phase 7 的验收标准满足；
2. `P0-VAL-01`、`P0-VAL-02`、`P0-MAP-01`、`P0-MAP-02`、失败工作簿和模板 P0 均有实际回归证据；
3. P1 无剩余项，或所有外部阻断有原始错误、环境需求与可执行解除步骤；
4. clean restore/build/unit/integration/docs test/pack dry-run/benchmark smoke 的真实结果已记录；
5. 无生产 `InternalsVisibleTo`、无 `Task.FromResult` 伪异步、无静默方向 fallback；
6. API baseline、breaking migration、docs 和 public XML 注释与最终代码一致；
7. 未执行 commit/push/tag/publish，用户自行审查和提交。

## 7. 下一步

使用 `/execute-plan`、`/run-plan` 或 `plan-executor` 从 `TASK-0.1` 开始。执行 Agent 应先读取本计划和当前任务目录中的状态文件；本目录当前仅含本计划，因此 Phase 0 必须先创建其他状态文件并重采集基线。
