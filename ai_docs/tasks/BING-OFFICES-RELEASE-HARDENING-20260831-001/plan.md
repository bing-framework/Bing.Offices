<!-- AI_PLAN_STATUS: READY -->
# Bing.Offices 发布前稳定化与 API 收敛实施计划

## 0. 任务信息

```yaml
task-id: BING-OFFICES-RELEASE-HARDENING-20260831-001
task-type: release-hardening-and-api-convergence
priority: P0
language: zh-CN
plan-status: READY
execution-mode: continuous-and-resumable
breaking-change: allowed-with-approved-api-diff-and-migration
auto-commit: false
auto-push: false
auto-tag: false
auto-publish: false
plan-output: ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/plan.md
```

本轮只创建本计划文件，不修改生产代码、测试、配置、数据库或已有任务资料；不运行实施、打包、测试、Benchmark、Git 提交、推送、Tag、PR 或 NuGet 发布。后续执行 Agent 必须持续实施至最终独立 Review 与 Go/No-Go，不得在 Phase 边界停止。

## 1. 输入、范围与证据完整性

### 1.1 已读取的仓库依据

- 根目录 `AGENTS.md`、`.github/copilot-instructions.md`、`.github/prompts/create-plan.prompt.md` 和 `.github/skills/chinese-comments/SKILL.md`。
- `README.md`、`docs/excel/*.md`、解决方案和项目文件、`common.props`、`framework.props`、`common.tests.props`、`version.props`。
- 当前生产主链：`NpoiExcelImporter`、`NpoiImportPlanBuilder`、`NpoiFailureWorkbookWriter`、`AtomicFileCommitter`、Workbook export request/builder、CSV pipeline 及 metadata/资源限制配置。
- 测试与性能设施：`Bing.Offices.Tests`、`Bing.Offices.Tests.Integration`、`Bing.Offices.Docs.Tests`、`Bing.Offices.ResourceProbe`、`Bing.Offices.Benchmarks`。
- 前序任务 `BING-OFFICES-RELEASE-HARDENING-20260827-001` 的 `plan.md`、`execution.md`、`review.md`、`decisions.md`、`api-diff.md`、`package-consumer.md`。

### 1.2 缺失或冲突的输入

用户提供的以下输入在当前工作区未找到，不能将其中的结论、完成度、快照 SHA 或发现视为已复现事实：

- `ai_docs/codebase-analysis/bing-offices-implementation-review-20260831.md`
- `merged_20260831085402.md`
- 用户给出的快照 SHA-256 `0a64224d62b0106032ec0e1d75eaf73d41cf260a8f66ffb5df99af8e3fdbe9c8` 的对应文件

执行开始时，`RH31-000` 必须再次搜索并尝试核验 SHA-256。若仍缺失，必须在 `00-baseline.md`、`03-decisions.md` 和最终 Review 中标为 `INPUT_MISSING`；只可依当前源码、Git 状态和新运行证据做结论。

用户要求“创建计划后立即开始实施”与当前 `plan-writer` 模式“唯一写入目标为本次 plan.md，禁止实施”冲突。按更高优先级模式约束，本轮在写入计划后停止；后续使用 `$execute-plan` 或等效入口从 `RH31-000` 连续执行。

### 1.3 任务范围

- 生产程序集：`Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi`。
- 质量门禁：单元测试、真实磁盘 Integration、Docs/package-only consumer、Benchmark、XML 文档、打包、兼容与最终独立 Review。
- 不在范围：Word/PDF 新功能、远程发布、远程分支修改、数据库和外部网络服务。

### 1.4 工作区保护

当前工作区已有大量未提交的发布硬化和中文 XML 注释改动。用户特别提示以下文件在前次会话后可能被用户或格式化工具修改，执行任何编辑前必须重新读取并合并现状，绝不回退：

- `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeDataConvertException.cs`
- `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeEmptyLineException.cs`
- `src/Bing.Offices.Core/Bing/Offices/Exceptions/OfficeHeaderException.cs`
- `src/Bing.Offices.Core/Bing/Offices/Metadata/PictureStyle.cs`

## 2. 当前实现、真实接入与完成度判断

### 2.1 技术基线

- 依赖方向为 `Abstractions <- Core <- Npoi`；Abstractions/Core 为 `netstandard2.0`，Npoi 由 `framework.props` 多目标构建。
- Unit 使用 xUnit，目标包括 netcoreapp3.1/net5/net6/net7/net8；Integration 仅 net6/net8；Docs consumer 与 Benchmark 为 net8。
- Core 使用 CsvHelper 33.1.0；Npoi 使用 NPOI；Benchmark 使用 BenchmarkDotNet 0.14.0。
- NPOI 导入在 `WorkbookFactory.Create` 前将输入复制进 `MemoryStream`，因此这是 DOM 管线，不能声明为真正流式、零 GC，或在 NPOI 解压/DOM 创建前阻断全部资源放大。

### 2.2 已进入真实调用链的能力

| 能力 | 现状 | 代码与测试证据 |
| --- | --- | --- |
| XLS/XLSX、异构多 Sheet、动态列、模板、样式、批注、图片、图表 | 已实现 | Workbook Request、NPOI importer/exporter、NPOI 扩展和 Unit/Integration 主链。 |
| Mapping Profile 四种模型、JSON/XML 映射、配置型校验 | 已实现 | Mapping plan factory、loader、Profile Registry、Docs consumer。 |
| Workbook validation、结构化错误、失败工作簿、关系绑定 | 已实现 | `NpoiWorkbookValidationPipeline`、`NpoiFailureWorkbookWriter`、`NpoiRelationBinder`。 |
| CSV 实体导入/导出及公式注入防护 | 已实现 | `CsvEntityPipeline`、`CsvEntityImporter/Exporter`、CSV tests。 |
| Excel/CSV 文件原子提交 | 已实现且接入 | 两个 File API 共用 `AtomicFileCommitter`；存在提交/清理和 Windows 锁定测试。 |
| 请求级 metadata | 已实现且接入 | `ExcelWorkbookMetadataOptions`、`Metadata(...)`、NPOI metadata 应用和模板测试。 |
| CSV 资源限制 | 已实现 | 输入字节、行、错误、字段、列上限，使用 `ResourceLimit` 和截断结果。 |
| API 门禁、Docs package consumer、资源探针、Benchmark 源码 | 基础设施存在 | `PublicApiContractTest`、Docs 包引用、ResourceProbe、BenchmarkDotNet classes。 |

### 2.3 已确认的实现不足

| 优先级 | 状态 | 证据 | 结论 |
| --- | --- | --- | --- |
| P0 | 部分完成 | `NpoiExcelImporter` 先计算 `existingSheets`，导入循环再次调用 `ResolveSheetIndex` | 同一 selector 结果未贯穿 plan、materialization 和 failure mapping，存在未来漂移风险。 |
| P0 | 已实现待复验 | `AtomicFileCommitter` 的 internal 文件系统边界；前序 Review 有双故障和锁定测试记录 | 需用当前工作树复跑 Excel/CSV Move/Replace+cleanup 及 Windows 场景，不能仅依历史报告发布。 |
| P0 | 已实现待复验 | `NpoiFailureWorkbookWriter` 有 `IFailureWorkbookFileSystem`、诊断和清理逻辑 | 需复验 destination 未污染、残留、真实锁定/目录冲突；不使用不稳定 ACL 或磁盘填满。 |
| P1 | 部分完成 | XLS/XLSX template metadata preserve/override 代码和测试存在 | 需补 XLS default preserve、顺序隔离、64 并发重开以及 package consumer 的 `Metadata(...)` 使用证据。 |
| P1 | 未收敛 | `PublicApiContractTest` 仍列出 mapper、plan、type map、validation concrete、legacy helpers/exceptions 等大量 public 类型 | 公开 API 仍混有用户入口、Provider SPI、兼容层和执行细节，需带迁移和 API diff 逐项治理。 |
| P1 | 部分完成 | `ExcelImport.cs` 仍有 `HeaderMatch`、`MaxColumnCount`、`EnabledEmptyLine`、`IgnoreEmptyLineAfterData`；`ExcelExport` 有 `AddNavigationSheet` | 术语和入口冗余；仅 `ExcelSetting.Default` 已移除，其余不得擅自删除。 |
| P1 | 未验证 | Benchmark classes 有 1K/10K/100K workload、缓存/规则/样式/Failure 场景，但当前任务无原始 artifacts | 代码不等于性能结果；必须绑定 commit/diff、环境和原始 JSON/Markdown 重新运行。 |
| P1 | 部分完成 | Docs consumer 以本地 `2.0.0` 包运行；前序文档只将 Bing 包目录设为 source 且依赖预热缓存 | package restore 前提尚不充分，必须区分本地 Bing source 与第三方依赖的可信 source/离线缓存。 |
| P1 | 未验证 | 当前工作区无新代际 Excel/WPS/LibreOffice 互操作产物 | 必须生成本轮 fixtures 并实际打开、保存、重开；缺客户端时明确 `NOT_VERIFIABLE`，不得判 PASS。 |
| P1 | 文档滞后 | README 仍只笼统描述 Excel/Word/Pdf，`nuget-migration.md` 写 `AddNpoi(): void` 而实际 API 应核验 | README、迁移、限制、Stream/File、性能和 release checklist 必须与最终 API 同步。 |

### 2.4 当前完成度和发布判断

用户提供的“约 80%、No-Go”是缺失评审文件的声明，当前不可验证。基于已读取源码和 20260827 独立复审，功能主链大多已实现，发布阻塞点主要集中在 selector 单次解析、P0 故障矩阵复验、metadata 并发/格式证据、API 收敛、性能原始结果、package restore 可复现性、客户端互操作及最终材料。

因此本计划的初始状态为：**实现成熟度约 75% 至 85%，发布成熟度未达到 Go，暂定 No-Go**。`RH31-000` 必须以当前 commit、diff、实际测试与环境数据重算并取代该区间；不得以 interface、DTO、Mock、测试方法名或 README 声明提高完成度。

## 3. 执行报告与证据规则

执行阶段必须在本任务目录创建并持续维护以下文件，均使用 UTF-8：

- `00-baseline.md`
- `01-plan.md`：本文件的执行副本或链接说明，禁止与 `plan.md` 产生冲突方案
- `02-progress.md`
- `03-decisions.md`
- `04-api-breaking-changes.md`
- `05-unit-test-report.md`
- `06-integration-test-report.md`
- `07-package-consumer-report.md`
- `08-benchmark-report.md`
- `09-documentation-report.md`
- `10-final-review.md`
- `11-final-summary.md`

每个完成结论必须可追溯为：`公共 API -> 参数校验 -> Core/内部实现 -> NPOI/CSV/IO -> 异常/取消/资源释放 -> 返回结果 -> Unit -> Integration -> Consumer/Benchmark（适用时）`。

状态只允许：`TODO`、`IN_PROGRESS`、`BLOCKED`、`DONE`、`VERIFIED`、`NOT_VERIFIABLE`。`DONE` 不等于 `VERIFIED`；无当前运行证据的历史实现只能记录为 `DONE` 或 `PARTIAL`。

## 4. Phase 0：基线、输入校验与可追溯性

### RH31-000 基线与缺失输入核验

- **优先级**：P0；**依赖**：无。
- **目标**：冻结实际工作树、环境、依赖、输入快照与前序任务差异，避免在陌生 dirty 改动上做错误结论。
- **确认文件**：根目录 props、所有 `.csproj`、前序任务文档、`PublicApiContractTest.cs`、当前四个用户提示变更文件。
- **步骤**：
  1. 读取并记录 `git status --short`、branch、HEAD、`git diff --stat`、`git diff --name-only`、`git diff --check`；不还原或覆盖任何陌生改动。
  2. 搜索 20260831 review/snapshot，存在时以 UTF-8 读取并计算 SHA-256；不存在时记录 `INPUT_MISSING`。
  3. 记录 `dotnet --info`、OS、CPU、内存、GC、NuGet sources/缓存边界、目标框架与 lock file 状态。
  4. 构建“最终生产符号 -> 测试方法”矩阵，至少覆盖 selector、metadata、AtomicFileCommitter、Failure Workbook、CSV limits、resource probe、public API 和 package consumer。
  5. 创建 `00-baseline.md`、`02-progress.md`、`03-decisions.md`，初始化其余报告模板；每次状态变更附命令、输出摘要、时间、commit/diff hash。
- **测试/验证**：仅基线命令；不因 locked restore 失败修改 lock files。若 `--locked-mode` 失败，记录项目与 `NU1004`，继续使用已恢复资产执行非 restore 验证。
- **风险**：工作区存在用户修改；历史报告中的通过数已过期。
- **验收**：所有输入和环境有来源；未知输入不被当作事实；没有 Git 写操作。

### RH31-001 当前全链回归与发布门禁基线

- **优先级**：P0；**依赖**：RH31-000。
- **目标**：建立本轮可比较的 build/test/pack/docs baseline，并分离项目问题、环境问题和前次修改引入的问题。
- **步骤**：
  1. 优先执行 locked restore；失败仅记录，不用 `--force-evaluate` 改写锁文件。
  2. 执行 Release build、net6/net8 Unit/Integration、Docs consumer，记录总数/失败/跳过和既有警告。
  3. 检查 Release 的三程序集 XML 文档、nupkg/snupkg 的 DLL/XML/nuspec/许可证/README metadata。
  4. 运行 `get_errors`、`git diff --check`；不要将 CRLF/LF 提示误写为补丁错误。
- **验收**：所有命令均有可重跑记录；No-Go/Blocker 清单来自当前结果。

## 5. Phase 1：P0 正确性与失败合同闭环

### RH31-101 Sheet selector 单次解析与稳定冲突合同

- **优先级**：P0；**依赖**：RH31-001。
- **目标**：让请求 selector 到物理 Sheet 的解析只执行一次，并复用于 mapping plan、导入、结果和 Failure Workbook。
- **确认文件**：`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`NpoiImportPlanBuilder.cs`、`ExcelWorkbookImportRequest.cs`、`ExcelImport.cs`、selector Unit/Integration tests。
- **候选文件**：新的 internal resolved-sheet value type/resolver；`NpoiFailureWorkbookWriter.cs`（仅为接受已解析映射而改）。
- **步骤**：
  1. 在 workbook 打开后建立不可变 `request -> physical index -> sheet/name` 结果；保留 missing 状态而不是丢弃后重新解析。
  2. 在任何 plan 编译、反射调度或 `.Single(...)` 前，拒绝 duplicate name、duplicate index、名称+索引命中同一 physical Sheet，并保留 selector/实际索引/名称。
  3. `NpoiImportPlanBuilder` 接受已解析 Sheet 描述而不是 `int` + workbook 再查；import loop、`resolvedSheetRequests` 和 failure writer 复用它。
  4. 统一缺失、隐藏、越界和冲突的 `ExcelImportErrorCode`/异常类型；不泄漏 LINQ `InvalidOperationException`。
- **API/兼容性**：优先 internal 改动；若异常文本或类型为公共合同，写入 `04-api-breaking-changes.md` 并在迁移中说明。
- **用例矩阵**：XLS/XLSX；大小写 comparer；重复名称/索引；name+index 同物理 Sheet；越界、缺失、隐藏；合法混合；相同 ItemType 但不同 mapping；Failure Workbook 回写关联。
- **验收**：主链仅有一次解析；结果在 plan/materialization 一致；所有冲突确定性失败。

### RH31-102 原子文件提交双故障矩阵复验与收口

- **优先级**：P0；**依赖**：RH31-001。
- **目标**：确认 Excel 与 CSV File API 在 write/flush/move/replace/cancel/cleanup 双故障下保持旧目标或不创建新目标，且不丢失主异常与清理诊断。
- **确认文件**：`AtomicFileCommitter.cs`、`ExcelStreamExtensions.cs`、`CsvStreamExtensions.cs`、`ReviewFixRegressionTest.cs`、`ExcelWorkbookRequestTest.cs`、`CsvTest.cs`、Integration test。
- **步骤**：
  1. 不扩展 public IO abstraction；复用 current internal `IAtomicFileSystem`，检查 `CreateNew`、同目录 staging、Flush(true)、Move/Replace 和 cleanup 的实际契约。
  2. 补齐或校正 Excel/CSV 的 Move+Delete 和 Replace+Delete 双故障；断言主提交异常、`Bing.Offices.{format}.TemporaryCleanupException`、Delete 被调用、旧内容/无新目标及 staging 可诊断。
  3. 以真实 Windows locked target 重跑，非 Windows 标为条件性 `NOT_VERIFIABLE`，但不把跳过算 PASS。
  4. 检查路径/父目录为系统边界，不添加路径遍历承诺或在异常中输出业务内容。
- **验收**：双格式、两种提交路径的 Unit 在 net6/net8 通过；Windows 真实磁盘 Integration 通过或明确环境阻塞。

### RH31-103 Failure Workbook 故障、清理和敏感诊断矩阵

- **优先级**：P0；**依赖**：RH31-001。
- **目标**：证实失败工作簿临时目录、创建、序列化、copy、取消、MaxBytes 和删除失败都不污染 destination，并将 cleanup 作为附加诊断而非覆盖主异常。
- **确认文件**：`NpoiFailureWorkbookWriter.cs`、`ExcelImportPolicies.cs`、`ExcelP0RegressionTest.cs`、`ExcelImporterIntegrationTest.cs`。
- **步骤**：
  1. 重新读取用户最近可能修改的目标文件后才编辑；保留当前 `IFailureWorkbookFileSystem` 的最小测试边界。
  2. 对 directory/create/copy failure 验证 destination 长度/字节未变、Delete 调用、临时残留状态、主异常与 inner exception。
  3. 使用真实目录冲突和 Windows 独占目标流覆盖真实 IO；不改 ACL、不填满磁盘。若 Windows 特定场景不可执行，记录 `NOT_VERIFIABLE`。
  4. 审计 DiagnosticSink 和 Exception.Data：不得含 cell 原值、表头、实体或工作簿业务内容；sink 抛异常不得遮蔽主异常。
  5. 验证 ErrorRowsOnly/AnnotatedOriginal 与 XLS/XLSX 受限写入不回归。
- **验收**：Unit 故障矩阵和真实磁盘场景有当前证据；失败输出、临时残留及敏感信息合同被断言。

### RH31-104 metadata 格式、隔离与消费者证据

- **优先级**：P1；**依赖**：RH31-001。
- **目标**：验证 request metadata 快照在 XLS/XLSX、template/no-template、顺序和并发导出下都隔离；不恢复 `ExcelSetting.Default` 静态共享状态。
- **确认文件**：`ExcelWorkbookMetadataOptions.cs`、`ExcelWorkbookExportRequest.cs`、`ExcelExport.cs`、`ExcelHelper.cs`、`NpoiExcelExporter.cs`、`ExcelWorkbookRequestTest.cs`、Docs consumer。
- **步骤**：
  1. 补 XLS 默认 preserve，核验六字段显式 override 的格式差异。
  2. 验证 `Build()` clone 后修改原 options 不污染 request；连续导出与 64 个不同 metadata 并发导出后重开逐一校验。
  3. Docs package-only consumer 增加 `Metadata(...)` 的编译和实际重开调用。
  4. 将 `ExcelSetting` 的现存实例型遗留定位为兼容 DTO 或删除候选；不在没有 API 分类账前继续 breaking。
- **验收**：无生产读取 static mutable metadata；格式、顺序、并发、consumer 均有 current-run 证据。

### RH31-105 资源、异常、Dispose 和取消边界

- **优先级**：P0；**依赖**：RH31-001。
- **目标**：确保 CSV/Excel 上限、错误分类、反射异常解包、Stream ownership、取消和 Dispose 与代码及文档一致。
- **确认文件**：`CsvEntityPipeline.cs`、`CsvHeaderBinder.cs`、`CsvImportOptions.cs`、`NpoiExcelImporter.cs`、`NpoiStreamCopier.cs`、`NpoiRelationBinder.cs`、resource probe/tests。
- **步骤**：
  1. 验证 CSV byte/row/error/field/column limits 对 seekable/non-seekable 输入、header 和后续 record 一致产生 `ResourceLimit` + truncation。
  2. 覆盖 enumerator/parser/writer 在 MoveNext、转换器、setter、validation 和 cancellation 抛异常时 Dispose；不 mock 内部业务。
  3. 审计反射 `TargetInvocationException`，保留原异常 stack；统一结构、配置、资源、取消、输出错误分类。
  4. 为 Excel 明确：MaxInputBytes 只限制复制前输入；图片上限仅覆盖映射图片列；NPOI DOM/ZIP/OLE 边界通过独立进程取证但不做不存在的压缩炸弹承诺。
  5. 不添加伪 async；量化 NPOI 不可中断阶段前后的 cancellation 检查，记录 ADR。
- **验收**：错误码、资源结果、stream ownership 和释放行为由 provider 主链测试证明；无 catch-all 静默 fallback。

## 6. Phase 2：API 收敛与 Breaking Change 治理

### RH31-201 API 分类账、consumer 搜索与 SPI 边界

- **优先级**：P0；**依赖**：Phase 1 合同冻结。
- **目标**：将 public 类型/成员分为用户 API、Provider SPI、兼容层、执行细节、无消费者候选；避免仅靠 `EditorBrowsable` 伪隐藏。
- **确认文件**：`PublicApiContractTest.cs`、三程序集 public API、`AssemblyInfo.cs`、Docs tests、NPOI DI registration。
- **候选收敛项**：`MappingConfigurationMerger`、document factory/loader concrete、`ExcelTypeMap*`、`ExcelPropertyMap`、plan factory/provider、value map/converter resolver、concrete validation rules、`UniqueTracker`、`SheetSetting`、`RegexConst`、旧 helpers、legacy exceptions、`ICellValueConverter`。
- **步骤**：
  1. 生成每个 public/protected 类型和成员的源码调用、测试调用、Docs fence、nupkg consumer、Provider 扩展证据。
  2. 将对外 SPI 限缩为 provider-neutral 只读 contract；生产程序集不得通过 IVT 向彼此穿透。
  3. 对无外部场景的执行细节改 internal/private；对可能被第三方使用的类型先给 `[Obsolete]` forwarding 或 next-major migration。
  4. 更新 public API baseline，增加 negative API test，检查 NPOI 类型泄漏和 mutable collection 暴露。
- **Breaking 策略**：每个删除/重命名必须写 old/new、最早移除版本、迁移代码、是否 source/binary breaking、consumer 验证；没有证据不得删除。
- **验收**：每个 exported type 有明确分类；NPOI 程序集只暴露批准的注册入口；无 production IVT。

### RH31-202 唯一推荐入口与术语治理

- **优先级**：P1；**依赖**：RH31-201。
- **目标**：减少同义 builder/option 名称，且使文档只有一条推荐主路径。
- **候选变化，必须先经分类账确认**：
  - `ExcelMapping.For<T>()` -> 方向化 `Import<T>()`/`Export<T>()`。
  - `Mapping(configuration/document)` -> 具名配置与文档入口。
  - `HeaderMatch` -> `RequireExpectedHeaders`；`MaxColumnCount` -> `MaxReadColumns`。
  - `EnabledEmptyLine` -> `ReportEmptyRows`；`IgnoreEmptyLineAfterData` -> `StopAtFirstEmptyRow`。
  - `AddNavigationSheet` -> `AddSheet(name, parents.SelectMany(...))`。
  - `AddNpoi()` 返回 `IServiceCollection`，前提是实际签名与二进制兼容影响确认。
- **步骤**：
  1. 逐项搜索仓库和 package consumer，决定保留、obsolete forwarding 或 next-major 删除。
  2. 对 bool 名称记录默认值、true/false 行为和 CSV/Excel 一致性；不能仅机械 rename。
  3. 迁移 Docs fence、README、API baseline 和 consumer；不保留两套独立实现。
- **验收**：每项有行为对照、API diff、migration snippet、Docs/consumer 编译证据；没有未批准的额外 breaking。

### RH31-203 API 契约自动化与版本治理

- **优先级**：P1；**依赖**：RH31-201/202。
- **目标**：将当前巨大手写类型列表升级为可维护的 public API 差异门禁，并与版本策略一致。
- **步骤**：评估并接入与现有多 TFM/lock file 兼容的 APICompat/PublicApiAnalyzers，或增强当前反射 member snapshot；为每个目标框架生成 diff；核验 `2.0.0` 与所有 breaking 的版本策略，未经用户决定不得自行升 major。
- **验收**：CI 能阻止意外 API 变更；`04-api-breaking-changes.md` 是唯一真实 migration ledger。

## 7. Phase 3：职责、目录与热路径重构

### RH31-301 导入解析、计划与物化分层

- **优先级**：P1；**依赖**：RH31-101、Phase 2 API 冻结。
- **目标**：使 `NpoiExcelImporter` 清晰表达 resolve -> plan -> validate -> materialize -> relation -> failure artifact，降低反射和重复解析耦合。
- **步骤**：按真实职责抽 internal resolved-sheet resolver、generic dispatch cache、column binder、resource runtime/source location；一个主要 public 类型一个文件；保留 NPOI adapter internal。
- **测试**：每个 collaborator 直接 Unit；facade 用真实 workbook Integration；cache hit/miss 和 exception unwrapping。
- **验收**：没有新的 public surface；main orchestrator 不再承担 selector、mapping、row materialization、relation 和 IO 全部细节。

### RH31-302 导出、Failure Workbook 与 CSV 边界复用

- **优先级**：P1；**依赖**：RH31-102/103、Phase 2。
- **目标**：导出/失败产物/CSV 共用明确的 output commit 语义，避免 duplicated staging/cleanup 逻辑；只提取能减少复杂度的协作对象。
- **步骤**：Exporter 只管 plan/workbook lifecycle；Failure Workbook 分解复制、rich text/style/drawing、summary 和 output lifecycle；CSV import/export/record reader/writer 分文件并维持单操作生命周期。
- **验收**：不引入多层空 abstraction；Failure Workbook/ErrorRowsOnly golden 和 File API 合同保持。

### RH31-303 Mapping plan/cache 与性能敏感实现

- **优先级**：P1；**依赖**：Phase 2、RH31-501 baseline。
- **目标**：根据 Benchmark/profiler 处理 generic reflection dispatch、accessor、ConversionContext、mapping cache key、ValidationRangeIndex 与 dynamic type cache。
- **步骤**：先为 cache hit/miss/eviction/concurrency、clone/merge 字段完整性建立测试；仅在有测量时编译访问器、预索引转换器/命名规则、减少无 converter context 分配、限定 dynamic cache。
- **验收**：优化前后结果一致；资源/性能收益和复杂度写入 benchmark report；不默认使用对象池/ValueTask/Source Generator。

### RH31-304 死代码、命名空间和兼容遗留清理

- **优先级**：P2；**依赖**：RH31-201～303。
- **目标**：消除明确无引用 helper、宽泛 fallback 和目录/namespace 不一致（例如拼写问题），不进行全仓格式化。
- **验收**：每个删除有搜索与 negative API 证据；没有超出批准 breaking table 的 public 改动。

## 8. Phase 4：测试、Integration 与 PackageConsumer

### RH31-401 P0 单元测试与生产符号追溯

- **优先级**：P0；**依赖**：Phase 1～3 相关实现。
- **要求**：xUnit 测试方法使用英文 `Method_State_Expected`，中文 XML 测试目的，AAA；仅 mock 文件系统、时间、IO 等外部边界，不 mock 被测内部实现。
- **最低矩阵**：selector、metadata 64 并发、AtomicFileCommitter 双故障、Failure Workbook、CSV limit/release、resource classification、formula cache、reflection/relations、mapping cache、API negative baseline。
- **验收**：`05-unit-test-report.md` 能映射每个新增/修改生产符号到精确测试方法；net6/net8 P0 无跳过。

### RH31-402 Integration、golden 与独立进程

- **优先级**：P0；**依赖**：RH31-401。
- **范围**：真实 XLS/XLSX 磁盘导入/导出、多 Sheet、template、style/comment/picture/chart、failure artifact reopen、atomic target lock、目录冲突、Windows locked destination、resource probe 子进程。
- **约束**：不依赖公网、生产数据库、sleep、随机等待或填满磁盘；Windows 特定断言需 `RuntimeInformation` gating，并以 `NOT_VERIFIABLE` 不取代 PASS。
- **验收**：`06-integration-test-report.md` 区分已运行、条件跳过、环境阻塞与失败。

### RH31-403 可复现 PackageConsumer

- **优先级**：P0；**依赖**：Phase 2 API 和本地 pack。
- **目标**：验证 nupkg 而不是项目引用，且文档命令能在声明条件下恢复 Bing 包与第三方包。
- **确认文件**：`Bing.Offices.Docs.Tests.csproj`、`DocsConsumerTest.cs`、`package-consumer.md`、pack 配置。
- **步骤**：
  1. 将本轮输出放入临时本地 package source，确保 assets 精确解析本轮版本，无 `NU1601`。
  2. 明确两种受支持方式：预热离线 `NUGET_PACKAGES`（记录其已含第三方依赖），或同时提供受信任第三方 source；不得声称空缓存仅凭本地 Bing source 能 restore。
  3. Docs consumer 必须实际调用唯一主 API、metadata、mapping、CSV、DI 和 XML docs；必要时新增临时独立 consumer，但不提交无价值 demo 项目。
  4. 检查 nupkg/snupkg 含 DLL、XML、nuspec 依赖、license/readme/repository metadata。
- **验收**：`07-package-consumer-report.md` 包含 source/cache 前提、assets 证据、命令和测试结果。

### RH31-404 全 TFM 构建与运行策略

- **优先级**：P1；**依赖**：RH31-001。
- **目标**：NPOI 的 netstandard2.0/2.1、netcoreapp3.1、net5、net6、net7、net8 和测试目标真实建构；安装 runtime 的测试必须运行。
- **验收**：缺 runtime 记录为环境 `BLOCKED`；net6/net8 Unit/Integration 是强制全绿，不将只 build 当作 test PASS。

## 9. Phase 5：性能、GC、内存与 Benchmark

### RH31-501 Benchmark 基线可信度与工作负载校准

- **优先级**：P0；**依赖**：正确性冻结。
- **确认文件**：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`MappingValidationBenchmarks.cs`、`StreamPipelineBenchmarks.cs`、ResourceProbe。
- **目标**：从实际参数进入工作负载的源码，生成绑定 commit/diff/environment 的原始 BenchmarkDotNet JSON/Markdown。
- **步骤**：
  1. 审核 `RowCount`、`FailureRowCount`、规则数、cache cold/hit/miss/eviction、Unique rows/columns 确实参与运行路径。
  2. 删除或隔离仅测 `PeakWorkingSet64` 的微基准结论；峰值使用独立子进程采样，区分 retained 和 peak，使用 `GC.KeepAlive`。
  3. 使用 BenchmarkDotNet 默认自适应，或明确记录 smoke 的 `SimpleJob(1,2,3)` 不能作为发布性能结论；正式基准禁止以三 iteration 代表 100K release budget。
  4. 先仅运行相关 benchmark filters，避免盲跑超大矩阵；储存原始 artifacts。
- **验收**：每个指标可追溯到源码输入；`08-benchmark-report.md` 不宣称 zero-GC/streaming/压缩炸弹防护。

### RH31-502 性能优化与预算

- **优先级**：P1；**依赖**：RH31-501。
- **目标**：只优化测量证实的热点，并在相同环境比较前后。
- **候选**：compiled accessor、反射 generic dispatch cache、ConversionContext 延迟创建、mapping/cache key、ValidationRangeIndex、ArrayPool finally 归还、导出 byte copy、style/rich text/picture/AutoSize 成本、bounded dynamic type cache。
- **规模与指标**：Excel 1K/10K/100K、CSV 大字段/大行、mapping 10/100/1K/10K rules、validation 重叠范围、failure/template/picture/chart、并发 1/4/16/64；Mean、error、alloc/op、Gen0/1/2、LOH、independent peak working set、throughput、tail latency、destination capacity。
- **验收**：每个优化对应事实基准/探针、正确性回归和 ADR；未达到的预算据实保留 No-Go 或 waiver。

## 10. Phase 6：文档、发布材料与最终独立 Review

### RH31-601 文档与 XML Documentation 同步

- **优先级**：P1；**依赖**：最终 API。
- **范围**：README、`docs/excel`、NuGet migration、XML docs。
- **步骤**：
  1. 按中文注释 Skill 校验改动 public API、构造、参数、返回、异常与 `inheritdoc`；删除不实“streaming/zero GC/保证回滚”说明。
  2. 更新唯一推荐路径、metadata、selector、mapping migration、CSV/Excel limits、template/style/comment/chart、failure workbook、Stream ownership、File atomicity、取消边界、性能复现和已知不支持项。
  3. `AddNpoi`、版本号、API 名称、支持 TFM 必须从实际源码/pack 生成，不使用过期文字。
  4. 所有 C# fence 继续由 Docs tests 从 markdown 编译执行；链接有效、无占位版本或重复导航。
- **验收**：`09-documentation-report.md` 记录每个页面和 fence 验证；README 不再声称当前未支持的 Word/PDF 能力。

### RH31-602 本地 pack、SBOM 与发布清单

- **优先级**：P0；**依赖**：RH31-403、RH31-601。
- **目标**：只生成本地产物和审计材料，不执行任何远程发布。
- **步骤**：验证三包和符号包、依赖/许可/SBOM、XML docs、README/repository metadata、supported TFM/runtime/OS、NPOI/CsvHelper 版本、API diff、breaking migration、release notes、known limitations。
- **验收**：发布清单可由清洁环境复现；任何遗漏保持 No-Go。

### RH31-603 办公客户端互操作

- **优先级**：P1；**依赖**：RH31-402。
- **目标**：以本轮生成的 XLS/XLSX fixture 验证 Excel、WPS 或 LibreOffice 的打开/保存/重开及 NPOI 二次解析。
- **步骤**：生成具有 manifest、source hash、generation id 的 template/style/comment/picture/chart/failure fixtures；自动化客户端若可用则运行，否则记录版本、缺失原因和人工步骤。
- **验收**：客户端不可用为 `NOT_VERIFIABLE` 而非 PASS；不复用旧任务产物充当当前证据。

### RH31-604 最终独立 Review 与 Go/No-Go

- **优先级**：P0；**依赖**：全部可执行任务。
- **目标**：由独立 reviewer 只读复核真实调用链、测试、包、性能、文档和工作区 diff，并得到发布判断。
- **Review 重点**：参数校验、异常、取消、Dispose、并发、API 重复、NPOI 泄漏、production IVT、伪 async、无界缓存、静态状态、Benchmark 参数真实性、dirty state/输入缺失。
- **Go 条件**：P0/Blocker 为零；P1 已修复或有具名且被批准的 release waiver；net6/net8 Unit/Integration、Docs、pack/consumer、相关 benchmark 全部有当前证据；互操作按发布政策处理；最终 docs/API/migration 一致。
- **输出**：`10-final-review.md` 和 `11-final-summary.md`，包含完成度、验证矩阵、残余风险、waiver、明确 Go/No-Go 与建议提交分组；不得自动 commit/push/tag/publish。

## 11. Breaking Change 与迁移控制表

| 候选变更 | 当前状态 | 计划策略 | 必须证据 |
| --- | --- | --- | --- |
| `ExcelSetting.Default` 移除 | 已在 2.0.0 工作树发生 | 复验迁移至 request `Metadata(...)`，不恢复共享状态 | API diff、XLS/XLSX/并发、consumer |
| `ExcelMapping.For<T>()` 方向化 | 候选 | 未经 API 分类账和用户版本决定不得删除 | 仓库/consumer 搜索、obsolete 或 next-major 表 |
| `Mapping(...)` 重载收敛 | 候选 | 具名入口后单实现，避免双链 | behavior diff、Docs/consumer |
| header/column/empty-row bool 命名 | 候选 | 记录默认与 true/false 语义，渐进 obsolete | Unit、API migration |
| `AddNavigationSheet` | 候选 | 仅在 `AddSheet+SelectMany` 输出等价时弃用 | golden/consumer |
| `ICellValueConverter` 与 legacy attributes/exceptions | 候选 | 先确认第三方 Provider/consumer，再定 adapter 或删除 | public API/consumer |
| plan/type map/validation concrete public -> internal | 候选 | 分离 SPI 与执行细节，不靠 IVT | negative API test、provider proof |
| `AddNpoi` 返回 `IServiceCollection` | 待核验 | 先检查当前签名、二进制影响和链式使用 | DI/consumer/API diff |

## 12. 验证命令

执行 Agent 先根据 RH31-000 确认环境，以下命令均来自现有 solution/project 配置。`--locked-mode` 失败时不得改 lock file；后续 `--no-restore` 仅可在资产确实可用时运行。

```powershell
dotnet --info
dotnet restore Bing.Offices.sln --locked-mode
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-build
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net6.0 --no-build
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-build
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-build
dotnet pack src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release --no-build --no-restore -o artifacts/packages-vnext
dotnet pack src/Bing.Offices.Core/Bing.Offices.Core.csproj -c Release --no-build --no-restore -o artifacts/packages-vnext
dotnet pack src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-build --no-restore -o artifacts/packages-vnext
dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks -- --filter "*StreamPipelineBenchmarks*" --exporters json markdown
git diff --check
```

- ResourceProbe、FailureWorkbook、Mapping/Validation、Regex/Unique 等 benchmark filter 必须按当前任务分批执行，并在报告记录完整参数与 artifacts 路径。
- Package consumer restore 必须明确 `NUGET_PACKAGES`、本地 Bing source 和第三方依赖 source/预热 cache 前提，不能以空缓存和单一本地 Bing source 的失败作为库缺陷。
- 不调用 `Publish.bat` 或 build 脚本中的发布 target，防止触发 NuGet publish。

## 13. 完成定义与依赖顺序

执行顺序：`RH31-000 -> RH31-001 -> Phase 1 P0 -> RH31-201 -> RH31-202/203 -> Phase 3 -> Phase 4 -> Phase 5 -> Phase 6 -> RH31-604`。不依赖项可并行，但每个 Phase 的 P0 失败必须阻止 Go 结论。

任务仅在以下条件同时满足时标记 `COMPLETE`：

1. 所有 P0 项达到 `VERIFIED`；P1 已修复或得到具名、带风险和批准人的 waiver；P2 有明确归属。
2. selector、metadata、原子文件、Failure Workbook、CSV/Excel 资源与释放合同有当前源码与运行证据。
3. API 分类与迁移账本完整，无未记录 public 变化、NPOI 泄漏或 production IVT。
4. net6/net8 Unit/Integration、Docs consumer、Release build、pack 和 diff 检查通过；其它 TFM 真实 build，缺 runtime 如实 `BLOCKED`。
5. Benchmark 原始 artifacts 与环境数据齐全，性能声明不超过可测范围。
6. XML docs、README、Docs fences、migration 和 release checklist 与最终 API 行为一致。
7. 互操作结果或不可验证原因按发布政策记录；未验证客户端不冒充 PASS。
8. `10-final-review.md` 和 `11-final-summary.md` 给出证据化 Go/No-Go；没有执行 commit、push、PR、Tag 或 NuGet publish。

## 14. 建议提交分组（仅供最终报告，不执行）

1. `fix(import): resolve sheet selectors once and reject physical conflicts`
2. `test(io): close atomic export and failure workbook fault matrices`
3. `test(export): prove metadata isolation across formats and concurrency`
4. `refactor(api): classify and converge public entry points`
5. `refactor(npoi): separate import orchestration from plan and materialization`
6. `perf(core): measure and optimize proven mapping and validation hot paths`
7. `test(release): validate package consumers, TFM and client interoperability`
8. `docs(release): align migration limits and release evidence`

这些仅是建议分组，不授权 Git 提交、推送、Tag、PR 或发布。
