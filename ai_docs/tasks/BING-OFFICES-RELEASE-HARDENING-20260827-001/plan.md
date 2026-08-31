<!-- AI_PLAN_STATUS: READY -->
# Bing.Offices 发布前正确性、API、测试、性能与文档闭环实施计划

## 0. 任务信息

```yaml
task-id: BING-OFFICES-RELEASE-HARDENING-20260827-001
task-name: Bing.Offices 发布前正确性、API、测试、性能与文档闭环
task-type: implementation-and-release-hardening
priority: P0
language: zh-CN
execution-mode: continuous-and-resumable
plan-status: READY
breaking-change: allowed-with-api-diff-and-migration
auto-commit: false
auto-push: false
auto-tag: false
auto-publish: false
```

- **唯一任务目录**：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/`。
- **本计划阶段边界**：只创建本 `plan.md`，不修改业务代码、测试、配置、数据库，不执行实现、Git 提交、推送、Tag 或 NuGet 发布。
- **后续执行方式**：由执行 Agent 从 `RH27-000` 开始持续执行 Phase 1～6；单项阻塞时记录证据并继续不依赖项，不在 Phase 边界自动停止。
- **发布安全**：实施阶段允许 restore、build、test、pack、PackageConsumer、Benchmark 和本地互操作取证；任何 commit、push、tag、publish 仍需用户另行明确授权。

## 1. 输入、证据与缺口

### 1.1 已读取并验证的依据

1. 用户附加的完整发布硬化需求，固定 Task ID 为 `BING-OFFICES-RELEASE-HARDENING-20260827-001`。
2. 根目录 `AGENTS.md`、`.github/copilot-instructions.md`、`.github/prompts/create-plan.prompt.md` 和 `.github/skills/chinese-comments/SKILL.md`。
3. `Bing.Offices.sln`、`common.props`、`common.tests.props`、`framework.props`、三个生产项目、Unit/Integration/Docs/Benchmark 项目和 Flubu `build/BuildScript.cs`。
4. `README.md`、现有 `docs/excel/*.md`、生产源码、公共 API 快照、关键单元/集成/文档测试和 Benchmark 源码。
5. 前序任务 `BING-OFFICES-RELEASE-HARDENING-20260825-01` 的 `plan.md`、`execution.md`、`review.md`，用于识别已完成修复、遗留 blocker 和过时结论。

### 1.2 输入缺口与冲突处理

- 用户指定的 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260827.md` **当前不存在**。不得编造其结论；执行开始时再次检查，仍缺失则在 `baseline.md` 和 `decisions.md` 记录证据缺口，以当前源码、可复现测试及前序独立 Review 为准。
- 仓库中未发现 `Mustang.Solution`、`vendor/methodologies` 或其他方法论归档，无法记录其 commit/SHA-256。执行时再次搜索；若仍不存在，记录“不适用/输入缺失”，不得以方法论替代运行证据。
- 当前计划工具没有执行任意终端命令的能力，因此本轮没有声称已运行 `git status`、`git diff`、SDK 查询、build/test/pack。`RH27-000` 必须首先补齐这些真实基线。
- 用户附加提示词要求“创建计划后立即实施”，但当前严格角色为 `plan-writer`，只允许写 `plan.md` 并在完成后停止。实施由后续 `/execute-plan`、`/run-plan` 或 `$execute-plan` 接续，不改变任务 ID。
- 前序 Review 的部分性能缺陷与当前 Benchmark 源码存在时间差：当前源码已经拆出 Failure/HeaderStyle/ValidationRange/Regex/Unique 基准并让参数进入 workload，但尚未读取到与当前源码对应的新原始结果。因此只能标为“代码已调整，证据待重跑”，不能标为 VERIFIED。

## 2. 当前实现与完成度判断

### 2.1 技术与项目基线

- 生产依赖方向：`Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`。
- Abstractions/Core 目标为 `netstandard2.0`；NPOI 为 `netstandard2.0;netstandard2.1;netcoreapp3.1;net5.0;net6.0;net7.0;net8.0`（由 `framework.props` 决定）。
- Unit 测试使用 xUnit，目标为 netcoreapp3.1/net5/net6/net7/net8；Integration 为 net6/net8；Docs consumer 与 Benchmark 为 net8。
- NPOI 导入先将输入复制到 `MemoryStream`，随后使用 `WorkbookFactory` 创建 DOM。它不是流式 XLSX parser，也不能承诺 0 GC、低峰值或在 DOM 创建前阻断所有解压后资源放大。
- CsvHelper 33.1.0、NPOI（版本来自 `version.props`/锁文件，执行时记录解析值）、BenchmarkDotNet 0.14.0。

### 2.2 已真实进入调用链的能力

| 能力 | 当前判断 | 证据摘要 |
| --- | --- | --- |
| XLS/XLSX 导入导出、异构多 Sheet | 已实现 | Workbook Request、NPOI importer/exporter、Unit/Integration 调用链存在。 |
| 固定列、typed 动态列、Mapping Profile/JSON/XML | 已实现 | Mapping plan factory、方向 Builder、文档 consumer 与大量请求测试存在。 |
| 配置优先级、Reset/Clear、缺失方向 fallback 策略 | 已实现并有回归 | Mapping patch/profile 测试及前序整改记录存在。 |
| 配置校验、Workbook Data Validation 子集、Unique | 已实现子集 | Workbook validation pipeline、Unique journal、XLS/XLSX 操作符测试存在；复杂公式/命名区域仍需扩展验证。 |
| 关系绑定、图片、模板、样式、批注、失败工作簿 | 已实现但能力有边界 | Provider 主链和真实工作簿测试存在；ErrorRowsOnly 不承诺复制全部 Excel 对象。 |
| XLSX 柱/线/饼图 | 已实现子集 | 导出定义和 NPOI 写入链存在；XLS 和复杂图表场景不是当前承诺。 |
| CsvHelper 实体导入导出 | 已实现基础 | 单次 CsvWriter、公式注入策略、Unique 与测试存在；资源上限和异常释放矩阵不完整。 |
| 公共 API 快照、Docs consumer、Integration、Benchmark | 已建立 | `PublicApiContractTest`、Docs Tests、Integration、BenchmarkDotNet 项目存在。 |

### 2.3 前序任务已修复、不得重复设计的行为

执行时先以当前测试重验，保持以下回归合同：

- `AddNpoi()` 注册完整默认规则；Configured/Workbook validation mode 隔离。
- Workbook Data Validation 的全部比较操作符、`LESS_OR_EQUAL` Formula1、HSSF 数值边界、`EmptyCellAllowed`。
- Workbook 校验失败不进入转换、配置校验或实体物化；Unique pending 正确回滚。
- `ExcelDateAttribute.Format` 基于原始文本精确校验。
- Failure Workbook 在临时文件中受 `MaxBytes` 限制，超限不污染 destination；非零表头 ErrorRowsOnly 正确。
- Failure Workbook 批注 Preserve/Append/Replace/Fail、XLS/XLSX 富文本 run、部分结构复制。
- `ExcelCellStyleReset`、XSSF 前景/背景及边框颜色、HSSF 颜色能力边界和样式缓存。
- Regex 缓存容量 256、FIFO、timeout、并发边界。
- Unique pending 增量计数、图片行索引、导出 dynamic key 预计算、HeaderAttribute 样式/字体缓存。

### 2.4 当前明确缺口

| 优先级 | 缺口 | 当前证据 |
| --- | --- | --- |
| P0 | 重复物理 Sheet selector 未统一拒绝 | Builder 只拒绝重复名称；重复索引、名称+索引命中同一 Sheet 会进入 importer，`NpoiImportPlanBuilder` 仍可能在 `.Single(...)` 处产生非契约异常。 |
| P0 | Workbook metadata 是进程级可变状态 | `ExcelSetting.Default` 有可替换静态实例；`ExcelHelper.PrepareWorkbook` 在未显式传值时读取该全局对象。 |
| P0 | 普通导出 File API 非原子 | `ExcelStreamExtensions.ExportToFile` 直接 `FileMode.Create` 目标文件；转换/序列化/写失败会留下截断目标。 |
| P0 | Stream 失败语义未完整合同化 | exporter 直接写 destination；不可寻址或中途失败流无法回滚，文档与测试不足。 |
| P0 | 失败工作簿临时文件清理失败被吞掉 | `NpoiFailureWorkbookWriter` 对 `IOException`/`UnauthorizedAccessException` 静默忽略，缺少结构化诊断与故障注入。 |
| P0 | 资源限制描述与真实边界不完整 | `MaxInputBytes` 只限制压缩/文件输入复制；DOM、shared strings、styles/drawings 和压缩炸弹边界尚无准确合同。CSV 缺少输入字节、行、错误、字段长度等完整门槛。 |
| P0 | CSV enumerator 异常释放需证明 | `CsvEntityPipeline` 手工取得 enumerator；需验证所有异常/取消路径都 Dispose。 |
| P0 | 公式缓存边界未完整覆盖 | 当前读取缓存值且暴露 provider-neutral context，但无缓存、陈旧缓存、错误缓存的最终合同和矩阵不完整。 |
| P1 | 公共 API 约 180 个顶层类型，执行细节公开 | 快照包含 `MappingConfigurationMerger`、具体 mapping/type map、内置校验实现、legacy exceptions/helpers 等大量候选。 |
| P1 | API 入口和术语重复 | `ExcelMapping.For<T>()`、方向 Builder、两个 `Mapping(...)`、`HeaderMatch`、`MaxColumnCount`、`EnabledEmptyLine`、`IgnoreEmptyLineAfterData`、`AddNavigationSheet` 并存。 |
| P1 | 生产源码仍有不实“单遍流式”注释 | importer/exporter 的 XML summary 与 DOM/直接写语义不一致。 |
| P1 | 反射与热路径成本仍需量化 | generic `MethodInfo.Invoke`、逐 cell reflection accessor、ConversionContext、mapping cache key、ValidationRangeIndex 等需 benchmark 决策。 |
| P1 | 办公客户端互操作仍阻塞 | 前序记录中 Excel COM `Workbooks.Open` 失败，LibreOffice/WPS 不可用，没有可靠 roundtrip PASS。 |
| P1 | 文档主题不完整 | 现有 docs 仅 6 个页面；Stream/File、模板/样式/图表、关系/图片、容量、性能和 release notes 尚未闭环。 |

### 2.5 总体评估

沿用用户提供的静态基线作为待复核起点：功能 85%、API 61%、测试 73%、性能 42%、发布准备度 54%、总体 78%。当前源码显示前序任务已显著加强正确性与测试，但 P0 物理 Sheet 冲突、静态 metadata、普通文件原子性、资源/临时文件合同仍未闭环，API 和性能证据也未完成。因此本计划开始时的真实发布判定仍为 **No-Go**；完成度百分比必须在 `baseline.md` 用最新 build/test/API/benchmark 证据重算，不能直接继承静态数字。

## 3. 总体执行策略与状态规则

- 状态只使用 `TODO`、`IN_PROGRESS`、`BLOCKED`、`DONE`、`VERIFIED`。
- `DONE` 只表示代码完成；调用链、测试、包或产物未通过时不得标 `VERIFIED`。
- 每个 Phase 后执行 `REVIEW -> FIX -> REGRESSION -> REVIEW`，BLOCKER/P0/P1 未闭环不得进入最终 Go。
- Phase 1 优先消除数据错误、跨请求污染、目标截断和临时文件残留；不得先做目录大迁移。
- Phase 2 的 breaking table 属于用户已明确允许的 next-major 治理输入，不再把“未单独批准”作为无限延期理由；但每个删除项仍必须有仓库消费者搜索、API diff、迁移路径和 PackageConsumer 证据。
- 目录重构只在 Phase 2 API 决策后进行，避免同一文件同时承担行为、API 与机械移动三种 diff。

## 4. Phase 1：正确性、失败语义与资源边界

### RH27-000 基线、工作区保护与追溯矩阵

- **优先级**：P0；**依赖**：无。
- **目标**：建立当前 commit/dirty state、环境、API、测试、包、Benchmark 和前序任务差异基线。
- **已确认范围**：`Bing.Offices.sln`、所有 `.csproj`/props/targets、`PublicApiContractTest.cs`、前序任务文档。
- **步骤**：
  1. 运行 `git status --short`、`git diff --stat`、`git diff --name-only`、`git diff --check`；保护所有陌生改动。
  2. 记录 branch/commit、SDK/runtime/OS/CPU/内存/GC、目标框架和锁定依赖；再次搜索缺失评审与方法论。
  3. 执行 restore、Release build、net6/net8 Unit/Integration、Docs、pack 基线；缺失旧 runtime 时只记录环境阻塞，不弱化项目目标。
  4. 生成 public 顶层类型/成员基线、生产 IVT 列表和“生产符号 -> 测试方法”矩阵。
  5. 创建并持续维护 `baseline.md`、`progress.md`、`decisions.md`、`api-governance.md`、`test-matrix.md`、`test-report.md`、`benchmark-plan.md`、`review.md`、`release-checklist.md` 及 artifacts 索引。
- **风险**：前序任务可能仍是未提交工作树；所有测量必须绑定 commit + diff hash。
- **验收**：环境、命令、通过/失败/跳过数及产物路径可复现；报告缺失和 dirty state 被显式记录。

### RH27-101 重复物理 Sheet selector 前置拒绝

- **优先级**：P0；**依赖**：RH27-000。
- **目标**：在行解析和 mapping `.Single(...)` 前，区分 selector 重复、物理 Sheet 重复和映射冲突。
- **已确认文件**：`ExcelImport.cs`、`ExcelSheetSelector*`、`NpoiExcelImporter.cs`、`NpoiImportPlanBuilder.cs`、`ExcelWorkbookRequestTest.cs`、Integration tests。
- **步骤**：
  1. Builder 前置拒绝重复名称（按配置 comparer）和重复索引；错误指出冲突 selector。
  2. Workbook 打开后一次性解析全部 selector，建立 request-selector -> physical index/name 表。
  3. 名称+索引或不同名称 comparer 结果命中同一 physical Sheet 时，抛稳定参数/配置异常，附 selector 与实际 Sheet 名/索引。
  4. 将已解析表传给 plan builder 和 import loop，避免重复 Resolve 及结果漂移。
  5. 对同一 plan group 的 Sheet 映射冲突给出领域异常，不让 LINQ `.Single` 泄漏。
- **测试矩阵**：重复名称、大小写 comparer、重复索引、名称+索引同物理 Sheet、越界索引、缺失名称、合法混合 selector、多 Sheet 相同模型/不同 mapping；XLS/XLSX 主链。
- **Mock 边界**：使用真实内存 Workbook，不 mock NPOI Sheet。
- **风险**：异常类型变化属于可观察行为；记录到 API/迁移文档。
- **验收**：所有冲突在计划/解析前确定性失败；合法混合请求保持通过。

### RH27-102 请求级 Workbook metadata 与静态状态移除

- **优先级**：P0；**依赖**：RH27-000；**breaking**：是。
- **目标**：Author/Company/Title/Subject/Category/Description 每次请求快照化，不共享可变对象。
- **已确认文件**：`ExcelSetting.cs`、`ExcelExport.cs`/request、`ExcelHelper.cs`、`NpoiExcelExporter.cs`、API snapshot、Docs consumer。
- **步骤**：
  1. 全仓搜索生产与消费者，定义 provider-neutral `ExcelWorkbookMetadataOptions`（或等价不可变快照）。
  2. 在 Workbook export request/builder 增加显式 metadata；Build 时深拷贝/冻结。
  3. NPOI 只读取 request metadata，不回退到静态可变对象。
  4. 移除 `ExcelSetting.Default` setter 和生产读取点；若保留一个 major 的 shim，只允许 `[Obsolete]` 创建独立快照，禁止继续暴露共享 mutable instance。
  5. 明确空值、默认值、XLS/XLSX 支持差异和模板 metadata 覆盖策略。
- **测试矩阵**：默认 metadata、显式字段、模板 preserve/override、顺序调用隔离、并发 64 个不同 metadata、Build 后修改源 options 不污染 request、包消费者迁移。
- **风险**：API 和默认作者值变化；必须给 before/after 和 member API diff。
- **验收**：生产路径无静态 mutable metadata；并发与顺序导出重开后值完全隔离。

### RH27-103 普通导出 Stream/File 失败合同与原子写入

- **优先级**：P0；**依赖**：RH27-000。
- **目标**：File API 不截断既有目标；Stream API 诚实描述部分写入边界。
- **已确认文件**：`NpoiExcelExporter.cs`、`ExcelStreamExtensions.cs`、`CsvStreamExtensions.cs`、`NpoiStreamCopier.cs`、Stream tests、Integration tests。
- **设计决策**：
  - File API 使用目标同目录随机临时文件，flush/close 成功后同卷 replace/move；既有目标替换策略兼容 Windows。
  - Stream API 默认直接写入且不承诺回滚；不得用无条件 `MemoryStream.ToArray()` 获得伪原子性。可选 staging 仅在有明确容量上限和真实需求时引入。
  - 调用方始终拥有 destination；失败后 position/length 合同按 seekable/non-seekable 分开。
- **步骤**：
  1. 在文件扩展边界验证路径、父目录和目标类型，不追随不受控临时目录。
  2. 实现同目录 staging、成功提交、失败清理、目标已存在/不存在两条路径。
  3. 转换器、序列化、取消、磁盘/目标写失败均不得覆盖旧文件。
  4. 为直接 Stream 写入记录可能部分写入；测试不可写、不可寻址、可寻址、中途失败和取消。
  5. CSV File API 同步采用一致原子合同；避免 Excel/Csv 行为分叉。
- **安全**：随机文件名、`CreateNew`、不记录敏感内容、不接受目录穿越假设；库不替调用方授权任意路径。
- **验收**：File 失败保留旧目标或不创建新目标；成功结果可重开；Stream 失败状态与文档/测试一致。

### RH27-104 Failure Workbook 临时文件与诊断治理

- **优先级**：P0；**依赖**：RH27-103。
- **目标**：临时目录、权限、命名、生命周期和清理失败均可诊断，取消也清理。
- **已确认文件**：`NpoiFailureWorkbookWriter.cs`、failure options/error model、P0 tests。
- **步骤**：
  1. 抽取可测试的临时文件/输出提交边界，不把 IO abstraction 扩散到业务层。
  2. 定义默认临时目录和可配置请求级目录；使用随机 `CreateNew`，限制文件共享。
  3. 无权限、磁盘满、序列化失败、目标复制失败、取消、删除失败分别产生稳定错误分类和结构化诊断。
  4. 主异常优先；清理失败附加为诊断/聚合信息，不静默吞掉，也不覆盖主错误。
  5. 保持 MaxBytes 超限不污染 destination 的既有合同。
- **测试**：故障注入文件系统边界 + Windows Integration；敏感内容不进入日志/异常；残留文件检查。
- **风险**：真实磁盘满难确定性模拟；Unit 用受控 boundary，Integration 用失败流/权限目录，禁止填满系统盘。
- **验收**：每条故障路径有自动证据；无未诊断残留；取消路径清理通过。

### RH27-105 Excel/CSV 资源限制真实语义

- **优先级**：P0；**依赖**：RH27-000。
- **目标**：文档与执行顺序一致，覆盖可控边界，不夸大对 NPOI DOM/压缩炸弹的保护。
- **已确认文件**：`ExcelImportPolicies.cs`、`NpoiExcelImporter.cs`、`NpoiStreamCopier.cs`、图片/Unique runtime、Csv options/pipeline、resource tests。
- **步骤**：
  1. 分层记录压缩输入字节、ZIP/OLE 容器、NPOI DOM、业务行/列/错误/图片/Unique、失败输出大小。
  2. XLSX 在 WorkbookFactory 前评估 NPOI/ZIP 可配置防护能力；XLS 单独定义 OLE 边界。无法可靠前置限制的项明确为部署隔离建议。
  3. 补极宽表、shared strings、styles、drawings、图片数量/总字节、压缩比等独立进程测试。
  4. 明确无图片映射时是否扫描图片：默认不扫描则 MaxPictures 不保护未映射 drawing，文档直述；如安全策略要求全局扫描，评估额外成本并配置化。
  5. CSV 增加输入字节、最大行、最大错误、最大字段长度和可选最大列数；在 parser/reader 边界生效。
  6. 所有上限错误使用稳定 ResourceLimit 分类，并有 truncation 元数据。
- **验收**：独立进程超限结果可复现；无“阻止所有解析峰值/压缩炸弹”的不实声明；默认上限与兼容策略有 ADR。

### RH27-106 配置迁移、异常分类与资源释放

- **优先级**：P0；**依赖**：RH27-000。
- **目标**：v1 非目标方向、反射异常、CSV 枚举器和错误分类端到端一致。
- **步骤**：
  1. 决定 v1 迁移的非目标方向为 `null` 还是显式空配置；与“缺失方向默认失败，ConventionFallback 才回退”一致。
  2. 建立损坏文件、配置、结构、行、关系、资源、取消、输出错误的类型/错误码矩阵。
  3. 审计 `NpoiRelationBinder` 和所有 reflection invoke，统一解包 `TargetInvocationException` 并保留 stack/原异常。
  4. `CsvRecordReader.Read(...).GetEnumerator()` 使用明确 `using/finally`；覆盖 MoveNext、转换、setter、validation、取消异常。
  5. 禁止 catch-all 静默 fallback；历史 exceptions 是否删除交 Phase 2。
- **验收**：每类错误有 provider 主链测试；enumerator/parser/writer 在成功、失败、取消均 Dispose。

### RH27-107 异步、取消与公式缓存 ADR

- **优先级**：P1；**依赖**：RH27-103、RH27-105。
- **目标**：只实现有真实 IO 收益的 async，明确 NPOI 不可取消阶段与公式语义。
- **步骤**：
  1. 审计所有循环、Stream/File copy 和序列化检查点；传递 token 到可取消边界。
  2. 测量 WorkbookFactory/NPOI Write 阶段取消延迟；文档说明无法中断的位置。
  3. 公式只读取缓存值，不执行重算；覆盖缺失缓存、陈旧缓存、错误缓存、类型不匹配和公式文本暴露。
  4. ADR 决定是否新增 async File/Stream API。禁止 `Task.Run`、`.Result`、`.Wait()` 和仅改签名的伪异步。
- **验收**：ADR、测试和文档一致；若延期 async，给出明确原因和未来触发条件。

### Phase 1 门禁

- RH27-101～106 的 P0 测试全部通过，无跳过；前序回归合同全绿。
- `decisions.md` 完成输出、资源、迁移、metadata、async/公式决策。
- 无新增 TODO、NotImplemented fallback、吞异常或静态可变状态。

## 5. Phase 2：API 收敛与 Breaking Change

### RH27-201 Public API 分类账与消费者取证

- **优先级**：P0；**依赖**：Phase 1 合同稳定。
- **目标**：对每个 public/protected/internal 类型与成员标记用户 API、Provider SPI、执行细节、兼容层、未使用、待验证。
- **步骤**：
  1. 从反射快照生成机器可读账本，关联源码、仓库调用、Docs、PackageConsumer 和替代入口。
  2. 将用户 API 与 Provider SPI 分离；SPI 使用 `EditorBrowsable(Never)`，但不靠生产 IVT 穿透程序集。
  3. 审计约 180 个公开顶层类型及成员，不能仅按类型名删除。
  4. 检查 NPOI 类型泄漏、mutable collection、implementation concrete type 和 friend assembly。
- **验收**：每个 public 类型有场景或明确迁移/删除结论；`api-governance.md` 可追溯。

### RH27-202 唯一推荐入口与术语统一

- **优先级**：P0；**依赖**：RH27-201；**breaking**：是。
- **目标**：形成一条导入/导出主路径和方向明确的 Mapping API。
- **计划变更**：
  - `ExcelMapping.For<T>()` -> `ExcelMapping.Import<T>()` / `Export<T>()`。
  - `Mapping(configuration/document)` -> 方向化且具名的配置/文档入口，避免仅靠重载类型区分。
  - `HasTitle`/`HasHeader` 统一为用户文档中的唯一术语。
  - `HeaderMatch` -> `RequireExpectedHeaders`；`MaxColumnCount` -> `MaxReadColumns`。
  - `EnabledEmptyLine` -> `ReportEmptyRows`；`IgnoreEmptyLineAfterData` -> `StopAtFirstEmptyRow`。
  - `AddNavigationSheet` -> `AddSheet(name, parents.SelectMany(...))`。
  - NPOI `AddNpoi()` 返回 `IServiceCollection`，符合 DI 链式约定。
- **兼容策略**：next-major 删除无价值同义入口；确有迁移价值者只保留一个明确周期的 `[Obsolete]` forwarding，禁止双实现。
- **验收**：README 只展示唯一主路径；API diff 和 PackageConsumer 同时验证新入口及批准的 shim。

### RH27-203 删除、降级和 SPI 最小化

- **优先级**：P1；**依赖**：RH27-201。
- **候选删除/降级**：`SheetSetting`、`RegexConst`、无消费者的 `TypeExtensions`/`ExpressionExtension`、legacy Office exceptions、`ExcelValueMap<T>`、`ICellValueConverter`、重复旧属性族、旧 CsvHelper overload/global state、无生产调用 NPOI helper。
- **候选 internal/private**：`ExcelTypeMapFactory`、`ExcelTypeMap<T>`、`ExcelPropertyMap`、binding resolver、`ExcelMappingPlanFactoryProvider`、默认 loader、merger/document factory、内置 validation concrete classes、`UniqueTracker`（若非 Provider SPI）。
- **步骤**：每项执行仓库/测试/示例/nupkg consumer 搜索；判断第三方 Provider 需求；只公开最小稳定 abstraction。
- **验收**：不存在仅因测试方便而 public 的实现；无生产程序集 IVT；删除项有 negative API baseline。

### RH27-204 API 工具化与迁移包

- **优先级**：P1；**依赖**：RH27-202/203。
- **目标**：将手写快照升级为可维护的 API 门禁。
- **步骤**：评估 PublicApiAnalyzers/APICompat 与现有 hash snapshot 的组合；生成 before/after 顶层类型和成员 diff；迁移表包含旧入口、新入口、版本、示例和行为差异。
- **验收**：CI 可检测意外 public API；NPOI 不泄漏；实际 nupkg consumer 通过。

### Phase 2 门禁

- 每个 public 类型都有用户或 Provider 场景；生产 IVT 为零。
- 所有 breaking change 有 API diff、迁移文档和 PackageConsumer。
- Phase 1 行为测试不变。

## 6. Phase 3：目录、职责与实现重构

### RH27-301 Import Builder/Request 与 importer 编排拆分

- **优先级**：P1；**依赖**：Phase 2 API 冻结。
- **目标**：一个主要 public 类型一个文件；编排方法直接显示 resolve -> plan -> validate -> materialize -> relation -> failure artifact。
- **已确认范围**：`ExcelImport.cs` 当前同时含入口、Workbook Builder、Sheet Builder；importer 仍含 selector、column binding、reflection dispatch、cell helpers 和 runtime 类型。
- **步骤**：按 `Importing/{Builders,Requests,Results,Options}` 拆分；提取 physical sheet resolver、column binder、resource runtime/source location；缓存 ItemType generic dispatch delegate；保持 internal。
- **验收**：无额外 public surface；新协作对象有直接测试；主编排无反复解析 selector。

### RH27-302 Export/Failure Workbook/Csv 职责拆分

- **优先级**：P1；**依赖**：RH27-103/104、Phase 2。
- **步骤**：
  1. Exporter 保留 plan、workbook lifecycle、commit 编排；Cell/Chart/Template/IO 分域。
  2. Failure Workbook 拆为 Row/StyleRichText/Drawing copier 和 Output Committer，复用单一临时提交合同。
  3. Csv importer/exporter/record reader/writer 分文件，共享单次操作生命周期，不创建多层空抽象。
- **验收**：每个 collaborator 有明确替换/测试/程序集边界价值；golden file 不变。

### RH27-303 Mapping contract 与 internal plan 分离

- **优先级**：P1；**依赖**：RH27-203。
- **目标**：Provider contract 只暴露只读所需字段，具体 plan/column/layout/cache/merger internal 化。
- **步骤**：校验 clone/merge/cache key 对所有字段完整；缓存反射 accessor delegate；没有 converter 时不创建 ConversionContext；命名规则/转换器预索引。
- **验收**：字段“输入 -> merge -> plan -> provider”追溯测试完整；cache hit/miss/eviction/concurrency 直接测试。

### RH27-304 目录、命名和死代码治理

- **优先级**：P2；**依赖**：RH27-301～303。
- **步骤**：namespace 与目录一致；修正 `ConditionalFormattin` 等拼写；删除无引用 helper/宽泛 fallback；不机械拆分微型 DTO，不全仓格式化。
- **验收**：public API 只发生 Phase 2 已批准变化；diff 可 review；编译与 golden regression 不变。

## 7. Phase 4：测试体系闭环

### RH27-401 P0 正确性与故障矩阵

- **优先级**：P0；**依赖**：Phase 1～3。
- **必须覆盖**：重复 selector；metadata 并发；损坏/空/随机/截断/加密/扩展名不符；导出转换/取消/序列化/中途写失败；临时目录权限/删除/目标提交；资源放大；Dispose；实际 nupkg consumer。
- **测试规则**：xUnit 英文 `Method_State_Expected`；每个测试中文 XML 目的和 AAA；真实 NPOI/Stream/DI，mock 仅限外部边界。
- **验收**：P0 无跳过，生产符号追溯矩阵完整。

### RH27-402 P1/P2 行为矩阵

- **P1**：Data Validation 命名区域/跨 Sheet/公式列表/Custom/日期系统/地址；Relation 异常与集合边界；CSV BOM/编码/换行/公式注入/超长字段/并发 options；Mapping v1/clear/profile lifecycle；模板保护/隐藏/链接；图表空/单行/非数值；Cancellation/Dispose。
- **P2**：继承/隐藏/init-only/只读属性；nullable enum/DateTimeOffset/Guid/Version/泛型；零列/最大 Sheet 名/Unicode/RTL/代理项；XML docs 和 SPI compatibility。
- **验收**：未完成 P1 有明确 release waiver；P2 进入 backlog 不伪装完成。

### RH27-403 Integration、golden 与独立进程

- **优先级**：P0。
- **范围**：真实磁盘 XLS/XLSX、多 Sheet、模板、图片、图表、失败工作簿重开；原子替换；权限/失败路径；大文件资源采集。
- **办公客户端**：使用 Bing.Offices 生产链生成、GenerationId + manifest/source/roundtrip SHA 绑定的 fixtures；Excel/WPS/LibreOffice 自动或人工打开、保存、重开并 NPOI 二次解析。缺失客户端保持 BLOCKED，不复用旧代际结果。
- **验收**：自动化与外部 blocker 分离；任何未完成客户端不得包装为 PASS。

### RH27-404 全 TFM 与 PackageConsumer

- **优先级**：P0。
- **步骤**：所有 TFM restore/build；已安装 runtime 执行 test；net6/net8 Unit/Integration 必须全绿；Docs 从本地 nupkg 恢复；另建临时 artifacts consumer 验证公开主路径、SPI（如承诺）和 XML docs。
- **验收**：包消费不依赖项目引用或源码 friend；旧 runtime 缺失按环境阻塞记录。

## 8. Phase 5：性能、GC 与 Benchmark

### RH27-501 Benchmark workload 校准

- **优先级**：P0；**依赖**：正确性冻结。
- **当前状态**：源码已拆出真实 FailureRowCount、HeaderAttribute、ValidationRowCount、Regex、Unique、cold/hit/miss/eviction 场景，但需要重新审查参数确实进入 workload 并生成当前版本产物。
- **步骤**：
  1. 使用 BenchmarkDotNet 默认自适应或经论证配置，避免 1/2/3 仅适合烟测的结论夸大。
  2. 分离 plan factory 创建、rule 构造、cold/hit/miss/eviction；MultiRule 不得预热后只测 hit。
  3. Failure/style/validation/regex/unique 各自只携带实际消费参数。
  4. 删除 PeakWorkingSet 微基准；峰值工作集用独立进程采样。
- **验收**：Reviewer 能从源码证明每个参数进入 workload；原始 JSON/Markdown 与 commit/diff hash 绑定。

### RH27-502 热路径优化及前后对照

- **优先级**：P1；**依赖**：RH27-501。
- **候选**：compiled typed accessor；无 converter 不创建 context；Dictionary/FrozenDictionary（仅兼容 TFM 可用时）；ValidationRangeIndex 低分配查询；ArrayPool buffer finally 归还；mapping cache key；ExportToBytes 重复复制；AutoSize/RichText/Style/Picture 成本；有界 dynamic Type cache。
- **原则**：每个优化先有 profiler/benchmark 证据；同一环境 before/after；无业务旁路；不默认引入 MemoryPool/对象池/stackalloc/ValueTask/struct/Source Generator。
- **验收**：正确性完全一致；收益、回归和复杂度取舍写入 `benchmark-report.md`。

### RH27-503 规模、并发与资源预算

- **场景**：Import/Export 1K/10K/100K/宽表；CSV 1M/大字段；Mapping 10/100/1K/10K rules；Validation 0/1/重叠/多列；模板/图片/图表/AutoFit/Failure；seekable/non-seekable；并发 1/4/16/64。
- **指标**：Mean/Error/StdDev、Allocated/op、Gen0/1/2、可验证 LOH、独立进程 Peak Working Set、吞吐和尾延迟、destination capacity；环境/GC 模式完整。
- **资源探针**：对 plan、UniqueTracker/values 分别定义存活/释放阶段并使用 `GC.KeepAlive`；禁止人工 LOH payload 代替生产 workload。
- **验收**：100K 行正式预算与结果；无“Low GC/Near-Zero/0 GC/streaming”不实结论。

## 9. Phase 6：文档、注释与发布准备

### RH27-601 XML Documentation 与源码表述

- **优先级**：P1；**依赖**：最终 API。
- **步骤**：按中文注释 Skill 补当前改动 public API 的 summary/typeparam/param/returns/exception；实现/override 用 `inheritdoc`；说明 Stream ownership、线程安全、失败状态、取消延迟和资源边界；删除 importer/exporter“单遍流式”等不实表述。
- **验收**：XML docs 无计划外缺失/格式警告；注释与真实行为一致。

### RH27-602 用户文档与可执行示例

- **优先级**：P1。
- **主题**：快速开始/DI；主导入导出路径；四类 Profile；Attribute/Profile/Document/Request 优先级；v1/v2 与 breaking migration；动态列/多 Sheet/关系/图片；错误/失败工作簿；模板/样式/批注/图表；Stream/File；并发/取消/容量；性能复现；Word/PDF 不在本次范围。
- **步骤**：复用并扩展 `docs/excel`，避免重复页面；从 Markdown 原文提取 C# fence 编译运行；关键示例从实际 nupkg 执行。
- **验收**：README 只展示唯一推荐路径；无失效链接、占位版本或过度承诺。

### RH27-603 Pack、SBOM 与发布清单

- **优先级**：P0。
- **步骤**：验证 nupkg 依赖、XML docs、snupkg、license、readme、repository metadata；生成 API diff、migration、release notes、SBOM/依赖许可/安全说明；列出支持 TFM/runtime/OS/NPOI/CsvHelper。
- **验收**：PackageConsumer 使用本轮本地包；只生成本地产物，不上传。

### RH27-604 最终独立 Review 与 Go/No-Go

- **优先级**：P0；**依赖**：全部任务。
- **Review 项**：真实调用链、参数/异常/取消/Dispose/并发、API 重复、internal 直接测试、生产 IVT、伪 async、隐藏分配/无界缓存/静态状态、Docs/API、一致 Benchmark workload。
- **门禁**：BLOCKER/P0=0；P1 修复或有批准 waiver；build/test/pack/consumer/benchmark 有真实证据；办公客户端 blocker 按发布政策明确处理。
- **输出**：`final-report.md` 给出真实 Go/No-Go、完成 Phase、剩余风险、breaking migration、建议提交分组；不得自动 commit/push/tag/publish。

## 10. Breaking Change 清单

| 变更 | 目标 | 迁移策略 | 必须证据 |
| --- | --- | --- | --- |
| 移除 `ExcelSetting.Default` mutable global | request metadata snapshot | 新 options 示例；必要时一个周期 obsolete shim，但不共享对象 | 并发隔离、API diff、consumer |
| `ExcelMapping.For<T>()` 方向化 | `Import<T>()`/`Export<T>()` | next-major 删除或单周期 forwarding | 双方向示例、negative baseline |
| 两个 `Mapping(...)` 收敛 | 具名 Document/Configuration | 机械迁移表 | Docs fences、consumer |
| 布尔语义方法重命名 | 自说明方法 | next-major rename，必要时 obsolete forwarding | 默认值与行为对照 |
| `AddNavigationSheet` 删除 | `AddSheet + SelectMany` | 文档迁移 | 相同输出 golden |
| `ICellValueConverter` 删除 | `IExcelValueConverter` | adapter 仅过渡期存在 | provider-neutral context tests |
| 执行细节 public -> internal | 最小 Provider SPI | SPI consumer 或明确无消费者 | APICompat/negative snapshot |
| legacy attributes/exceptions/helpers 删除 | Excel 前缀属性和结构化错误 | old/new 对照 | package consumer 编译 |
| `AddNpoi` 返回值调整 | `IServiceCollection` | source-compatible 时直接改，否则 major | DI chaining test |

## 11. 验证矩阵

| 变更域 | Unit | Integration | PackageConsumer | Docs | Benchmark |
| --- | --- | --- | --- | --- | --- |
| Sheet selector | 必须 | XLS/XLSX | API 编译 | 错误合同 | 不适用 |
| Metadata | 必须并发 | 重开 metadata | 必须 | 必须 | 并发/分配可选 |
| Stream/File 原子性 | 必须故障流 | 必须真实磁盘 | 必须 | 必须 | seek/non-seek/capacity |
| 临时文件 | 故障注入 | Windows 权限/清理 | 不适用 | 安全策略 | failure workload |
| 资源限制 | 边界 | 独立进程 | Options 编译 | 必须 | 1K～100K/峰值 |
| 配置/异常/CSV | 必须 | 主链 | 必须 | 必须 | CSV/Mapping |
| API 收敛 | snapshot/negative | DI | 必须 | fences | 不适用 |
| 内部重构 | 直接 internal + facade | golden | 必须 | 无行为变化 | 前后对照 |
| 图表/模板/图片 | 必须 | 重开/客户端 | 示例 | 必须 | 专项 |
| 发布包 | API contract | 本地包 | 必须 | 必须 | 结果归档 |

## 12. 真实命令基线

执行 Agent 先确认 restore 成功，再按仓库真实项目运行：

```powershell
dotnet --info
dotnet restore Bing.Offices.sln
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore
dotnet pack src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release --no-build --no-restore
dotnet pack src/Bing.Offices.Core/Bing.Offices.Core.csproj -c Release --no-build --no-restore
dotnet pack src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-build --no-restore
dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks -- --filter "*TargetBenchmark*" --exporters json markdown
git diff --check
```

- Benchmark filter 必须替换为实际分组，不一次盲跑全部超大矩阵。
- Docs Tests 的本地包源/版本需与本轮 pack 输出对齐，禁止误消费缓存中的旧 2.0.0 包。
- Flubu 脚本包含 `nuget.publish` 目标且默认配置会注册发布链，除非明确选择不触发 publish 的安全 target，否则优先使用上述显式 dotnet 命令。

## 13. 完成定义

只有以下条件全部成立，任务才可标记 `COMPLETE`：

1. Phase 1～6 所有可执行任务达到 VERIFIED；BLOCKER/P0 为零。
2. P1 已修复或具名、带风险和批准人的 release waiver；P2 有明确后续归属。
3. 重复物理 Sheet、metadata 隔离、文件原子性、Stream 失败、临时文件、资源/CSV/公式合同有真实测试。
4. API 分类、唯一入口、before/after diff、迁移和实际 nupkg consumer 完成；无生产 IVT。
5. net6/net8 Unit/Integration、Docs、Release build、pack 全绿；其他 TFM 至少构建，运行时缺失如实 BLOCKED。
6. Benchmark workload 参数真实、原始 artifacts 完整、100K 预算明确，不宣称未达到的 streaming/async/低 GC/0 GC。
7. Excel/LibreOffice/WPS 自动化或人工互操作结果与阻塞按发布政策真实记录，不混用旧代际产物。
8. README/docs/XML docs/release notes 与最终 API 和行为一致。
9. `final-report.md` 给出证据化 Go/No-Go。
10. 未自动执行 commit、push、tag、PR 或 NuGet 发布。

## 14. 建议提交分组（仅建议，不执行）

1. `fix(import): reject duplicate physical sheet selectors`
2. `feat(export): scope workbook metadata to requests`
3. `fix(io): make file exports atomic and diagnose temporary cleanup`
4. `feat(resources): enforce documented excel and csv limits`
5. `refactor(api): converge mapping and workbook entry points`
6. `refactor(core): separate provider contracts from execution details`
7. `test(release): complete failure resource package and interop matrices`
8. `perf(benchmarks): align workloads and publish resource evidence`
9. `docs(release): finalize migration limits and release guidance`

本计划明确不执行上述提交，也不执行 push、tag 或 NuGet 发布。
