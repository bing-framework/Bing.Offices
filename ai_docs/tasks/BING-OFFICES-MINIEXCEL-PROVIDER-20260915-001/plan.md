# Bing.Offices MiniExcel Provider 实施计划

## 1. 任务信息

- Task ID：`BING-OFFICES-MINIEXCEL-PROVIDER-20260915-001`
- 计划日期：2026-09-15
- 角色：Planner（本文件只规划，不实施）
- 目标：新增独立 `Bing.Offices.MiniExcel` Provider，在共同能力范围内复用现有 Workbook Request、Mapping、Converter、Validation、异常和文件提交契约，并用真实性能与资源证据验证其大数据价值。
- 计划状态：待人工批准
- 预计 Phase：17
- 预计 Task：26

## 2. 已读取证据

### 2.1 仓库规则与设计

- `AGENTS.md`
- `README.md`
- `.github/workflows/ci.yml`
- `common.props`、`framework.props`、`common.tests.props`、`Directory.Build.targets`
- `ai_docs/excel/00-overview.md` 至 `08-validation.md`、`implementation-progress.md`
- `docs/excel/README.md`
- 既有 `ai_docs/tasks/` 中 API、资源、消费者、Benchmark 和发布记录

### 2.2 已确认源码

- 公共契约：`IExcelExporter`、`IExcelImporter`、Workbook Export/Import Request 与 Result、`ExcelResourceLimits`、`IFileExportCommitter`、异常体系。
- 公共映射：`IExcelMappingPlanFactory`、`IExcelMappingWorkbookPlan`、`IExcelMappingPlan`、固定/动态列计划、Profile、Attribute、JSON/XML、Fluent、Converter、Validation、`UniqueTracker`。
- NPOI 实现：Exporter/Importer、Plan Builder、Sheet Writer/Executor、Row Materializer、Relation Binder、Failure Workbook、Async Staging、Stream Copier、XLSX ZIP Preflight。
- 工程门禁：Solution、双 TFM 测试、API snapshot、Pack、net6/net8 PackageReference Consumer、Benchmark/Resource Probe。

### 2.3 外部版本证据（实施时必须重新确认）

- NuGet 官方页面在计划日显示稳定版 `MiniExcel 1.46.0`，包含 `netstandard2.0` 与 `net8.0` 资产，可覆盖 Provider 的 `net6.0;net8.0` 目标：https://www.nuget.org/packages/MiniExcel
- 官方仓库将 2.0 标为 preview，并说明其 API/包结构正在变化；P0 不采用 preview：https://github.com/mini-software/MiniExcel
- 官方文档显示 v1 系列具有 `SaveAsAsync`/`QueryAsync`，且从 v1.25.0 起接收 `CancellationToken`；这不能替代本仓库对真实异步、延迟枚举和取消时机的源码/测试 Spike。

## 3. 当前基线与完成度

### 3.1 已真正实现

1. Abstractions/Core 已提供 Provider-neutral Workbook Request、结果、错误、Mapping Plan、Converter、Validation、Profile、动态列和原子文件提交接口。
2. `ExcelMappingPlanFactory` 在 Core 中完成 `Attribute < Profile < JSON/XML Document < Request Fluent` 合并，预绑定 converter/validation，并缓存不可变映射计划。
3. NPOI 已实现 XLS/XLSX、多 Sheet、动态列、模板、样式、多级表头、关系、图片、批注、图表、Failure Workbook、资源限制及同步/外围异步 IO。
4. `ExcelStreamExtensions` 已为任意 `IExcelExporter`/`IExcelImporter` 提供 byte[] 与文件读取便利入口；`DefaultFileExportCommitter` 已提供原子提交。
5. CI 已执行 Release build、双 TFM Unit/Integration、Docs、资源矩阵、Benchmark 发现、三个包的 pack/内容检查和 net6/net8 包消费者。

### 3.2 未实现或部分实现

- MiniExcel Provider：**0%**。仓库没有 MiniExcel 项目、包引用、实现、测试、文档或发布配置。
- 第二 Provider 共享基础：**大部分已具备，但尚未抽取执行级公共逻辑**。`NpoiRelationBinder`、两个 Plan Builder、错误收集、部分行物化、Async Staging、Stream Copier 与 ZIP Preflight 被锁在 NPOI 程序集内部。
- Provider-neutral 文档：部分完成。`IExcelExporter.ExportAsync` 与 `IExcelImporter.ImportAsync` XML 注释仍写死 NPOI DOM。
- Provider 并存：未定义完整契约。NPOI 使用 `TryAddTransient`，同一容器注册两个 Provider 时“先注册者生效”，需要固化并文档化。
- MiniExcel 能力、异常、资源、性能：全部待真实 API/文件/Benchmark 验证，不能用第三方宣传替代证据。

### 3.3 不是完成证据的内容

- 现有公共接口不代表 MiniExcel 已适配。
- MiniExcel 官方支持某项功能，不代表能兑现 Bing.Offices 的完整语义。
- `QueryAsync` 返回 `Task<IEnumerable<T>>` 不自动等于逐行异步流式消费。
- `IEnumerable<IDictionary<string, object>>` 适配存在不代表没有每行 Dictionary 分配或提前物化。
- 能生成可打开的 XLSX 不代表列顺序、日期、decimal、动态列、错误上下文、取消与流所有权满足契约。

## 4. 计划前关键判断

1. **公共抽象是否足够**：足够承载第二 Provider 的 P0；无需新增 MiniExcel 专属 Request/Attribute/Profile。
2. **NPOI 泄漏**：公共签名未泄漏 NPOI 类型；两处接口 XML 注释泄漏 NPOI DOM，需改为 Provider-neutral 描述。
3. **应下沉的逻辑**：关系绑定、映射计划分组/运行时泛型调度、XLSX ZIP 预检明确不依赖 NPOI；错误收集、转换/校验行物化为 mixed，先提取最小 Provider-neutral kernel；Sheet/Cell、图片、模板和 Failure Workbook 保持 Provider-specific。
4. **最小 P0**：XLSX 普通/多 Sheet、固定与动态列、Stream/File/byte[]、现有 Mapping/Converter/Validation、Sheet name/index、header/body 策略、关系、资源限制、异常、真正可验证的 async/cancellation。
5. **不能在 P0 承诺**：XLS、Chart、Failure Workbook、图片导入/导出、复杂 Comment/Style/Header/Merge、模板完整保真。
6. **Public Capability API**：P0 不需要。Provider 内部 preflight 足以 fail-fast；能力矩阵放文档和测试，不扩张公共 API。
7. **Runtime Provider Resolver**：P0 不需要。应用启动时选一个 Provider；交叉测试使用直接实例或独立容器。
8. **Async Pipeline**：不整体重构。先 Spike MiniExcel v1.46.0 的同步/异步边界；只有确需复用时才把 staging/copy 最小下沉 Core，禁止 `Task.Run`、`.Result`、`.Wait()`。
9. **XLSX ZIP Preflight**：应下沉 Core。当前实现仅依赖 ZIP/XML/公共资源限制，唯一 NPOI 依赖是命名和异常 provider 值。
10. **Breaking Change**：预计无公共成员删除或签名变更；新增 MiniExcel 包及 public Provider/DI 扩展是 additive。XML 注释中性化不构成二进制 Breaking Change。
11. **TFM**：Provider 延续 `net6.0;net8.0`。稳定 MiniExcel 1.46.0 的资产表面兼容，但必须用 restore/build/consumer 实证。
12. **测试/Benchmark 缺口**：MiniExcel 的职责级、真实 XLSX、交叉契约、异步取消、资源、包消费者和 1K~1M Benchmark 全部缺失。
13. **可发布剩余工作**：完成以下 17 Phase、报告、API snapshot/审批、包内容验证、NPOI 回归和发布前独立 review；本任务不发布 NuGet。

## 5. 目标调用链

```text
ExcelWorkbookExportRequest
  -> Provider 内部 capability/preflight
  -> Core IExcelMappingPlanFactory.CreateWorkbook<T>(Export)
  -> MiniExcel 延迟行适配器（固定列 + 动态列 + converter/value map）
  -> MiniExcel SaveAs/SaveAsAsync XLSX
  -> 调用方 Stream 或 IFileExportCommitter 原子提交
```

```text
MiniExcel Query/GetReader 原始行
  -> Provider 原始值适配为 ExcelCellValue
  -> Core Mapping Plan getter/setter/converter/value map
  -> Core Validation binding + UniqueTracker
  -> Entity + 结构化 ExcelImportError
  -> Core relation binder
  -> ExcelWorkbookImportResult
```

禁止形成 `MiniExcel Attribute -> MiniExcel 直接映射实体` 的旁路。

## 6. 功能与能力矩阵（计划目标，不是完成声明）

| Feature | NPOI 当前 | MiniExcel P0 目标 | 分级/处理 |
|---|---:|---:|---|
| XLSX | Yes | Yes | P0 |
| XLS | Yes | No | Unsupported，Plan 阶段 fail-fast |
| Stream/File/byte[] | Yes | Yes | P0 |
| Multi Sheet / 异构 DTO | Yes | Yes | P0 |
| Dynamic Column | Yes | Yes | P0，延迟枚举 |
| Attribute/Profile/Fluent/JSON/XML | Yes | Yes | P0，共用 Core plan |
| Value Mapping/Converter | Yes | Yes | P0 |
| Validation/Unique/动态校验 | Yes | Yes | P0 |
| Relations | Yes | Yes | P0，共用 Core binder |
| Template | Yes | 待 Spike | P1；语义不完整则 Unsupported |
| 基本 Header/NumberFormat | Yes | 待 Spike | P0/P1，以实测定界 |
| 复杂 Style/Merge/Multi Header | Yes | 待 Spike | P1/P2 或 Unsupported |
| Image Export/Import | Yes | 待 Spike | P2 或 Unsupported |
| Comment | Yes | 待 Spike | P2 或 Unsupported |
| Chart | Yes(XLSX) | No | Unsupported |
| Failure Workbook | Yes | No | P2/Unsupported；P0 返回错误集合 |
| Streaming/Large Dataset | Partial/DOM | Yes（待证） | P0 + 性能门禁 |

## 7. 文件范围

### 7.1 已确认将修改

- `Bing.Offices.sln`
- `version.dev.props`
- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/IExcelExporter.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/IExcelImporter.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/`（仅最小共享执行组件）
- `src/Bing.Offices.Core/Bing/Offices/Imports/`（relation/error/materialization kernel）
- `src/Bing.Offices.Core/Bing/Offices/IO/`（仅经 Spike 证明需要的共享 staging/copy）
- `build/ApiSnapshot/Program.cs`、`build/api-snapshot-baseline.json`
- `.github/workflows/ci.yml`
- `benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj`、`Program.cs`
- `README.md`、`docs/excel/README.md`、`ai_docs/excel/*.md`

### 7.2 已确认将新增

- `src/Bing.Offices.MiniExcel/Bing.Offices.MiniExcel.csproj`
- `src/Bing.Offices.MiniExcel/AssemblyInfo.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Exports/*`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/*`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Extensions/MiniExcelServiceCollectionExtensions.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/*`
- MiniExcel 职责级 Unit/Integration/Cross-provider 测试文件
- MiniExcel Benchmark、controlled enumerable/stream、真实 XLSX fixture
- `ai_docs/excel/09-providers.md`
- 本任务的 execution/review/matrix/report/final 文档

### 7.3 候选文件（Phase 0/1 后定稿）

- 新建独立 `tests/Bing.Offices.MiniExcel.Tests`、`tests/Bing.Offices.MiniExcel.Tests.Integration`、`tests/Bing.Offices.ProviderContract.Tests`，或在现有 Unit/Integration 项目中以清晰目录隔离。选择标准是双 Provider 依赖隔离、测试速度和 API snapshot 复用，不为项目数量本身优化。
- 新建 `Consumer.Net6.MiniExcel`/`Consumer.Net8.MiniExcel`，或把现有消费者参数化为 Provider 矩阵。必须仍通过真实 `.nupkg` 的 `PackageReference`，不能改成 ProjectReference。
- `NpoiWorkbookPlanKeyBuilder`、两个 Plan Builder、`NpoiRelationBinder`、`ExcelImportErrorCollector`、部分 `NpoiImportRowMaterializer`、`NpoiAsyncStaging`、`NpoiStreamCopier` 的移动/拆分文件名由 Phase 1 的依赖分类决定。

## 8. 分阶段执行任务

### 8.0 Task 字段继承规则

下列规则是本节 26 个 Task 的组成部分，用于避免在每个 Task 中机械重复；Task 自身字段覆盖本规则：

- **当前状态**：等同该 Task 的“现状/证据”；未实现项均为 Missing，抽取项均为 Existing in NPOI/Not reusable，Spike 项均为 Needs Investigation。
- **源码证据**：等同“现状/证据”与第 2、3、7 节列出的已确认文件；候选文件不能在 Phase 0/1 结论前当作确定修改项。
- **涉及/新增/修改文件**：等同该 Task 的“修改范围”并受第 7 节文件清单约束；不得顺手修改清单外业务文件。
- **调用链**：以第 5 节为总体链路，各 Task 的“实施步骤”是该链路对应节点；任何偏离必须写入 execution/decision 记录。
- **公共 API 影响**：除 ME-1000、ME-1001 的 additive Provider/DI API 外均为 None；ME-100 只改 XML 文档。
- **Breaking Change**：所有 Task 默认 None；如实施证据要求删改公共成员、改变 NPOI/DI/错误行为，必须停止并取得成员级审批，不能用本计划自动授权。
- **性能与 GC/Resource**：热路径 Task（ME-301/401/501/600/700/900/1400/1401）必须给出 allocation、缓存、枚举、临时文件或集合生命周期证据；其余 Task 至少确认无新增热路径。
- **异常路径**：生产 Task 必须沿用现有 BingOffices exception/observer contract，保留 provider/operation/stage/cause；取消继续抛 `OperationCanceledException`。
- **并发/异步影响**：除 ME-700/701/1401 外默认不引入共享可变状态或新异步边界；发现影响时必须补直接测试。
- **测试**：每个生产修改 Task 必须增加职责级直接测试；文档/基线 Task 用内容、命令与 diff 证据验证。
- **验证命令**：使用第 9 节与该 Task“验证”指定的 filter/项目；Phase 0 确认新项目路径后将精确命令写入 execution 和 CI。

### Phase 0：Baseline / 真实 MiniExcel API Spike

#### Task ME-000：冻结基线与需求矩阵

- **目标**：记录 clean/dirty 状态、当前构建测试、包与 API 基线，建立 `requirements-matrix.md`。
- **现状/证据**：计划时工作树 clean；Solution 无 MiniExcel；CI 硬编码 3 个包。
- **修改范围**：仅任务文档；不得改生产代码。
- **实施步骤**：保存 commit/SDK/OS；执行 restore/build/双 TFM tests/API snapshot/pack 基线；把每项需求标为 Existing/Partial/Missing/Investigate。
- **依赖**：可访问现有 NuGet 源。
- **验证**：使用第 9 节基线命令；记录完整退出码和失败原因。
- **风险**：历史文档存在已过时的 TFM 描述，以 csproj 与本次命令为准。
- **验收标准**：基线可复现，未把历史报告当作当前结果。

#### Task ME-001：MiniExcel 1.46.0 API/行为 Spike

- **目标**：在隔离实验中确认稳定版的导入、导出、多 Sheet、动态 schema、async、cancel、流所有权、模板及异常行为。
- **现状/证据**：官方元数据显示 1.46.0 稳定且 TFM 表面兼容；2.0 仍 preview。
- **修改范围**：候选临时 Spike 测试或正式 Provider 测试 fixture；不把 preview 放入生产依赖。
- **实施步骤**：核对包依赖树/许可证；编译真实签名；用 controlled enumerable/stream 验证枚举时机、`QueryAsync` 返回后的读取时机、CancellationToken 和 Dispose；生成/读取多 Sheet XLSX。
- **依赖**：ME-000。
- **验证**：最小可执行测试 + `.nupkg` 资产检查；结论写入 `provider-capability-matrix.md`。
- **风险**：官方文档和当前包签名可能不一致；v1 async 可能只是异步包装或返回延迟同步枚举。
- **验收标准**：锁定 `MiniExcelPackageVersion=1.46.0` 或记录拒绝原因/稳定备选；所有 P0 设计均有 API 证据。

### Phase 1：Provider-neutral Contract 审查与最小下沉

#### Task ME-100：中性化公共契约并分类 NPOI helper

- **目标**：清除 XML 注释泄漏，形成 Provider-specific/neutral/mixed 分类。
- **现状/证据**：两个公共 async 接口 remarks 写死 NPOI；Core/Abstractions 无 NPOI 类型引用。
- **修改范围**：两个接口注释、分类记录；签名不变。
- **实施步骤**：改为“Provider 的同步阶段不保证异步”；逐个记录 helper 的依赖和所有者；禁止批量搬迁。
- **依赖**：ME-000。
- **验证**：`rg` 确认 Abstractions/Core 无 NPOI DOM 描述或类型；双 TFM API diff 无成员变化。
- **风险**：无二进制风险；XML 文档语义必须仍准确描述 NPOI。
- **验收标准**：公共文档中性，分类可追溯到源码。

#### Task ME-101：下沉计划分组、关系绑定与 XLSX 安全预检

- **目标**：只提取被两 Provider 真实复用的执行能力。
- **现状/证据**：Plan Builder/Relation Binder/ZIP Preflight 核心不依赖 NPOI DOM，但异常 provider 名被硬编码。
- **修改范围**：Core 新增 internal helper；NPOI 改用共享实现；必要时用 internal context 参数传 provider。
- **实施步骤**：先搬测试再搬代码；保持 cache key 的隔离/命中/失败后状态；relation 保持空键、重复键、未匹配和 source location 语义；ZIP 预检保持 position 恢复、DTD/路径/压缩比/大小限制。
- **依赖**：ME-100。
- **验证**：NPOI 原职责级测试全部通过；新增 Core helper 的成功/失败/取消测试。
- **风险**：异常 Provider/Stage/Message 漂移，或内部缓存键错误跨请求复用。
- **验收标准**：NPOI 行为不变，MiniExcel 不复制三套逻辑。

### Phase 2：创建 Provider 项目

#### Task ME-200：建立 Bing.Offices.MiniExcel 包

- **目标**：创建并接入独立 Provider/包。
- **现状/证据**：NPOI 目标 `net6.0;net8.0`，依赖 Core；版本集中在 `version.dev.props`。
- **修改范围**：新 csproj/AssemblyInfo、solution、版本属性、测试/benchmark 引用。
- **实施步骤**：沿用 `framework.props`；引用 Core 和稳定 MiniExcel；不得引用 NPOI；单独设置准确 PackageTags/Description；开启 XML docs/symbol/source link/readme/license 继承。
- **依赖**：ME-001。
- **验证**：restore、双 TFM build、dependency graph、pack 内容；检查包中无 NPOI 依赖。
- **风险**：net6 选择 netstandard2.0 资产带来的依赖版本冲突。
- **验收标准**：项目单独构建/打包，依赖方向符合目标图。

### Phase 3：P0 Export Pipeline

#### Task ME-300：导出 preflight 与计划构建

- **目标**：把 Workbook Request 编译成 MiniExcel 可消费且不会静默降级的导出计划。
- **现状/证据**：Core plan 已包含固定/动态列、getter、converter、value map、style/layout 元数据。
- **修改范围**：`MiniExcelExporter`、内部 preflight/plan builder；不新增公共 Request。
- **实施步骤**：校验 XLSX、Sheet 名/重复、模板/图表/图片/批注/复杂样式等能力；按 item type 调用共享 plan builder；第三方异常统一翻译为现有 BingOffices 异常。
- **依赖**：ME-101、ME-200。
- **验证**：每一 unsupported 特性在写任何字节前失败，Provider=`MiniExcel`、Operation/Stage 稳定。
- **风险**：漏检导致残缺文件却返回成功。
- **验收标准**：支持范围进入写入链，不支持范围 fail-fast。

#### Task ME-301：延迟行适配与 XLSX 写入

- **目标**：实现普通/多 Sheet、固定/动态列、基础类型和格式导出。
- **现状/证据**：NPOI writer 不可复制；MiniExcel 接受 object/dynamic/DataReader 等形态，具体方案由 Spike 决定。
- **修改范围**：Export row adapter/sheet writer/options builder。
- **实施步骤**：预编译列 getter；按 plan 顺序产出 header/value；converter 后执行 value mapping；保持 numeric/date/decimal/null/enum/bool 语义；动态列按稳定 key 合并；禁止 `source.Select(...).ToList()`。
- **依赖**：ME-300。
- **验证**：完整读取生成 XLSX，断言 Sheet/header/列序/单元格类型和值；controlled enumerable 证明单次、按需枚举。
- **风险**：每行 Dictionary、字符串格式化或 reflection 成为主要分配热点。
- **验收标准**：P0 导出契约通过，100K 不发生整表业务对象复制。

### Phase 4：P0 Import Pipeline

#### Task ME-400：Workbook/Sheet 读取与结构策略

- **目标**：实现 XLSX Stream/File/byte[]、多 Sheet、name/index、header/start row、空行与列范围策略。
- **现状/证据**：公共 request 已表达全部策略；MiniExcel 原始 API 是否保留物理位置需 Spike 确认。
- **修改范围**：Importer、sheet resolver/executor/runtime。
- **实施步骤**：先缓冲/预检再解析；一次解析 selector 并拒绝冲突；保留一基错误坐标；处理大小写、trim、缺列、未知/重复列、required header、MaxReadColumns。
- **依赖**：ME-101、ME-200。
- **验证**：真实 XLSX 正反例；完整 `ExcelWorkbookImportResult`/error 字段断言。
- **风险**：dynamic 返回字典可能丢失重复 header 或精确 column index。
- **验收标准**：共同场景与 NPOI 结构结果一致，无法保留的语义明确 fail-fast。

#### Task ME-401：原始值、转换、物化与错误契约

- **目标**：统一经过 `ExcelCellValue -> converter -> value map -> setter`。
- **现状/证据**：NPOI row materializer 为 mixed，Core plan 已预绑定转换器和 setter。
- **修改范围**：Core 最小 materialization kernel + MiniExcel raw value adapter。
- **实施步骤**：定义 object 到 `ExcelCellValue` 的 kind/raw/formula-cached-value 规则；保持日期系统、nullable、enum、decimal、culture；捕获用户扩展异常并保留 Stage/位置。
- **依赖**：ME-400。
- **验证**：DateTime/serial/string date、decimal 大数/scale/科学计数、nullable、非法值、公式缓存值完整测试。
- **风险**：MiniExcel 已做隐式类型转换，导致原始值/错误上下文不可逆。
- **验收标准**：不绕过现有 converter pipeline，错误字段完整或明确记录已批准差异。

### Phase 5：Mapping / Converter / Validation

#### Task ME-500：共享所有 Mapping 来源与缓存语义

- **目标**：证明 Attribute/Profile/Fluent/JSON/XML/ValueMap 均由 Core plan 生效。
- **现状/证据**：Factory 已实现优先级和缓存；Provider 只应消费计划。
- **修改范围**：MiniExcel tests，必要的 provider plan adapter；不得加入 MiniExcel Attribute。
- **实施步骤**：同一 DTO 分别构造五类来源和冲突合并；重复渲染；不同方向/Sheet/Profile/配置隔离；失败后重试。
- **依赖**：ME-301、ME-401。
- **验证**：直接断言输出/导入实体及 cache hit/miss/isolation/failure state。
- **风险**：Provider render plan 错误放入公共 mapping cache。
- **验收标准**：mapping plan 可跨 Provider 复用，render plan 只在 MiniExcel 内部且 key 含所有相关输入。

#### Task ME-501：Validation 与 Unique 资源边界

- **目标**：复用 Required/Date/Regex/Range/MaxLength/MaxValue/Unique、自定义/命名/动态校验。
- **现状/证据**：bindings 与 `UniqueTracker` 已 Provider-neutral；NPOI runtime 生命周期在 Provider 内。
- **修改范围**：共享 runtime kernel、MiniExcel validation tests。
- **实施步骤**：按 request validation mode 执行；每次 import 隔离 tracker；遵守 MaxTrackedUniqueValues/MaxErrors；失败行不污染增量状态。
- **依赖**：ME-401。
- **验证**：每个默认规则独立职责测试；10W/100W Unique 内存记录和超限异常。
- **风险**：跨行 HashSet 无界增长、静态状态串请求。
- **验收标准**：错误码/坐标/消息与 NPOI 合同一致，资源上限有效。

### Phase 6：Dynamic Column / Multi Sheet / Relations

#### Task ME-600：动态列、多 Sheet 与关系闭环

- **目标**：完成异构 Sheet、动态列和父子关系 P0。
- **现状/证据**：Request/plan 已表达能力；relation 已可下沉。
- **修改范围**：MiniExcel adapters/tests + Core binder。
- **实施步骤**：验证不同 Sheet 不同 DTO/mapping；动态 key/alias/order/placement/unknown policy/converter/validator；导入后绑定 parent-child/navigation。
- **依赖**：ME-301、ME-501、ME-101。
- **验证**：无匹配父行、重复/空 key、多个 Sheet、动态列参与完整 roundtrip。
- **风险**：为关系保留全部实体与 source location，限制极大导入的 streaming 上限。
- **验收标准**：功能正确；文档明确 relation 模式需要工作簿级物化，不能宣称 O(1) 内存。

### Phase 7：Async / Cancellation / Stream / File IO

#### Task ME-700：异步与流所有权合同

- **目标**：实现真实异步 IO 和可观察取消，不伪装底层同步阶段。
- **现状/证据**：公共流由调用方拥有；NPOI 用 staging；MiniExcel v1 async 边界待 Spike。
- **修改范围**：Exporter/Importer async、可选共享 staging/copy。
- **实施步骤**：所有可异步 Read/Write/Flush 使用 async API并传 token；同步 CPU/第三方段明确；不 dispose 调用方流；处理 seekable/non-seekable/slow/throwing/cancellation stream；定义 Position 变化。
- **依赖**：ME-001、ME-301、ME-401。
- **验证**：检测真实 `ReadAsync`/`WriteAsync` 调用；前/中/提交前取消；无 `Task.Run/.Result/.Wait()`。
- **风险**：第三方延迟枚举发生在 Task 完成后，使异常和取消逃出公共观察边界。
- **验收标准**：OperationCanceledException 标准化，stream 仍可由调用方使用，临时资源释放。

#### Task ME-701：文件原子提交与取消恢复

- **目标**：复用 `IFileExportCommitter`，保证失败/取消不损坏既有文件。
- **现状/证据**：Core 已有真实同步/异步 flush、replace/move/cleanup。
- **修改范围**：Exporter file 入口和直接职责测试。
- **实施步骤**：与 NPOI 相同调用 committer；写入失败保留原异常分类；提交失败使用 `BingOfficesFileCommitException`。
- **依赖**：ME-700。
- **验证**：新文件/覆盖、写中取消、commit 前取消、flush/replace/delete 失败矩阵和真实文件证据。
- **风险**：直接 Stream 不具备回滚，必须文档化。
- **验收标准**：无残留临时文件，既有目标字节完全不变。

### Phase 8：Unsupported Feature / Capability Preflight

#### Task ME-800：内部 capability 与 fail-fast

- **目标**：集中判断能力，禁止静默忽略。
- **现状/证据**：已有 `BingOfficesUnsupportedFeatureException`，不需新增异常类型或 public capability API。
- **修改范围**：MiniExcel internal feature validator。
- **实施步骤**：至少拒绝 XLS、Chart、Failure Workbook 以及 Spike 判定不完整的 Template/Style/Merge/Header/Image/Comment；支持策略必须有正向测试。
- **依赖**：ME-001、ME-300、ME-400。
- **验证**：每特性独立测试，断言未写输出、异常 code/provider/operation/stage。
- **风险**：request 中复合功能漏检。
- **验收标准**：所有矩阵项均为 Supported with tests 或 Unsupported with tests，无“静默部分成功”。

### Phase 9：Security / Resource Limits / XLSX Preflight

#### Task ME-900：输入安全与资源预算

- **目标**：MiniExcel 不绕过现有文件、ZIP/XML、行列、错误、图片、唯一值限制。
- **现状/证据**：NPOI ZIP preflight 可中性化；MiniExcel 内部 sharedStrings 缓存行为待验证。
- **修改范围**：Core preflight + MiniExcel importer/runtime tests。
- **实施步骤**：限制 input bytes、ZIP entries/ratio/uncompressed/worksheet/sharedStrings/styles/XML chars/depth；禁止 DTD/path traversal/重复 entry；公式仅读缓存值且不执行。
- **依赖**：ME-101、ME-400。
- **验证**：恶意/超限 XLSX 真实 fixture；失败后 stream position/临时文件/缓存目录清理。
- **风险**：第三方在 preflight 后仍产生不可控峰值；需进程级长测记录残余风险。
- **验收标准**：所有公共上限有直接测试，安全失败发生在第三方解析前。

### Phase 10：DI / Public API / Consumer Package

#### Task ME-1000：DI 与 Provider 选择语义

- **目标**：新增 `AddBingOfficesMiniExcel()`，注册现有服务与 MiniExcel 实现。
- **现状/证据**：NPOI 使用 `TryAdd*`，两 Provider 同容器时先注册者生效。
- **修改范围**：新 DI extension；不新增 resolver/factory。
- **实施步骤**：复用 default validation/mapping/file services；Exporter/Importer transient；固化“首个 Provider 注册生效”，文档建议每容器只选一个；交叉测试直接实例或独立容器。
- **依赖**：ME-200、ME-301、ME-401。
- **验证**：null、链式返回、幂等、生命周期、自定义服务保留、NPOI/MiniExcel 双注册两种顺序。
- **风险**：调用顺序隐藏配置错误；P0 接受但必须明确记录，后续 resolver 单独立项。
- **验收标准**：行为确定且测试化，不改变 `AddBingOfficesNpoi()` 既有语义。

#### Task ME-1001：API snapshot 与真实包消费者

- **目标**：把第四个生产程序集纳入双 TFM API 治理与 net6/net8 PackageReference 消费。
- **现状/证据**：snapshot 和 CI 当前硬编码三程序集/三包；消费者直接使用 NPOI 类型。
- **修改范围**：snapshot program/baseline/tests、consumer matrix、pack verification。
- **实施步骤**：新增程序集路径/身份；分类 public Provider/DI 类型；禁止第三方类型泄漏；consumer 只依赖 Bing 接口并完成 DI、Stream/File/byte[] roundtrip。
- **依赖**：ME-1000。
- **验证**：snapshot capture/diff、包解压、net6/net8 restore/build/run。
- **风险**：修改 approved baseline 需要成员级审批；未审批不得把新增程序集伪装为既有 baseline。
- **验收标准**：双 TFM snapshot、成员追溯、迁移记录和消费者报告完整。

### Phase 11：Unit Tests

#### Task ME-1100：默认实现职责级 Unit Test

- **目标**：每个受影响默认实现都有独立测试。
- **现状/证据**：MiniExcel 测试为零；项目规则禁止只靠综合测试间接覆盖。
- **修改范围**：MiniExcel unit tests、Core 抽取测试、现有 NPOI 回归。
- **实施步骤**：按 Exporter/Importer/Plan/RowAdapter/Materializer/Validation/DI/Preflight/Async/Stream/Unsupported 分类；测试完整值、文件结构或完整结构化错误。
- **依赖**：Phase 1-10。
- **验证**：双 TFM filtered + full unit tests。
- **风险**：只 mock MiniExcel 会错过真实序列化行为。
- **验收标准**：`symbol-test-map.md` 映射最终生产符号到项目/方法/关键行为。

### Phase 12：Integration Tests

#### Task ME-1200：真实 XLSX 集成矩阵

- **目标**：用真实文件验证 Office 结构和 roundtrip。
- **现状/证据**：现有 Integration 只引用 NPOI。
- **修改范围**：MiniExcel Integration tests/resources。
- **实施步骤**：覆盖普通、多 Sheet、动态列、日期/decimal/nullable/enum/bool、ValueMap、五类 Mapping、Validation、关系、non-seekable、文件提交；用独立 reader 或双方交叉读取避免自证。
- **依赖**：ME-1100。
- **验证**：net6/net8 Integration；完整 workbook 内容断言。
- **风险**：同一 Provider 写后读可能掩盖格式错误。
- **验收标准**：`integration-test-report.md` 记录 fixture、命令、结果和互操作限制。

### Phase 13：Cross-provider Contract Tests

#### Task ME-1300：共同能力契约套件

- **目标**：同一 Request 在共同支持集上得到等价业务结果。
- **现状/证据**：无第二 Provider；不要求二进制 hash 相同。
- **修改范围**：共享 contract fixture/test project（按 Phase 0 决策）。
- **实施步骤**：比较 Sheet、header、列序、固定/动态值、converter/value map、validation error 全字段、import result、异常 code；明确允许差异。
- **依赖**：ME-1200。
- **验证**：NPOI 与 MiniExcel 参数化执行，XLS/HSSF 与 XLSX/XSSF 原 NPOI 成功/失败测试继续通过。
- **风险**：过度追求底层单元格完全一致削弱 Provider 特性。
- **验收标准**：contract 比较业务语义，不比较 ZIP/hash。

### Phase 14：Performance / Benchmark / GC / 1M

#### Task ME-1400：Benchmark 与 controlled streaming 证据

- **目标**：测量而非预设 MiniExcel 的性能价值。
- **现状/证据**：Benchmark 只引用 NPOI；MiniExcel 宣称低内存但本仓库无证据。
- **修改范围**：Export/Import/ProviderComparison/RealIo/Stream benchmark 与 probe。
- **实施步骤**：1K/10K/100K/500K/1M，10/30 列、混合类型、动态列、多 Sheet、Stream/File；记录 Mean/Median/P95（基础设施支持时）、Allocated、Gen0/1/2、Peak Working Set、file size、rows/s、MB/s、cold/repeat；controlled enumerable 检测提前枚举/ToList。
- **依赖**：ME-1300。
- **验证**：100K 必跑；500K/1M 在 long-running/manual gate，无法执行须记录资源原因；输出原始 BenchmarkDotNet artifact。
- **风险**：每行 Dictionary、MemoryStream 双份 copy、Unique/relations 全量状态、sharedStrings cache 造成峰值。
- **验收标准**：`benchmark-report.md` 含环境、版本、原始指标、GC/内存和基于数据的推荐场景。

#### Task ME-1401：并发与资源矩阵

- **目标**：验证 transient Provider、静态委托/plan cache、临时资源在 10/50/100 并行请求下可控。
- **现状/证据**：现有静态 delegate cache 与 singleton mapping plan factory；MiniExcel 静态配置未知。
- **修改范围**：ResourceProbe/并发 tests。
- **实施步骤**：成功/第三方异常/取消/流异常/文件异常矩阵；检查文件句柄、临时文件、枚举器、shared strings cache、对象状态；设置并记录并发限制建议。
- **依赖**：ME-700、ME-900、ME-1400。
- **验证**：真实文件资源矩阵，不以单次 GC 数字代替泄漏证据。
- **风险**：第三方全局配置或缓存导致请求间污染。
- **验收标准**：资源报告包含可重复证据与部署限制。

### Phase 15：Documentation

#### Task ME-1500：Provider、API 与性能文档

- **目标**：让消费者按能力选择 Provider，且不直接依赖 MiniExcel API。
- **现状/证据**：README 仅列 NPOI；Excel docs 把实现描述为 NPOI DOM。
- **修改范围**：README、docs/Excel docs、新 `09-providers.md`。
- **实施步骤**：加入 `AddBingOfficesMiniExcel()`、公共接口示例、能力矩阵、unsupported、stream ownership、async 边界、resource/security、large dataset、两 Provider 选择/并存规则；同步版本与 TFM。
- **依赖**：ME-1400、ME-1401。
- **验证**：Docs tests 编译；示例不含 `using MiniExcelLibs`；链接检查。
- **风险**：文档先于实测结论承诺 P1/P2。
- **验收标准**：文档只声明已有测试支持的能力。

### Phase 16：Release Readiness / Final Review

#### Task ME-1600：CI、报告与最终门禁

- **目标**：达到可发布状态但不执行发布。
- **现状/证据**：CI 包数量和 snapshot/consumer/benchmark contract 均需扩展。
- **修改范围**：CI、所有任务报告、最终追溯。
- **实施步骤**：加入第四包 build/test/pack/verify/consumer；常规 CI 只跑 smoke benchmark，full benchmark 手动；生成 `execution.md`、`review.md`、两矩阵、`api-diff.md`、四测试/性能报告、`final-report.md`；独立 review 后修复 NEEDS_FIX。
- **依赖**：全部前序 Task。
- **验证**：第 9 节全量门禁、`git diff --check`、包内容、API 审批、NPOI 回归。
- **风险**：普通 CI 被长测拖慢；snapshot baseline 未经成员审批。
- **验收标准**：DoD 全绿或明确 BLOCKED 且不宣称完成；不 commit/push/publish。

#### Task ME-1601：最终生产符号追溯与范围收口

- **目标**：满足“最终生产符号 -> 测试方法”门槛，并确认未扩大为全库重构。
- **现状/证据**：项目规则要求成员级追溯；本计划包含若干候选下沉点。
- **修改范围**：`symbol-test-map.md`、final report、git diff 审计。
- **实施步骤**：逐符号列关键行为/测试项目/方法名；审计 public API、第三方类型泄漏、Task.Run/Result/Wait、ToList 热路径、NPOI 行为差异、无关文件变更。
- **依赖**：ME-1600。
- **验证**：`rg` 审计 + 独立 reviewer 对 plan/execution/diff/报告验收。
- **风险**：为复用而过度抽象 Core。
- **验收标准**：所有 P0 符号有直接测试，P1/P2/Unsupported 与矩阵一致，无未批准删除公共成员。

## 9. 仓库真实验证命令

以下命令来自当前 csproj、CI 和 API snapshot 工具；Executor 应在 PowerShell UTF-8 环境执行并把结果写入报告。

```powershell
dotnet restore .\Bing.Offices.sln
dotnet build .\Bing.Offices.sln -c Release --no-restore

dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Docs.Tests\Bing.Offices.Docs.Tests.csproj -c Release --no-restore

dotnet build .\build\ApiSnapshot\ApiSnapshot.csproj -c Release --no-restore
dotnet run --project .\build\ApiSnapshot\ApiSnapshot.csproj -c Release --no-build -- --identity-self-test true

dotnet pack .\src\Bing.Offices.Abstractions\Bing.Offices.Abstractions.csproj -c Release --no-restore -o .\artifacts\packages
dotnet pack .\src\Bing.Offices.Core\Bing.Offices.Core.csproj -c Release --no-restore -o .\artifacts\packages
dotnet pack .\src\Bing.Offices.Npoi\Bing.Offices.Npoi.csproj -c Release --no-restore -o .\artifacts\packages
dotnet pack .\src\Bing.Offices.MiniExcel\Bing.Offices.MiniExcel.csproj -c Release --no-restore -o .\artifacts\packages

dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- --list flat
dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- --filter "*ProviderComparison*"

git diff --check
```

新测试项目和 consumer 项目的精确命令只能在 ME-000 决定项目布局后加入 CI；不得在实施前臆造路径。

## 10. 测试与追溯最低要求

最终至少存在并追溯到具体测试方法的职责：

- `MiniExcelExporter`：普通/多 Sheet/动态列/完整类型/Mapping/converter/value map/stream/file/async/cancel/fail-fast。
- `MiniExcelImporter`：selector/header/row/column/whitespace/动态列/converter/validation/error/relation/resource/stream/async/cancel。
- `MiniExcelServiceCollectionExtensions`：注册、幂等、生命周期、自定义服务保留、双 Provider 顺序。
- 每个 Core 抽取：缓存隔离、命中/未命中、重复执行、失败后状态、NPOI 与 MiniExcel 直接测试。
- 每个 unsupported feature：完整结构化异常和零残缺输出。
- NPOI：XLS/HSSF 与 XLSX/XSSF 成功/失败不回归；未知实现实际异常类型固定。
- File/Async：真实 IO、资源矩阵、目标文件保护、临时文件清理、调用方 stream ownership。
- Output：断言完整 workbook/文件内容/结构化错误，不只断言字符串片段。

## 11. Breaking、兼容与迁移

- **预计 Breaking Change：无**。
- 允许的 public additive API：`MiniExcelExporter`、`MiniExcelImporter`、`AddBingOfficesMiniExcel()`；实际 public 面以 API review 为准。
- 不新增 `IMiniExcel*`、MiniExcel Request/Profile/Attribute、runtime resolver 或 public capability enum/interface。
- 不删除或改签名现有成员；公共 XML 注释中性化。
- 不改变 NPOI 默认格式、Failure Workbook、Mapping、异常类型或 DI 首注册语义。
- 迁移仅为新消费者从 `AddBingOfficesNpoi()` 主动改用 `AddBingOfficesMiniExcel()`；业务层继续注入 `IExcelExporter`/`IExcelImporter`。

## 12. 关键风险与停止条件

1. 若 MiniExcel v1.46.0 无法在 net6/net8 同时满足核心 API，ME-001 必须停止生产实现并提交版本/TFM 决策，不得私自采用 preview 或移除 net6。
2. 若原始读取 API 无法保留重复 header、物理坐标、raw/cached formula value，相关能力必须降级为 Unsupported 或提交明确兼容审批，不能伪造错误上下文。
3. 若多 Sheet API 必须完整物化全部数据，需以实测内存决定是否仍纳入 P0；不得用 marketing “streaming” 结论替代。
4. Relation、Unique、Failure Workbook 天然需要跨行/全局状态；文档必须区分“逐行读取”与“整个业务结果 O(1) 内存”。
5. 每行 Dictionary 是主要 GC 风险；先 benchmark，再选择 DataReader/custom enumerable/schema adapter，禁止提前引入复杂池化。
6. MiniExcel 内部临时 shared strings/cache 文件必须进入清理和并发矩阵。
7. 未经成员级 API baseline 审批、P0 cross-provider 通过、100K 证据、包 consumer 通过，不得标记可发布。

## 13. Definition of Done

- 独立 `Bing.Offices.MiniExcel` 项目/包存在且不依赖 NPOI。
- 实现现有 `IExcelExporter` 与 `IExcelImporter`，不建立第二套 Mapping/API。
- P0 XLSX、普通/多 Sheet、动态列、五类 Mapping、Converter、Validation、Relation、Stream/File/byte[] 通过直接测试。
- Async 使用真实异步 IO；取消、所有权、临时资源与原子提交有真实证据。
- Unsupported 功能全部在输出前结构化失败。
- XLSX preflight 与公共资源限制在两个 Provider 生效。
- NPOI Unit/Integration/API/Consumer 全部无回归。
- net6/net8 MiniExcel PackageReference consumer 成功。
- 双 TFM API snapshot、成员级审批和最终生产符号追溯完成。
- 100K 必测；500K/1M 完成或记录环境限制；性能结论来自真实数据。
- Unit/Integration/Package Consumer/Benchmark/Resource/Final 报告完整。
- README、Excel 文档、Provider capability matrix 与实测一致。
- CI smoke、pack 内容、`git diff --check` 通过。
- 不自动 commit、push、PR 或发布 NuGet。

## 14. 本计划范围摘要

- **P0**：稳定版 Spike；独立 Provider；XLSX 导入导出；多 Sheet；动态列；全部公共 Mapping/Converter/Validation；关系；Stream/File/byte[]；async/cancel；资源/安全；DI；API/包消费者；交叉契约；100K 与可扩展到 1M 的 benchmark。
- **P1/P2**：模板、复杂样式、merge/multi-header、图片、批注、Failure Workbook，仅在独立 Spike 和完整语义测试后实施；不得阻塞 P0。
- **Unsupported（P0 默认）**：XLS、Chart、Failure Workbook，以及未被 Spike 证明完整支持的模板/复杂样式/merge/header/image/comment。
- **后续独立 Task**：运行时 Provider Resolver/Factory、大范围 NPOI 重构、preview 2.0 迁移、复杂 Office DOM 能力对齐。
