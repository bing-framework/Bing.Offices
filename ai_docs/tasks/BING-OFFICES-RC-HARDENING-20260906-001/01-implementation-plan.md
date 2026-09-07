<!-- AI_PLAN_STATUS: READY -->
# Bing.Offices RC 发布加固实施计划

## 1. 任务元数据

```yaml
task-id: BING-OFFICES-RC-HARDENING-20260906-001
task-type: implementation-release-hardening
priority: P0
language: zh-CN
execution-mode: continuous-resumable
baseline-branch: master
baseline-commit: 958e5b4886c2fe8df80ece3218d9fab1c57a0ec6
baseline-worktree: clean
breaking-change: allowed-with-evidence
auto-commit: false
auto-push: false
auto-create-pr: false
auto-tag: false
auto-publish-nuget: false
```

本计划基于 2026-09-06 当前仓库只读取证生成。规划阶段没有修改生产代码、测试、项目文件或现有报告，也没有运行 restore/build/test/pack/Benchmark。执行时必须重新确认 Git 状态；若基线 commit 或工作树已变化，不得套用旧数字，应在 `00-baseline.md` 记录差异并以执行时源码为准。

## 2. 当前基线与增量范围

上一任务 `BING-OFFICES-RC-CLOSURE-20260905-001` 已完成异常、DateTimeOffset、弃用主链、包消费和部分性能证据，但最终为 `No-Go`。当前提交已包含该轮成果。本任务不重复实现已验证内容，只复验并处理以下真实增量：

1. `ExcelStreamExtensions.ExportToFile` 与 `CsvStreamExtensions.ExportToFile` 仍在 exporter 公共观察边界之外调用 `AtomicFileCommitter`，文件 Create/Flush/Replace/Move 失败不会进入 exporter 的 Observer 链。
2. `ExcelMappingDiagnostic` 仍是 public；JSON/XML 的两个 `out IReadOnlyList<ExcelMappingDiagnostic>` 重载始终返回空集合，私有 `Load*Document` 仍携带无作用的 diagnostics 参数。
3. Public API hash 门禁缺少正式成员 baseline；现有 canonicalizer 不覆盖 abstract/sealed、泛型约束、参数名和默认值、属性访问器可见性及关键 attribute。
4. `TryAddPicture` 在 `Workbook.AddPicture` 后发生 `CreatePicture/Resize` 可恢复失败时可能返回 `false` 但留下副作用，合同尚未冻结。
5. `CsvEntityPipeline.cs` 约 863 行、`NpoiFailureWorkbookWriter.cs` 约 857 行、`NpoiExcelImporter.cs` 约 630 行、`NpoiExcelExporter.cs` 约 565 行；职责仍集中。
6. NPOI 仍目标 `netcoreapp3.1;net6.0;net8.0`。netcoreapp3.1/net6 已 EOL，且现代依赖对 netcoreapp3.1 产生支持警告；仓库没有已发布消费者约束。
7. 已有 1k/10k/100k Excel、CSV、Failure Workbook ShortRun，但缺少可比前后 baseline、批准预算、CSV 1M、图片/模板/样式/验证等完整矩阵、Failure Workbook 双 DOM 独立资源峰值和取消延迟分位数。
8. 文件提交、NPOI 扩展、文档和 warning-as-error 发布门禁尚未形成一次完整闭环。

明确非目标：本轮不实现伪异步 API，不用 `Task.Run` 包装 NPOI；不声称 Excel 真正流式、低 GC 或零分配；不新增生产程序集间 IVT；不自动提交或发布。

## 3. 开工协议

执行 Agent 开始前必须：

1. 完整读取根 `AGENTS.md`、本计划、上一任务 `final-report.md`/`review.md`/API/Benchmark/Resource 报告、`.editorconfig`、Props/Targets、CI 与打包配置。
2. 检查是否出现 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260906.md`；存在则逐项纳入，缺失则记录，不伪造。
3. 执行 `git status --short`、`git diff --stat`、`git rev-parse HEAD`，保护执行开始前的任何用户修改。
4. 创建其余任务报告文件，并运行：

```powershell
node .agents/scripts/task-state.mjs start BING-OFFICES-RC-HARDENING-20260906-001 --source <harness>
```

5. 确认 `.agents/runtime/current-task.json` 为 active、mode=`plan-execution`、taskId 正确。
6. 不修改本计划 checkbox；执行状态写入 `progress.md` 与各报告。

## 4. 强制设计决策

### 4.1 文件导出必须成为 exporter 正式能力

推荐方案：给 `IExcelExporter`、`ICsvExporter` 增加正式文件导出方法，由具体 exporter 在自身唯一公共异常边界内完成内容生成和原子提交。现有扩展方法可以保留为薄委托或删除重复入口，但不能继续自行调用 committer。

Core 提供最小 Provider SPI，例如 `IFileExportCommitter`，位于 Abstractions；默认实现位于 Core，标记 `EditorBrowsable(Never)`。SPI 只负责同目录临时文件、Flush、Replace/Move 与 cleanup secondary diagnostic，不持有 Observer。Exporter 在同一个 `try/catch` 中调用 committer，并由自身 dispatcher 观察最终 `BingOfficesException`，从而保证：

- 内容生成异常保持原实例；
- 文件系统阶段才生成 `BingOfficesFileCommitException`；
- Observer 与调用方拿到同一实例；
- Dispatcher 的实例标记确保每操作最多一次；
- direct construction 使用默认 committer，DI 可替换测试实现；
- 第三方 provider 可实现或复用稳定 SPI，不依赖反射、Service Locator 或生产 IVT。

若执行时发现直接扩展 `IExcelExporter`/`ICsvExporter` 会造成不可接受的泛型合同问题，可改用独立的 `IExcelFileExporter`/`ICsvFileExporter` 用户接口，但必须保证 DI 和扩展只走一个路径，并在 Breaking 报告解释选择。禁止保留两套各自提交文件的实现。

### 4.2 TryAddPicture 的失败原子性

`false` 必须表示没有新增 picture data、shape 或 drawing 副作用。执行时先通过可替换的 internal provider adapter 注入 `AddPicture/CreatePicture/Resize` 各阶段失败，确认 NPOI 能否可靠补偿：

- 若能完整回滚，则实现 rollback 并测试 picture/drawing/relationship 数量恢复；
- 若 NPOI 无法可靠回滚，则 `TryAddPicture` 只在进入任何可变操作前返回 `false`。一旦开始修改 Workbook，后续失败必须抛出明确的 `BingOfficesException` 子类并文档化可能的部分修改，不能返回假失败。

不要用 catch-all；取消、OOM、StackOverflow 原样传播。

### 4.3 TargetFramework 决策

执行 Phase 0 时搜索 README、CI、包元数据和 Git 历史中的明确支持承诺。若没有已批准的旧运行时兼容目标，则本计划默认：

- Abstractions/Core 保持 `netstandard2.0`；
- Npoi、Unit、Integration 与 PackageConsumer 移除 `netcoreapp3.1` 和 `net6.0`，运行目标收敛到 `net8.0`；
- CI、条件 PackageReference、lock file、API snapshot 循环、文档和报告同步清理。

若维护者明确要求保留某个 EOL TFM，则必须在 `00-baseline.md` 写明业务依据、支持期限、依赖警告和运行时安全责任，且相应测试必须实际运行。不能只为“兼容”保留无维护承诺的目标。

### 4.4 API baseline 与审批

在任何 public 修改前，从干净基线 commit 生成完整 before snapshot；不得直接用修改后的 hash 覆盖旧常量。新版 snapshot canonicalizer 至少编码：

- 类型 kind/visibility/abstract/sealed/static/base/interfaces；
- 泛型参数及 class/struct/new()/类型约束；
- 成员 visibility/static/abstract/virtual/override/sealed；
- 返回类型、方法名、泛型参数和约束；
- 参数 ref/out/in、类型、名称、optional/default；
- 属性/事件类型和 getter/setter/add/remove 可见性；
- 影响调用与序列化的关键 attribute。

生成 before/candidate 成员行和机器可读 diff。正式 baseline 文件必须带 schema、生成工具版本、baseline commit、approvedBy、approvedAt；在维护者没有填入审批信息前，状态只能是 `BLOCKED`，禁止为了绿测自动批准。

### 4.5 性能预算

执行 Agent先在基线 commit/等价 worktree采集 baseline，再在候选代码用相同机器、Job、输入和电源策略运行。建议预算只能作为提案写入报告，必须由维护者确认；未批准时不得判 Go。默认回归审查阈值提案：常用 1k/10k 场景 Mean/P95 不高于 baseline 10%，Allocated 不高于 baseline 5%，错误/取消路径不得增加泄漏；100k/1M 与 Failure Workbook 使用业务容量阈值单独审批。

## 5. Phase 0：可复现基线与门禁输入

### 工作项

1. 创建 `00-baseline.md` 与 `progress.md`，记录 commit、dirty state、OS/CPU/内存、SDK/Runtime、TFM、NuGet source、lock file 和工具版本。
2. 检查 `.sln`、全部 csproj、Props/Targets、CI、fixtures、docs、package assets、symbols、SourceLink、ApiSnapshot 输入。
3. 运行全仓扫描并分类结果：TODO/FIXME、NotImplemented/NotSupported、Obsolete/legacy/migration、伪异步、public/protected/internal、IVT、reflection/Compile、缓存和大分配。
4. 运行 locked restore、Release build、Unit/Integration/Docs 基线；保留 TRX。额外执行 warning-as-error 探针，先分类既有 warning，再决定代码修复或精准项目级例外，禁止全局 NoWarn。
5. 生成当前 public API before snapshot 与成员分类表。
6. 使用当前 Benchmark 源在未修改基线采集至少 smoke + 代表性 1k/10k，并创建独立 baseline artifacts；大规模场景可在 Phase 5 完成。
7. 建立“生产符号 → 直接测试方法 → 集成/包/性能证据”的追踪表，后续持续更新。

### 验收

- 基线失败和新回归可明确区分；
- before API snapshot 可重复生成；
- 所有命令和退出码进入报告；
- 未改生产代码前完成基线冻结。

## 6. Phase 1：异常、NPOI 扩展、日期与资源正确性

### 6.1 文件提交观察闭环

按 4.1 实现唯一文件导出路径，覆盖 Excel/CSV：Create、内容写入、Flush、Replace、Move、目标存在/不存在、取消、cleanup 二次失败、锁定/权限、已有目标保护和临时文件清理。

直接测试必须断言异常完整结构和引用相等：最终类型、Code、Operation、Provider、Stage、InnerException、Observer 次数、`Assert.Same(observer.Exception, caller.Exception)`。文件提交行为变化属于公共接口和默认实现变化，必须有独立默认实现测试类。

### 6.2 public NPOI 扩展审计

逐项审计 `WorkbookExtensions`、`SheetExtensions`、`RowExtensions`、`CellExtensions`、`CellStyleExtensions`、`FontExtensions`、图表与 DI 扩展：receiver null、范围/offset、HSSF/XSSF、未知 provider、取消/致命异常和返回副作用。

优先处理：`GetExcelFormat`、`GetSheets`、`SetAllSheetAutoCompute`、`GetAllPictureInfos`、`RemovePictures`、`MovePictures`、`TryAddPicture`。七个扩展容器保持 public。图片多参数 API 的值对象化仅在能够减少真实错误且 API diff 明确时实施，避免同时保留等价 overload。

### 6.3 日期复验

保持共享 parser，不新增第二套日期体系。通过 public API 覆盖：DateTime/nullable、默认 `yyyy-MM-dd`、显式 Culture、闰日/边界/空白/非法、DateTimeOffset 显式/固定 offset、XLS 1900、XLSX 1900/1904、公式缓存日期、CSV、两个 TZ 的导入→导出→重开→再导入。文档明确 DateTimeOffset 是 ISO 文本，不是 Excel 原生数值日期，其排序/公式行为按文本处理。

### 6.4 Failure Workbook 和输入资源

在现有输入/ZIP/XML限制基础上，增加 Failure Workbook build budget。优先设计显式配置对象或现有 limits 成员，分别限制：候选错误行、复制单元格、图片字节、目标工作簿估算对象量和序列化字节。预算检查必须在分配扩大前执行；超限抛 `BingOfficesResourceLimitException` 或产生明确失败诊断，取消/失败不写部分 destination。

对 `ErrorRowsOnly` 先评估低内存两阶段方案；若 NPOI DOM 无法避免双持有，则使用独立子进程测量并提供默认上限/拒绝策略，不把 `MaxSerializedBytes` 当峰值内存限制。

### Phase 1 验收

- 文件提交异常可观察、同实例、一次；
- NPOI 扩展失败合同直接测试完整；
- 日期矩阵和双 TZ 通过；
- Failure Workbook 在预算前可控拒绝；
- 相关 Unit/Integration 全绿后更新 progress。

## 7. Phase 2：废弃删除、TFM 与 API 收敛

### 7.1 删除空 diagnostics API

删除：

- `ExcelMappingDiagnostic`；
- JSON/XML 两个 `out IReadOnlyList<ExcelMappingDiagnostic>` public 重载；
- `LoadJsonDocument`/`LoadXmlDocument` 的 diagnostics 参数、空 List 分配和相关 using；
- Public API 分类、snapshot、测试、XML docs、示例和“迁移诊断”文字。

删除前后分别 `rg`。迁移方式是直接调用无 out 的 v2 API；不加 Obsolete 过渡层。

### 7.2 清理其他陈旧入口

复核并删除旧 mapping/CsvHelper/Expression/Regex/Legacy 测试名和误导文案残留。测试名称与当前职责一致，不能保留 `LegacyExcelAsyncBytes`、`ProviderSpiAndCompatibilityOverloads` 等虚假语义。

### 7.3 TFM 收敛

按 4.3 决策修改 `framework.props`、条件依赖、测试 TFM、CI、ApiSnapshot、PackageConsumer、locks 和 docs。每个保留 TFM 必须完成 build/test/package asset/consumer 证据；netstandard2.0 通过独立合同消费者编译验证。

### 7.4 Public API 工具和 Breaking 报告

实现 4.4 的完整 snapshot，生成 before/candidate/diff，并在 `02-api-breaking-changes.md` 列出删除、新增、修饰符、默认值、TFM、异常与 namespace 变化。分类只允许 User API、Provider User API、Provider SPI；SPI 使用 `EditorBrowsable(Never)`，不隐藏兼容层。

### Phase 2 验收

- 删除符号在 src/tests/benchmarks/docs/XML/API 中无残留；迁移表可引用旧名字但不可编译；
- 七个 NPOI 扩展容器仍 public；
- candidate diff 无未解释变化；
- 未获审批前 API 门禁保持 BLOCKED，而不是伪 PASS。

## 8. Phase 3：小步职责重构

重构按下列顺序逐步进行，每一步先移动/抽取一个职责、运行最近测试，再进入下一步。namespace 尽量保持稳定，目录迁移不应顺带改 public API。

### 8.1 CSV

1. 将 partial exporter/importer 机械拆为 `Csv/Exporting/CsvEntityExporter.cs` 与 `Csv/Importing/CsvEntityImporter.cs`。
2. 抽取列计划/头绑定、scalar formatting、conversion、validation/error collection 到现有或新 internal 类型。
3. Importer/Exporter 只保留参数校验、orchestration、公共异常边界和结果组装。
4. 复用同一日期/Mapping/Validation 实现，不复制 helper。

### 8.2 Failure Workbook

在已有 `NpoiFailureWorkbookDiagnostics` 基础上拆分：selection/plan、ErrorRows workbook builder、row/cell/style/metadata copier、picture/validation copier、annotation/summary、serialization/temp commit。每个类型只有单一职责；取消、错误顺序、样式和流所有权必须保持。

### 8.3 NPOI Import/Export

- Importer：抽取 source buffering/preflight/workbook open、sheet resolution、runtime resource context；复用现有 plan builder、sheet executor、row materializer、relation binder。
- Exporter：抽取 workbook/template creation、metadata、chart assembly 与 serialization；复用 plan builder 和 sheet writer。
- 泛型缓存必须有 miss/hit/concurrency/不同模型隔离测试；缓存 key 不得持有请求实例或无界用户数据。

### 8.4 Mapping 和 Abstractions 文件治理

将 `ExcelMappingPlanFactory` 的 key/cache、compile/bind、immutable runtime model 分离；先复用现有接口，只有出现第二实现或第三方 provider 边界才新增 interface。把 `ExcelImportPolicies.cs`、`BingOfficesExceptions.cs`、`ExcelCellStyle.cs`、`IExcelMappingPlan.cs`、`MappingProfileContracts.cs` 中主要 public 类型按职责拆文件，namespace 和签名不变。

### Phase 3 验收

- orchestration 类只协调，不再实现复制/转换/序列化细节；
- 新 internal 默认实现有独立职责测试；
- 没有 Service Locator、生产 IVT 或无界新增 cache；
- 每小步 API diff 无意外变化，Unit/Integration/Docs 通过。

## 9. Phase 4：完整测试与真实包消费

### 9.1 Unit

在附件要求基础上重点增加：文件导出观察全阶段、NPOI unknown/null/range、TryAddPicture post-mutation failure、Failure build budget、API canonicalizer 各维度、TFM asset、并发 cache/registry/observer。公共接口/默认实现/Provider 分支分别直接测试，禁止只靠综合测试。

### 9.2 Integration

真实 XLS/XLSX/CSV、模板/公式/样式/图片/图表/批注/富文本、多 Sheet/关系/动态列、两种 Failure Workbook、文件锁/权限/replace/move/取消/temp cleanup、DI/Observer/Profile。Windows 本机完整执行；Linux 通过 CI 或等价 runner，不能把 Windows 文件语义外推。

### 9.3 并发与取消

同一实例 1/4/16/64 并发，覆盖多模型/租户；在 input copy、ZIP preflight、row loop、NPOI write、Failure copy、file commit 中途确定性取消，断言无死锁、串扰、缓存污染、部分目标和临时文件泄漏。

### 9.4 PackageConsumer

每次最终生产变更后才冻结 pack。独立 consumer 只用 PackageReference 和全新 NuGet cache，覆盖 DI、Excel/CSV、Profile、JSON/XML、Observer、文件提交、NPOI 扩展和异常结构。结构化检查 assets 无 project 类型，逐一比较包内 DLL 与 Release 输出 SHA-256；审计 XML docs、snupkg、SourceLink、README、LICENSE、repository metadata 和依赖闭包。

### 报告

生成 `03-unit-test-report.md`、`04-integration-test-report.md`、`05-package-consumer-report.md`。每份记录环境、命令、退出码、数量、耗时、artifact、失败/skip 与未覆盖风险。Coverage collector 若缺失，新增正式工具并输出报告；不可用时标 BLOCKED，不能写 0% 或猜测。

## 10. Phase 5：Benchmark 与资源证据

### 10.1 正确性与基线

- Release、无 debugger、固定电源；BDN MemoryDiagnoser。
- before/current 使用相同机器和 Job；cold/warm、hit/miss 分开。
- 普通 BDN 方法不报告 `Process.PeakWorkingSet64`；PeakWorkingSet/managed heap/LOH 每场景独立子进程采样。
- 保留 console、Markdown、CSV、JSON、HTML 和环境清单；不删除不利结果。

### 10.2 必跑矩阵

1. Excel XLS/XLSX import/export：1k/10k/100k。
2. CSV import/export：1k/10k/100k/1M。
3. 模板、图片、图表、样式、公式、批注、动态列、多 Sheet。
4. Failure Workbook：0.1%/10%/100% 错误率，AnnotatedOriginal/ErrorRowsOnly。
5. Reflection old baseline vs cached delegate candidate；mapping cold/hit/miss；Profile explicit/scan。
6. Validation 小/大范围；MemoryStream/FileStream/non-seekable/byte[]。
7. DateTime/DateTimeOffset/date-only；Observer 0/1/10/failure。
8. 正常/恶意 ZIP；并发 1/4/16/64；取消延迟 P50/P95/P99。

记录 Mean/Median/Error/StdDev/Ratio、rows/s、Allocated、Gen0/1/2、LOH objects/bytes、peak managed heap、PeakWorkingSet、P50/P95/P99、文件大小、cache hit/miss 和模型数。

### 10.3 优化门槛

先测后改。只优化与行/列/Sheet/图片规模线性增长的热点；ArrayPool/Span/ObjectPool 等只有实测收益且异常/取消归还安全时采用。若高分配来自 NPOI DOM 架构，报告事实并提出后续 SAX/SXSSF RFC，不在本任务假装流式。

### Phase 5 验收

- `06-benchmark-report.md` 与原始 artifacts 完整；
- baseline/candidate 可比；
- 预算已由维护者明确批准才可 PASS；否则 `BLOCKED/No-Go`；
- 无低 GC/零分配/真正流式过度声明。

## 11. Phase 6：文档、独立 Review 与发布门禁

### 11.1 文档结构

将现有文档迁移/补齐到附件要求的主题文件。避免复制两套说明；旧文件若改名，更新所有链接和 Docs.Tests 输入。至少覆盖 getting started、DI、import/export、CSV、profiles、JSON/XML v2、日期/时区、validation、模板/样式/图片/图表、异常/Observer、资源/内存、NPOI 扩展、provider authoring、breaking changes、known limitations。

所有 C# fence 从 Markdown 原文提取并由 Docs.Tests 或最终 nupkg consumer 编译执行。明确 Excel 是 buffered DOM，CSV 读取流式但结果物化，byte[] 双份内存成本，DateTimeOffset 为文本，Failure Workbook 资源限制不是任意内存硬上限。

### 11.2 XML Documentation 与 warning gate

public API 补齐 summary/typeparam/param/returns/exception/ownership；interface 实现优先 inheritdoc。启用文档分析器并只修本轮/最终 public surface；不得通过全局 CS1591 NoWarn 达成门禁。Release warning-as-error 应在精准处理已知第三方/EOL警告后通过。

### 11.3 独立 Review

由未参与主要实现的 Reviewer 从 public API 重新走调用链，写 `07-final-code-review.md`。发现 P0/P1 后执行 Agent 必须修复、补测试、重跑受影响和全量验证，再复审；不能把可修复问题直接留给用户。

### 11.4 最终顺序

```text
locked restore
→ Release build / warning-as-error
→ Unit（全部支持 TFM）
→ Integration（全部支持 TFM/平台）
→ candidate API diff/approval check
→ clean pack
→ fresh-cache NuGet-only consumer
→ Docs compile/run
→ Benchmark smoke
→ full Benchmark/resource probes
→ independent review
→ git diff/status/check
```

生成 `08-release-readiness.md`。只有 API 批准、预算批准、全部支持平台/TFM测试、包消费和独立 Review 均无 P0/P1 才能 Go；任何一项缺失必须 No-Go/Conditional Go。

## 12. 必需交付文件

执行时创建并维护：

```text
ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260906-001/
├─ 00-baseline.md
├─ 01-implementation-plan.md
├─ 02-api-breaking-changes.md
├─ 03-unit-test-report.md
├─ 04-integration-test-report.md
├─ 05-package-consumer-report.md
├─ 06-benchmark-report.md
├─ 07-final-code-review.md
├─ 08-release-readiness.md
├─ plan.md
├─ execution.md
├─ review.md
├─ progress.md
└─ artifacts/
```

`execution.md`/`review.md` 为仓库工作流协议文件；编号报告为用户要求的可上传证据。不要用占位文件冒充完成。

## 13. 最终生产符号到测试追踪要求

至少在 Unit 报告维护下列映射：

| 生产职责 | 必须有的直接测试 |
| --- | --- |
| 新文件导出接口与默认 committer | Create/write/flush/replace/move/cancel/cleanup、Observer same/once、Excel/CSV |
| NPOI public extensions | 每容器 null/valid/unsupported；受影响 Provider 成功/失败分支 |
| TryAddPicture | preflight failure 无副作用、post-mutation failure 回滚或 throw 合同、fatal/cancel |
| JSON/XML loader 删除 | 剩余 v2 overload 正反例、删除成员反射/API diff |
| API canonicalizer | 修饰符、约束、optional/default、accessor、attribute 的 golden lines |
| TFM 变更 | csproj/CI/asset/package consumer 矩阵 |
| Failure build limits | hit/miss、边界前后、两模式、取消/cleanup/目标不变 |
| 拆分后的默认实现 | 每个关键 internal 职责独立测试类 |
| 缓存/registry | hit/miss、隔离、并发、执行上下文优先级、重复渲染 |

SQL 元数据测试规则本任务不适用，因为不修改 `Bing.Data.Sql`；执行报告需注明不适用，不得运行生产数据库。

## 14. 外部决策点与停止条件

以下是允许最终保持 BLOCKED 的外部输入：

1. 维护者对完整 API member diff 的批准身份和时间；
2. 维护者对 TFM 支持范围的明确反向决策（若不同于本计划默认）；
3. 性能/资源业务预算、批准人和可比基线环境；
4. Linux/macOS runner 或对应 CI 结果。

除上述外部输入外，普通构建、测试、重构和 Benchmark 失败都不是停止理由。执行 Agent 应继续修复所有本地可处理问题。最终若外部输入仍缺失，`execution.md` 使用 `PARTIAL` 或 `BLOCKED`，`08-release-readiness.md` 必须为 No-Go；不得更新 hash、放宽断言、Skip 测试或自批预算。

## 15. Definition of Done

- 文件导出异常进入统一 Observer 链，同实例且一次；内容错误与文件提交错误分类准确。
- public NPOI 扩展的 null/range/unsupported/Try 副作用合同闭环，扩展容器仍 public。
- 日期、资源、取消、Dispose 和跨平台文件语义有实际证据。
- 空 diagnostics API 与所有陈旧残留删除，不加 Obsolete 兼容层。
- TFM 决策进入代码、CI、包和文档；每个保留目标真实验证。
- 完整 API snapshot/diff 可审计，只有维护者批准变化进入正式 baseline。
- 大类按职责小步拆分，关键默认实现有直接测试，无生产 IVT/Service Locator/无界 cache。
- Unit、Integration、Docs、PackageConsumer、Benchmark/Resource 分报告与原始 artifacts 完整。
- clean pack 后 fresh-cache consumer 与 Release DLL 身份一致。
- 独立 Review 无开放 P0/P1。
- `08-release-readiness.md` 给出有证据的 Go/No-Go。
- 未 commit、push、PR、tag 或 publish。

## 16. 建议实施顺序

```text
基线/完整 before API snapshot
→ 文件导出观察闭环
→ NPOI Try/Throw 与图片副作用
→ 日期/Failure 资源复验与补强
→ 删除 diagnostics 与陈旧 API
→ TFM 收敛
→ 完整 candidate API diff
→ CSV/Failure/NPOI/Mapping 小步拆分
→ Unit/Integration/并发取消
→ clean pack + PackageConsumer
→ baseline/candidate Benchmark/Resource
→ Docs/XML/warning gate
→ 独立 Review + 修复
→ 最终门禁与 task-finish
```

执行过程中不得在 Phase 边界停下；每完成一个重要子任务立即更新 `progress.md`，记录修改文件、行为、命令、结果、阻塞与下一步。
