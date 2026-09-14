# Bing.Offices RC 第二轮收敛实施计划

## 1. 任务信息

- Task ID：`BO-RC-20260908-002`
- 计划日期：2026-09-08
- 当前分支/HEAD：`master` / `e4093f554e7e81d642ff6722a1f8e44d185d39da`
- 规划基线：当前工作树包含 `BO-RC-20260907-001` 尚未提交的既有改动；执行前必须记录 staged/unstaged/untracked 分层，不得覆盖或回退。
- 任务性质：RC 第二轮发布证明、Async 文件边界、包消费、API 收敛、资源矩阵纠偏及有证据的热点优化。
- 本阶段：只生成计划，不修改业务源码、测试、配置、基线或报告，不执行 commit/push/PR/tag/publish。
- SQL 元数据规则：不适用；本仓库不含 `Bing.Data.Sql`，本轮计划在低优先级工程治理中清理该上下文污染。

## 2. 规划依据与范围

本计划基于用户提供的第二轮需求，并核对以下真实内容：

- `AGENTS.md`、`.agents/skills/create-plan/SKILL.md`；
- `framework.props`、`common.props`、`common.tests.props`、`asset/props/*.props`、各 `.csproj`；
- `.github/workflows/ci.yml`、`build/ApiSnapshot/*`、`build/api-snapshot-baseline.json`；
- Excel/CSV importer、exporter、stream extensions、atomic file committer、NPOI staging/template 调用链；
- Unit、Integration、Docs、Benchmark、ResourceProbe 及上一任务 PackageReference-only consumer；
- `BO-RC-20260907-001` 的 plan、execution、review、API/资源/测试/最终报告。

范围内：Runtime/包资产、Async 文件集成、Template 边界、PackageReference consumer、API 分类与重复入口、Resource Matrix、CSV/Unique/cache key/staging 性能、Benchmark/报告、Office 规则和 README 精修。

范围外：新增无关 Excel Feature、替换 NPOI、数据库/Word/PDF、无数据支持的架构重写、自动发布或 Git 写操作、将整个 Excel DOM 伪装为 fully async/zero allocation。

## 3. 当前实现与完成度判断

| 维度 | 当前真实状态 | 证据与结论 |
| --- | --- | --- |
| TFM 配置 | 已实现，需本轮重建证据 | Abstractions/Core 为 `netstandard2.0`；`framework.props`、NPOI、Unit、Integration 为 `net6.0;net8.0`；Benchmark/ResourceProbe/Docs 为 net8 合理单目标。上一任务本地双 TFM 已通过，但不能替代本轮 clean restore/build/pack。 |
| Unit/Integration | Unit 较完整；Async 文件 Integration 不足 | 上一任务 Unit 双 TFM各 716/716、Integration 各 15/15、Docs 10/10。现有 Integration 主要覆盖同步 FileStream、锁定、DI 与配置；没有职责级 Excel/CSV `*FromFileAsync`/`ExportToFileAsync` 矩阵。 |
| Async Runtime | 核心已实现，Template 边界部分完成 | CSV 使用真实 async Reader/Writer/Stream；Excel ImportAsync 使用 async copy；ExportAsync 使用 staging + CopyToAsync；生产未发现 Task.Run/Result/Wait。模板仍在 `CreateWorkbook` 中把 `request.Template` 直接交给同步 `WorkbookFactory.Create`。 |
| Atomic File | 已实现，需真实文件集成补证 | `ExportToFileAsync` 经 `IFileExportCommitter.CommitAsync`；Unit 已覆盖取消/清理/锁定，真实 XLS/XLSX/CSV 文件矩阵不足。 |
| Package Consumer | 一次性证据完成，仓库门禁未固化 | 上一任务 artifacts 下有真实 `PackageReference Include="Bing.Offices.Npoi" Version="2.0.0"` 的双 TFM consumer 并输出 `package-consumer-ok`，但它位于任务产物目录，不是稳定可重复的 `tests/Bing.Offices.Consumer.Net6/Net8` 门禁。 |
| Public API | 分层和 baseline 已建立；存在重复入口 | 当前 snapshot 为每个 TFM 168 个类型/1021 个成员，baseline 含 net6/net8、`approvedBy=jian玄冰`、`approvedAt`，但 `baselineCommit` 不等于当前 HEAD；Provider User API 已恢复。四个 Core `ExportToFile*` extension 仅转发同名接口方法，具备删除理由，但属于 Breaking Change。上一台账仍保留 165 等过期数字，需重新生成。 |
| 扩展测试 | 已完成 | 公开扩展 85/85 直接调用，独立 Review Round 13 PASS。若删除四个重复扩展，门禁/追溯/API baseline 必须同步变为 81 个，不能保留假签名。 |
| Resource Matrix | 设计存在关键缺口 | `StagingResourceMatrix` 虽有 `SemaphoreSlim(1,1)`，但只创建 `Math.Min(concurrency, MaxActualParallelism)` 个 operation；4/16/64 实际都只提交 1 个请求，`queuedRequestCount` 是推算值，不是实测排队。 |
| CSV writer | 明确热点，未优化 | `CsvRecordWriter.Write/WriteAsync` 每条记录新建/Dispose `CsvWriter`；异步每条还 Flush。字段/公式逻辑集中，可在不复制业务语义的前提下改为一次 writer 生命周期。 |
| UniqueTracker | 有基准，分配热点未治理 | `BeginRow/CommitRow/RollbackRow` 对 pending 字典 Clear；每行首次 reserve 又创建 HashSet/Dictionary，100K×多列产生短命集合。已有 UniqueJournal benchmark，但无 before/after。 |
| Mapping cache key | 已识别并有局部微基准，生产仍高分配 | 生产每次 key 创建均匿名对象 → JSON UTF-8 byte[] → SHA256 → Base64，并每次创建 serializer options/SHA256；现有 benchmark 只比较两种序列化路径，未独立覆盖真实 CacheKeyCreation/PlanHit/Miss。 |
| Excel staging | 功能完成，IO 放大未量化 | ExportAsync 为 NPOI → staging → destination；ExportToFileAsync 再经 atomic temp，可能双临时文件/二次写。现有 Stream benchmark 主要使用 MemoryStream，ResourceProbe 未完整输出实际 temp bytes/count/disk amplification。 |
| Benchmark | 部分完成 | 已有 Sync/Async、1K/10K/100K、CSV 1M、Unique/Dynamic；缺真实 FileStream、Delayed/Throttled Async、稳定 P95/P99/throughput、热点 before/after 和 staging 计数。 |
| 文档/工程治理 | 部分完成 | README 仍称“很多功能只建立基本结构”；`CellExtensions.ConditionalFormattin.cs` 拼写错误；根 AGENTS 含整段无关 SQL 规则。 |
| CI/依赖恢复 | 存在发布阻塞风险 | CI 使用 `dotnet restore Bing.Offices.sln --locked-mode`，但 `.gitignore` 忽略 `packages.lock.json`；上一任务已记录 clean restore 出现 `NU1004`。CI 也未运行 package consumer、验证包来源或上传对应发布证据。 |
| 外部门禁 | 未完成 | GitHub push/PR 双 TFM CI 最终绿灯与实际生产入口共享 `SemaphoreSlim(1,1)` 不在当前仓库可证明范围内，仍是发布条件。 |

按本轮最终验收项计算，当前约完成 `62%`：Runtime/Async/API/基础测试已成熟；主要缺口集中在 Async 文件 Integration、Template 明确合同、稳定 package consumer、真实并发排队、四类性能证据和本轮报告。当前不是“功能骨架”阶段，但也不能仅凭上一任务结果直接发布。

## 4. 设计决策与不可破坏合同

1. Abstractions/Core 保持 netstandard2.0；NPOI/Unit/Integration 保持 net6/net8；不为形式统一扩展 Benchmark/ResourceProbe/Docs TFM。
2. 保持 Excel NPOI DOM 同步边界；禁止 Task.Run、`.Result`、`.Wait()` 伪异步。
3. 2C4G 服务器批准策略保持：RequestedConcurrency 为 1/4/16/64，NPOI DOM ActualParallelism 最大为 1；所有逻辑请求必须真实创建并排队。
4. Provider User API（NpoiExcelImporter/Exporter 与 NPOI extension 容器）保持 public；Provider SPI 依据第三方实现需要判定，不以“不常用”为 internal 理由。
5. 删除四个重复 `ExportToFile*` extension 是本计划的 Breaking 候选决策：先生成成员级 diff/迁移说明并取得维护者审批，再改正式 baseline；替代调用为同签名接口实例方法。
6. Template Async 先做 characterization。若 AsyncOnlyReadStream 失败，优先比较 async staging 成本；若收益/复杂度不成立，则固化“模板读取同步”的明确合同，不保留失败测试。
7. 性能优化必须 before → candidate → after；吞吐/分配改善不超出重复测量噪声时，不保留复杂实现。
8. CSV 字段映射、转换、公式防护、动态列逻辑必须 Sync/Async 共用；只允许 IO 调度差异。
9. 不追求 Excel 全局 0 GC。报告按 Zero/Near-Zero/Reduced/Allocation Heavy 分类。
10. 所有新文本/报告/TRX 清单按 UTF-8；不覆盖当前用户已有 dirty changes。

## 5. 分阶段实施计划

### Phase 0（P0）：冻结当前工作树与重建发布基线

#### BO-RC2-P0-01 工作树、工具链与 TFM 基线

- **目标**：在不借用上一任务结论的前提下，证明当前工作树可重复 restore/build/test。
- **现状/证据**：HEAD 为 `e4093f...`，工作树包含上一任务 staged/unstaged/untracked 改动；TFM 配置已存在。CI 使用 locked restore，但 lockfile 被全局忽略，上一任务 clean restore 已出现 `NU1004`。
- **已确认文件**：`framework.props`、`common*.props`、`asset/props/target.feature.props`、`asset/props/package.props`、生产/测试 `.csproj`、`.github/workflows/ci.yml`、`.gitignore`、现有 `packages.lock.json`。
- **候选文件**：`ai_docs/tasks/BO-RC-20260908-002/baseline.md`、`artifacts/test-results/*`。
- **实施步骤**：记录 branch/HEAD/status 分层、dotnet info、OS；先决定并落实可重复依赖策略：优先对参与 solution restore 的项目提交 lockfile 并为其设置精准 `.gitignore` 例外，若不采用 lockfile 则必须同步移除 CI `--locked-mode` 并记录理由；在 clean checkout 验证；随后 Release solution build `/m:1`；单独确认 NPOI net6/net8 输出；按 CI 顺序运行 Unit/Integration 双 TFM及 Docs，避免既有 staging 临时前缀被并行测试进程互相污染。
- **依赖**：无。
- **验证**：`dotnet restore Bing.Offices.sln --locked-mode`；`dotnet build Bing.Offices.sln -c Release --no-restore /m:1`；`dotnet test ... -f net6.0/-f net8.0 -c Release --no-build --no-restore`。
- **风险**：NuGet 漏洞源不可访问产生既有 NU1900；需区分依赖恢复失败、warning 与代码失败。
- **验收标准**：clean checkout 的依赖恢复策略自洽且无 `NU1004`；NPOI 两套 DLL 实际存在；测试不是 0 executed；所有结果归档为本 Task TRX。

#### BO-RC2-P0-02 Pack 与资产检查

- **目标**：证明最终 nupkg 含正确编译资产和 XML 文档。
- **现状/证据**：上一任务曾 pack/consumer 成功，本轮尚无新产物。
- **已确认文件**：三个生产 `.csproj`、package props/version、CI pack 命令。
- **候选文件**：任务专属 `artifacts/packages/`、包清单/hash 报告。
- **实施步骤**：Release pack 到任务目录；用 Zip 结构化读取 nupkg；检查 NPOI `lib/net6.0`、`lib/net8.0`，Abstractions/Core netstandard2.0，DLL/XML/pdb 和 dependency group；记录 SHA-256。
- **依赖**：P0-01。
- **验证**：`dotnet pack Bing.Offices.sln -c Release --no-build --no-restore -o <task-packages>`；PowerShell/.NET Zip API 检查资产。
- **风险**：solution pack 混入非 packable 项目不等于失败；必须以三个目标包和包内资产为准。
- **验收标准**：每个包版本/路径/hash/TFM/dependency 可审计，零 ProjectReference 泄漏。

### Phase 1（P0）：真实 Async 文件 Integration 与 Template 边界

#### BO-RC2-P1-01 Excel Async File Integration

- **目标**：以真实 FileStream 证明 XLS/XLSX ImportFromFileAsync 与 ExportToFileAsync。
- **现状/证据**：现有 Integration 的 XLS/XLSX path roundtrip 使用同步流；Async cancellation 多为 Unit/MemoryStream doubles。
- **已确认文件**：`tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs`、Excel stream extensions、`NpoiExcelImporter/Exporter`、file committer。
- **候选文件**：按职责新增 `ExcelAsyncFileIntegrationTest.cs` 及小型 test helper。
- **实施步骤**：XLS/XLSX 参数化真实路径；验证新目标、替换目标、导入结果、导入/导出后立即独占打开/删除；预取消与可控 mid-copy/write；锁定目标、写入异常；对失败前后原文件 hash、临时文件集合和句柄做快照。
- **依赖**：P0。
- **验证**：Integration net6/net8 定向后全量；0 skipped。
- **风险**：Windows FileShare 锁语义与 Linux 不同；跨平台断言按平台能力分支但不得 skip 核心原子性。
- **验收标准**：两格式、两 TFM、成功/取消/锁定/异常均有真实文件证据，原文件不损坏且无 staging 残留。

#### BO-RC2-P1-02 CSV Async File Integration

- **目标**：证明 CSV 文件异步 IO 的编码、内容和取消合同。
- **现状/证据**：生产为真实 async；现有 Integration 仅同步 Stream roundtrip。
- **已确认文件**：`CsvEntityImporter/Exporter`、CSV stream extensions、Integration 项目。
- **候选文件**：`CsvAsyncFileIntegrationTest.cs`。
- **实施步骤**：真实 UTF-8 文件往返中文、逗号、引号、CRLF/LF、quoted newline；加入大文件但控制 CI 时长；预取消/mid cancellation、已存在目标、锁定/异常、文件释放与 temp cleanup。
- **依赖**：P0。
- **验证**：双 TFM 定向和全量 Integration；内容按完整预期记录/字段断言。
- **风险**：基于时间 CancelAfter 的测试易抖动；使用可控 stream/filesystem gate 或足够稳定的取消触发点。
- **验收标准**：两 TFM 真实 FileStream 路径全绿，0 skip，原始 TRX/测试文件大小归档。

#### BO-RC2-P1-03 Template Async Characterization 与决策

- **目标**：明确并验证模板输入是否支持 async-only read，而不夸大 Excel Async。
- **现状/证据**：`NpoiExcelExporter.CreateWorkbook` 同步调用 `WorkbookFactory.Create(new NpoiNonDisposingStream(request.Template))`，预期 AsyncOnlyReadStream 会失败。
- **已确认文件**：`NpoiExcelExporter.cs`、`ExcelWorkbookRequestTest.cs`、`AsyncPipelineTest.cs`、`docs/excel/async-io.md`。
- **候选文件**：`TemplateAsyncBoundaryTest.cs`、内部 template staging helper（仅在决策 A 时）。
- **实施步骤**：先用 XLS/XLSX AsyncOnlyReadStream characterization；记录同步 Read 调用。若失败，测量 template async staging 的内存/磁盘/生命周期；A：若可接受，先异步复制到 seekable staging，再同步交 NPOI，并覆盖 leaveOpen/异常/取消；B：若不实现，测试固化明确异常/同步读取边界并更新 XML/README，禁止宣称模板完全异步。
- **依赖**：P0，决策需先于最终 Benchmark/文档。
- **验证**：双 TFM Unit + 真实模板 Integration；模板流 leave-open/dispose、失败清理和取消。
- **风险**：实现 staging 可能形成第三份临时文件或重复 request；不得为通过 AsyncOnly 测试无条件放大大模板资源。
- **验收标准**：A 或 B 有量化理由、稳定测试和一致文档；不存在模糊“整个 Excel 输入 fully async”表述。

### Phase 2（P0）：永久 PackageReference-only Consumer

#### BO-RC2-P2-01 Net6/Net8 Consumer 门禁

- **目标**：把上一任务的一次性 artifact consumer 提升为可重复的仓库级发布验证。
- **现状/证据**：旧 consumer 已覆盖 DI、direct provider、Async bytes 和 NPOI extensions，但位于 `ai_docs/.../artifacts`，且未完整覆盖 Sync/File/CSV file。
- **已确认文件**：旧 `PackageConsumer.csproj/Program.cs/NuGet.Config`、nupkg 配置。
- **候选文件**：`tests/Bing.Offices.Consumer.Net6/`、`tests/Bing.Offices.Consumer.Net8/`、`build/RunPackageConsumers.ps1` 或同等 UTF-8 runner；不加入普通 solution ProjectReference 图。
- **实施步骤**：两个独立项目仅 PackageReference 最终本地 feed 包；隔离 NuGet cache/feed；验证 DI、IExcelImporter/Exporter、Npoi 直接构造、NPOI extensions、Excel/CSV Sync+Async Stream/File；检查 `project.assets.json` 不含 project reference；打印结构化成功标记；接入 CI 的 pack 后阶段，并上传 nupkg、consumer 日志及来源/hash 清单作为 artifact。
- **依赖**：P0-02、P1。
- **验证**：分别 restore/build/run net6/net8；清单记录实际运行 runtime 与包 hash。
- **风险**：固定 2.0.0 可能误吃全局缓存；runner 必须传任务 feed、清空独立缓存并禁用隐式外部同名包来源。
- **验收标准**：两个 consumer 在本地及 CI 均从本轮 nupkg 独立成功，报告可证明包来源与源码输出一致，失败会阻断发布门禁。

### Phase 3（P0/P1）：API 治理、重复入口与 Snapshot

#### BO-RC2-P3-01 Provider User API/SPI 全量复核

- **目标**：更新全量 API governance，不重新无脑 internal 化。
- **现状/证据**：上一任务已有 165 类型分类和 public provider；台账仍含已审批状态漂移。
- **已确认文件**：`PublicApiContractTest.cs`、上一任务 `api-classification.md`、三个程序集公开类型、AssemblyInfo。
- **候选文件**：本任务 `artifacts/reports/api-governance.md`。
- **实施步骤**：反射生成全部 exported type ledger；重点逐项回答 MappingConfigurationMerger、MappingProfileRegistry、UniqueTracker、ExceptionDispatcher、FileCommitter、Mapping Plan SPI、FactoryProvider、ValidationRules/DateTime rule 的第三方 Provider 需要；保持仅测试友元 IVT，拒绝生产程序集互友。
- **依赖**：P0。
- **验证**：分类测试、IVT 测试、public API ledger 无遗漏；报告 Category 只用指定五类。
- **风险**：EditorBrowsable(Never) 不等于 internal；不得据此误判兼容性。
- **验收标准**：每个公开类型有 Assembly/Visibility/Category/Decision/Reason，Provider User API 保持 public。

#### BO-RC2-P3-02 删除重复 ExportToFile Extensions（Breaking Gate）

- **目标**：移除只转发同名接口方法的四个冗余扩展，减少入口歧义。
- **现状/证据**：Excel/CSV 同步及异步 extension 均仅参数校验后调用接口同名实例方法；接口是正式替代入口。
- **已确认文件**：`CsvStreamExtensions.cs`、`ExcelStreamExtensions.cs`、对应职责测试、公开扩展门禁/追溯、Docs/consumer/API baseline。
- **候选文件**：breaking migration、candidate snapshot/diff。
- **实施步骤**：先生成删除四成员的 proposal；迁移为 `exporter.ExportToFile(...)` / `ExportToFileAsync(...)`；维护者批准后删除、更新调用点和测试；公开扩展计数从 Core 19→15、总计 85→81；禁止保留同 namespace 隐式转发兼容层。
- **依赖**：P3-01；正式 baseline 写入依赖维护者审批。
- **验证**：全仓无静态调用残留；Unit/Docs/consumer 编译运行；成员级 API diff 仅含批准删除。
- **风险**：源码 Breaking，虽然替代签名等价，静态类调用方必须迁移；未获审批则保持现状并在报告标记 BLOCKED，不自行更新 baseline。
- **验收标准**：有 Reason/Replacement/Breaking 记录和真实 approvedBy/approvedAt；无意外 API 漂移。

#### BO-RC2-P3-03 双 TFM API Snapshot Release Gate

- **目标**：冻结本轮最终 API。
- **现状/证据**：当前 baseline 双 TFM已批准并通过；若执行 P3-02 必然产生预期 diff。
- **已确认文件**：`build/ApiSnapshot/*`、baseline、CI compare、PublicApiContractTest。
- **候选文件**：本任务 candidate snapshot、成员级 diff、审批记录及批准后的 `build/api-snapshot-baseline.json`。
- **实施步骤**：生成 net6/net8 candidate；比较 Framework、exported types、分类和成员；记录 candidate 对应的 HEAD、dirty-state hash、程序集和 nupkg hash，禁止把上一任务的 `approvedBy/approvedAt` 自动继承为当前候选审批；维护者对成员级 diff 审批后更新 baseline；再次 compare；保留旧/新 hash 和迁移台账。
- **依赖**：P3-01/P3-02、P2 consumer。
- **验证**：`dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release --no-build -- --root output/release --baseline build/api-snapshot-baseline.json`；双 TFM API tests。
- **风险**：不能把 hash 相等替代成员审查，也不能由 Executor 自批。
- **验收标准**：零未批准 diff；baseline 含双框架、分类、approvedBy/approvedAt。

### Phase 4（P0）：Resource Matrix 真实请求并发

#### BO-RC2-P4-01 请求创建与单槽位门控纠偏

- **目标**：真实提交 1/4/16/64 个请求并测量排队资源，保持 ActualParallelism=1。
- **现状/证据**：当前只创建 `Math.Min(concurrency,1)` 个 operation，queued 数和 measured active 是推算值。
- **已确认文件**：`StagingResourceMatrix.cs`、ResourceProbe Program、上一任务 formal/smoke artifacts。
- **候选文件**：ResourceProbe 职责测试、任务专属 JSONL/Markdown。
- **实施步骤**：创建 `Enumerable.Range(0, requestedConcurrency)` 全部任务；每任务进入共享 SemaphoreSlim(1,1) 前后用原子计数记录 queued/active/maxActive；等待全部请求完成；采样进程/GC/LOH/temp 文件；不把 64 个 100K workbook 同时构建。
- **依赖**：P0；必须在正式性能测量前完成。
- **验证**：1/4/16/64 每档 submitted/completed 精确相等；maxActive=1；64 档实测 maxQueued≥63（允许采样字段另记，但事件计数必须确定）；0 error/residue。
- **风险**：100K×64 串行耗时很长；smoke 用 1K 验证队列完成，formal 用受控工作量/超时并明确资源包络，不用减少任务数作弊。
- **验收标准**：报告分列 RequestedConcurrency、SubmittedRequests、CompletedRequests、ActualParallelism、MaxQueuedRequests，全部来自实测。

#### BO-RC2-P4-02 2C4G 预算复验

- **目标**：在修正后的请求模型下复验已批准预算。
- **现状/证据**：旧 formal 36/36 的活动槽资源可参考，但不能证明排队成本。
- **已确认文件**：`tests/Bing.Offices.ResourceProbe/Program.cs`、`StagingResourceMatrix.cs`、上一任务 resource formal/approval 报告。
- **候选文件**：任务专属 ResourceProbe runner 配置、JSONL、`artifacts/reports/resource-report.md`；不改变生产并行度策略。
- **实施步骤**：固定 SDK/OS/电源/参数，warmup + 重复样本；分别记录 active operation 包络和 queue overhead；TempFile 默认策略；保留 PeakWorkingSet/LOH/GC/temp 数据。
- **依赖**：P4-01、P6 staging instrumentation。
- **验证**：结构化 budget evaluation，失败不得改阈值掩盖。
- **风险**：本地机器非 2C4G；可通过进程/容器限制或明确环境差异，不能把未限制机器结果直接称服务器通过。
- **验收标准**：1/4/16/64 全部完成且符合批准阈值；实际部署入口 gate 仍作为外部发布条件。

### Phase 5（P1）：CSV Hot Path 优化

#### BO-RC2-P5-01 单 CsvWriter 生命周期

- **目标**：消除每条记录创建/Flush/Dispose CsvWriter 的确定性热点。
- **现状/证据**：`CsvRecordWriter.Write/WriteAsync` 每调用构造 CsvConfiguration/CsvWriter；ExportCore 每行调用一次。
- **已确认文件**：`CsvPipelineSupport.cs`、`CsvEntityExporter.cs`、CSV Unit/Integration、CsvPipelineBenchmarks。
- **候选文件**：内部 record writer session/共享字段写入 helper；不新增 public API。
- **实施步骤**：先保存 1K/10K/100K/1M before；每次导出创建一次 CsvWriter/配置；header+rows 共用；Sync NextRecord，Async NextRecordAsync，末尾各 flush 一次；公式防护/字段格式化共用函数。
- **依赖**：P1 CSV integration、P0 benchmark baseline。
- **验证**：完整 CSV 字符串 parity、UTF-8/引用/换行/公式/动态列/取消、双 TFM；BDN before/after。
- **风险**：取消时 buffered data、leaveOpen 和 flush 异常语义可能变化；必须直接测试。
- **验收标准**：行为零漂移；Allocated 明显下降且吞吐不回退。建议保留阈值：100K allocated 至少下降 15% 或 Gen0 显著下降，Mean 不劣化超过 5%；否则说明噪声/回退候选。

### Phase 6（P1）：Unique、Cache Key 与 Staging 性能治理

#### BO-RC2-P6-01 UniqueTracker Allocation

- **目标**：减少每行 pending HashSet/Dictionary 短命分配。
- **现状/证据**：字典 Clear 后集合失去引用，下一行按 key 重新 new；已有 1/5 unique columns × 10K/100K benchmark。
- **已确认文件**：`UniqueTracker.cs`、Unique 职责测试、`UniqueJournalBenchmarks`。
- **候选文件**：内部可复用 per-key buffers/pool，不改变 public 构造器和语义。
- **实施步骤**：记录 before；设计 active key 与 reusable collection 分离，Begin/Commit/Rollback Clear 集合而非删除实例；保持 comparer、max count、first row、rollback 原子性；跑 after。
- **依赖**：P0；可与 P6-02 benchmark准备并行，代码改动分开。
- **验证**：重复/空白/null/max/commit/rollback/first-row 直接测试；BDN 10K/100K、1/5列。
- **风险**：复用集合可能保留大容量导致 retained memory 增长；同时报告 allocated 与 retained/PeakWorkingSet。
- **验收标准**：职责测试全绿；100K/5列 allocated/Gen0 有实质改善且 PeakWorkingSet不越预算，否则不保留复杂化。

#### BO-RC2-P6-02 Mapping Cache Key

- **目标**：分离并降低 cache hit 每次 key 计算的 allocation，同时守住语义隔离。
- **现状/证据**：生产每次 JSON UTF8 + SHA256 + Base64，并创建 options/SHA实例；现有 benchmark 未使用完整真实 key/plan hit/miss 三分法。
- **已确认文件**：`ExcelMappingPlanCacheKey.cs`、plan factory/cache、MappingValidationBenchmarks、cache isolation tests。
- **候选文件**：稳定 fingerprint/value key 内部实现、独立 `MappingCacheKeyBenchmarks`。
- **实施步骤**：增加 CacheKeyCreation/PlanCacheHit/PlanCacheMiss；确认配置对象可变性与 version 生命周期；先静态复用 serializer options/现代 hash API等低风险改进，再评估预计算 fingerprint；不以引用 identity 替代 semantic equality。
- **依赖**：P3 API 分类（确认保持 internal）、P0 baseline。
- **验证**：tenant/version/model/direction/config semantic equality、hit/miss/eviction、配置变化隔离、重复渲染；before/after BDN。
- **风险**：缓存错误比性能回退严重；任何 collision/陈旧 key 风险均否决优化。
- **验收标准**：隔离测试全绿；CacheKeyCreation 和 PlanCacheHit allocated 明显下降，miss 不显著退化；报告列出 hash/碰撞策略。

#### BO-RC2-P6-03 Excel Async Staging IO 放大

- **目标**：量化 Export/ExportAsync/ExportToFile/ExportToFileAsync 的临时写放大，再决定是否改策略。
- **现状/证据**：当前 Async staging + atomic temp 可能两次临时写；旧 ResourceProbe未形成四入口的 temp count/bytes 对照。
- **已确认文件**：`NpoiExcelExporter.cs`、INpoiAsyncStaging 实现/factory、file committer、StreamPipelineBenchmarks、ResourceProbe。
- **候选文件**：可观测 staging/file-system adapter、`ExcelStagingBenchmarks`、probe JSONL。
- **实施步骤**：1K/10K/100K 对照四入口，采集 elapsed/allocated/GC/PeakWS/temp count+bytes；如可获取再记进程 DiskBytesWritten。先不重写；只有现策略越预算或双写明显影响吞吐时，比较 Memory/TempFile/Hybrid，禁止大文件 byte[]。
- **依赖**：P1 Template 决策、P4 instrumentation。
- **验证**：输出文件 hash/可重开、取消/异常 cleanup；before/candidate 同机重复。
- **风险**：Windows 磁盘计数包含无关 IO；temp adapter 计数作为最低可信证据。
- **验收标准**：四入口资源差异可量化；策略维持或变更均有数据和审批，不凭直觉改写。

### Phase 7（P1）：Benchmark 真实 IO 与统计完善

#### BO-RC2-P7-01 Memory 与 Real IO 分轨

- **目标**：区分 API/state-machine overhead 与真实 IO throughput/latency。
- **现状/证据**：当前主要 MemoryStream；缺 FileStream、DelayedAsyncStream、ThrottledAsyncStream。
- **已确认文件**：Benchmark csproj、Program、StreamPipelineBenchmarks、MappingValidationBenchmarks。
- **候选文件**：`RealIoPipelineBenchmarks.cs`、受控 stream doubles、汇总脚本。
- **实施步骤**：保留 MemoryStream；增加文件同步/异步、延迟/限速流；参数 1K/10K/100K，CSV 1M按成本单独；warmup/repeat；child process 采集 PeakWS/P95/P99和吞吐。
- **依赖**：P5/P6 candidate 稳定后。
- **验证**：Benchmark 列表审查、Dry smoke、正式 run；原始 BDN JSON/CSV/日志完整。
- **风险**：BDN 默认统计不直接给业务端到端 P95/P99/PeakWS；用独立 probe 补充，不能伪造列。
- **验收标准**：报告清楚标注 Memory vs Real IO，包含 Mean/Median/P95/P99/throughput/allocated/Gen0/1/2/PeakWS。

#### BO-RC2-P7-02 Before/After 决策报告

- **目标**：让 CSV/Unique/cache/staging 每项优化可回溯。
- **现状/证据**：上一任务存在多批 Dry/formal 数据，但提交、机器和参数并不完全一致，不能直接作为本轮统一 before/after。
- **已确认文件**：`tests/Bing.Offices.Benchmarks/*`、既有 BDN 产物和 ResourceProbe JSONL。
- **候选文件**：本任务 Benchmark artifacts 与 `artifacts/reports/benchmark-report.md`，不在报告阶段顺手改生产代码。
- **实施步骤**：固定机器、SDK、commit/dirty hash、参数；before/candidate 各至少三次；计算中位数和变化率；按阈值保留/回退；分类 allocation 结论。
- **依赖**：P5/P6/P7-01。
- **验证**：每个摘要值链接到原始文件；无只报最佳样本。
- **风险**：当前工作树 dirty，before 必须在改热点前先保存，不能事后重造。
- **验收标准**：四项都有 Before/After/Decision；不声称整个 Excel zero GC。

### Phase 8（P2）：文件、Agent 规则与 README 治理

#### BO-RC2-P8-01 文件名与 Office 专属 AGENTS

- **目标**：修复拼写并移除无关 SQL 上下文污染。
- **现状/证据**：`CellExtensions.ConditionalFormattin.cs` 拼写错误；根 AGENTS 有整段 Bing.Data.Sql 门槛。
- **已确认文件**：上述源码文件、`AGENTS.md`。
- **候选文件**：无新增。
- **实施步骤**：仅 rename 为 `CellExtensions.ConditionalFormatting.cs`，确认 SDK glob和历史 diff；AGENTS 保留 UTF-8/Windows规则，替换为 Excel/CSV/NPOI、双 TFM、Async、API、测试、Benchmark/资源和追溯规则。
- **依赖**：所有 P0/P1 代码完成后，避免规则中途变化干扰执行。
- **验证**：build/test、rg 无旧拼写和 SQL 规则；编码检查。
- **风险**：Windows 大小写/rename 检测；使用非破坏性显式移动并核对 Git diff。
- **验收标准**：无行为/API diff；Office Agent 上下文准确。

#### BO-RC2-P8-02 README/Async/Breaking 文档精修

- **目标**：文档反映成熟能力和真实边界。
- **现状/证据**：README 仍称“很多功能只建立基本结构”；Async 文档已有 NPOI 边界但需同步 Template 决策、重复 extension 迁移。
- **已确认文件**：`README.md`、`docs/excel/*.md`、DocsConsumer tests。
- **候选文件**：Breaking migration/Release Notes 文档及最终 package consumer 示例；仅在现有目录结构确有对应职责时新增。
- **实施步骤**：改为核心 Excel/CSV 已实现；列出 Mapping/Validation/Dynamic/Failure/Async/Resource；同步 Template A/B、接口实例方法迁移、net6 EOL说明和 stream ownership；示例由最终 package consumer 编译。
- **依赖**：P1 Template、P3 API、P6 staging 决策。
- **验证**：Docs tests、consumer、过期术语搜索。
- **风险**：文档先于最终决策会漂移；最后统一同步。
- **验收标准**：无“骨架”误述、无 fully async/zero GC误导、Breaking replacement 可直接使用。

### Phase 9（P0 收口）：报告、全量门禁与独立 Review

#### BO-RC2-P9-01 任务报告集

- **目标**：形成当前 Task 独立、互相一致的发布证据。
- **现状/证据**：上一任务报告只能作为历史参考；本任务尚无绑定当前最终代码、包 hash 和测试产物的独立报告。
- **已确认文件**：上一任务报告模板、本任务各阶段生成的 TRX、nupkg、BDN、JSONL、API diff 和审批记录。
- **候选文件**：
  - `artifacts/reports/unit-tests.md`
  - `integration-tests.md`
  - `package-consumer.md`
  - `benchmark-report.md`
  - `resource-report.md`
  - `api-governance.md`
  - `BO-RC-20260908-002-final.md`
- **实施步骤**：按用户指定字段生成七份报告；final 固定为 20 节：本轮目标、修改摘要、net6/net8、Async Integration、Template Async、NuGet Consumer、API Governance、Resource Matrix、CSV 性能、UniqueTracker 性能、Mapping Cache 性能、Async Staging、Benchmark、GC/Memory、Breaking Changes、Unit Test、Integration Test、Remaining Issues、Release Blockers、Release Recommendation；记录 Task-ID、HEAD、dirty state、环境、命令、退出码、时间、TRX/BDN/JSONL/hash；明确 Executed/Skipped/Reason 和 Remaining/Blockers。
- **依赖**：全部前序 Phase。
- **验证**：报告数字由结构化原始产物解析；API/测试/资源计数一致；旧任务证据只作历史参考。
- **风险**：手工复制数字漂移；优先脚本化汇总。
- **验收标准**：Unit、Integration、Package Consumer、Benchmark、Resource、API Governance 六类专项报告及 20 节 final 均存在，所有声明可追溯。

#### BO-RC2-P9-02 最终固定顺序门禁

- **目标**：一次性证明最终候选。
- **现状/证据**：CI 已定义双 TFM 门禁，上一任务本地结果通过；本轮热点、consumer、并发和可能的 Breaking 变更完成后必须重新执行，不能沿用旧结果。
- **已确认文件**：`.github/workflows/ci.yml`、solution/project 配置、API snapshot runner、测试与探针入口。
- **候选文件**：最终命令日志、TRX、package/API/benchmark/resource 汇总和独立 `review.md`；生产源码不在本任务中再次修改。
- **实施步骤**：locked restore → Release build `/m:1` → Unit net6→net8 → Integration net6→net8 → Docs → pack → Consumer net6→net8 → API compare → benchmark/resource budgets → forbidden async/IVT search → diff-check →独立 Review。
- **依赖**：P9-01 前先生成原始结果，报告最后更新。
- **验证**：使用仓库真实命令；测试顺序运行以避免共享 temp prefix 竞态；GitHub CI/生产入口无法本地证明的项目保留 external gate。
- **风险**：性能运行时间长；不得用 smoke 替代 formal，只可分批执行并记录。
- **验收标准**：本地所有代码门禁通过、Review 无 MUST_FIX/SHOULD_FIX；未 commit/push/publish。

## 6. API Breaking 迁移提案

若维护者批准 P3-02，删除：

- `ExcelStreamExtensions.ExportToFile`
- `ExcelStreamExtensions.ExportToFileAsync`
- `CsvStreamExtensions.ExportToFile<T>`
- `CsvStreamExtensions.ExportToFileAsync<T>`

迁移为接口实例调用，参数顺序/返回/取消合同不变。`ExportToBytes*`、`ImportFromBytes*`、`ImportFromFile*` 继续保留。删除前发布 candidate diff；删除后更新 public extension coverage 从 85 到 81，并在 consumer/docs/Release Notes 提供 before/after。未获真实审批则不删除、不更新 baseline，并把该项标记为外部阻塞。

## 7. 最终验收矩阵

- [ ] NPOI net6/net8 实际 Release 输出与 nupkg 资产已验证
- [ ] Unit net6/net8、Integration net6/net8 均实际执行且 0 failed/0 skipped
- [ ] Excel XLS/XLSX Async File Integration 完成
- [ ] CSV Async File Integration 完成
- [ ] Template Async 边界按 A/B 决策完成测试、测量和文档
- [ ] Consumer.Net6/Net8 仅从本轮 nupkg restore/build/run
- [ ] Provider User API 保持 public；SPI 全量分类；无生产程序集间 IVT
- [ ] 四个重复 ExportToFile extension 已获批删除，或明确保持并记录阻塞
- [ ] API snapshot net6/net8 零未批准差异
- [ ] Resource Matrix 真实提交并完成 1/4/16/64 请求，ActualParallelism 最大 1
- [ ] Requested/Submitted/Completed/Actual/Queued 指标分离且实测
- [ ] CSV writer、UniqueTracker、cache key、Async staging 均有 before/after/decision
- [ ] MemoryStream 与 Real IO benchmark 分轨，正式指标和原始产物完整
- [ ] 生产无 Task.Run/Result/Wait 伪异步
- [ ] 文件名、AGENTS、README/Async/Breaking 文档与最终合同一致
- [ ] Unit/Integration/Package/Benchmark/Resource/API Governance/Final 报告生成
- [ ] 独立 Review 无未解决 MUST_FIX/SHOULD_FIX
- [ ] GitHub 双 TFM CI 最终绿灯（外部门禁）
- [ ] 实际生产入口共享单槽位 gate 已确认（外部门禁）
- [ ] 未自动 commit、push、PR、tag 或 NuGet publish

## 8. 推荐执行顺序与并行边界

严格顺序：`P0 baseline/pack → P1 Async Integration/Template → P2 Consumer → P3 API/approval/snapshot → P4 Resource correction → P5/P6 hotspot before-after → P7 formal benchmark → P8 governance/docs → P9 final gates/review`。

可并行的窄任务：Excel 与 CSV Async Integration；两个 Package Consumer；CSV/Unique/cache 三组 benchmark fixture；报告原始数据解析。主执行者必须负责 Template/API/资源策略决策、跨组整合和最终验证。任何生产优化都必须等待对应 before 数据，任何 API 删除都必须等待维护者审批。
