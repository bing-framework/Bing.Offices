# Bing.Offices RC 收敛与发布准备实施计划

## 1. 任务信息

- Task ID：`BO-RC-20260907-001`
- 计划基线：`master` / `4ecfe3fa7acea6b9d22d00e52e102c5ed1cf31bc`
- 基线工作区：clean（`git status --porcelain=v1` 无输出）
- 计划日期：2026-09-07
- 任务性质：RC 前兼容性恢复、API 治理、真实异步 IO、测试与发布证据闭环
- Breaking Change：允许，但必须有成员级差异、迁移策略和维护者审批
- 本计划阶段：只规划；不得修改业务源码、测试、配置，不得 commit/push/PR/tag/publish
- SQL 元数据规则：不适用；本任务不涉及 `Bing.Data.Sql`

## 2. 依据与范围

本计划基于以下真实内容形成：

- 根目录 `AGENTS.md`、`.agents/skills/create-plan/SKILL.md`；
- 用户提供的 `BO-RC-20260907-001` RC 收敛要求；
- `framework.props`、`common.props`、`common.tests.props`、生产/测试/Benchmark `.csproj`；
- `.github/workflows/ci.yml`、`Bing.Offices.sln`、`build/ApiSnapshot/*`；
- Excel、CSV、文件提交、NPOI Stream/DOM、DI 注册的真实生产调用链；
- Unit、Integration、Docs、Benchmark、ResourceProbe 与 Package Consumer 代码；
- `BING-OFFICES-RC-CLOSURE-20260905-001`、`BING-OFFICES-RC-HARDENING-20260906-001` 的历史报告。

历史报告只能证明当时的执行结果。执行本计划时必须在新 Task ID 下重新生成结果，不得把旧报告直接复制为通过证据。

### 2.1 范围内

- `Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi`；
- Unit、Integration、Docs Tests、PackageReference-only Consumer；
- API Snapshot、Benchmark、ResourceProbe、CI、README 与 Excel 文档；
- 本任务要求的 `artifacts/reports/*.md` 和原始测试/Benchmark/资源产物。

### 2.2 范围外

- Word、PDF 与数据库功能；
- 新增无关 Excel 功能；
- 自动发布 NuGet、Git 提交、Push、PR、Tag；
- 为降低 API 数量而进行无调用链证据的大规模 internal 化；
- 未获批准的性能优化或与 RC 门禁无关的目录重写。

## 3. 当前实现与完成度判断

| 维度 | 状态 | 当前证据 | 结论 |
| --- | --- | --- | --- |
| Excel/CSV/Mapping/Validation | 已真实实现 | `IExcelImporter`/`IExcelExporter`、`ICsvImporter`/`ICsvExporter` 均有生产实现；既有 Unit/Integration 覆盖 XLS/XLSX/CSV、映射、动态列、校验、失败工作簿 | 不是接口或 Stub；业务核心不重写 |
| 资源与异常边界 | 大部分已实现 | 已有 `IFileExportCommitter`、`AtomicFileCommitter`、异常 Observer、XLSX 预检、Failure Workbook budget 与直接测试 | 复用并扩展 Async，不另建平行体系 |
| Target Framework | 部分完成 | Abstractions/Core 为 `netstandard2.0`；`framework.props`、NPOI、Unit、Integration、CI 仅 `net8.0` | `net6.0` 发布与验证链缺失 |
| API 分层 | 部分完成 | 历史分类为 141 User API、9 Provider User API、15 Provider SPI；无生产程序集 IVT | `NpoiExcelImporter/Exporter` 仍 internal；需按新要求重新审计而非沿用旧结论 |
| Namespace | 未收敛 | NPOI 扩展源码位于 NPOI 程序集，但 namespace 仍为 `Bing.Offices.Extensions` | 与建议的 `Bing.Offices.Npoi.Extensions` 存在契约漂移 |
| Async API | 未实现 | 四个核心接口只有同步方法；生产代码未发现 `ImportAsync`/`ExportAsync`/`CommitAsync` | 本轮最大实现缺口 |
| 真异步 IO | 未实现 | `NpoiStreamCopier` 使用 `Read/Write`；CSV 使用同步 Reader/Writer；AtomicFileCommitter 使用同步 FileStream/Flush | CancellationToken 目前主要是同步检查点 |
| NPOI 同步边界 | 已确认 | `WorkbookFactory.Create` 与 `workbook.Write` 是同步 DOM 解析/序列化入口 | Async 设计必须诚实保留该同步阶段 |
| Sync 测试 | 较完整 | 历史 net8 Unit 为 492 passed + 1 API 审批阻塞；Integration 15/15；Docs 10/10 | 新改造后必须按双 TFM 重跑 |
| Async 测试 | 未实现 | 未发现 AsyncOnlyReadStream/AsyncOnlyWriteStream 或生产 Async 测试 | 必须新增真实 IO、取消、清理、并发测试 |
| Package Consumer | 部分完成 | 历史只有 net8 PackageReference-only consumer；包中 NPOI 只有 `net8.0` | 必须新增 Net6/Net8 两个独立消费者 |
| Benchmark/资源 | 部分完成 | 已有 1K/10K/100K、CSV 1M、ResourceProbe；`UniqueJournal` 在测量内插值字符串，Dynamic Cold 在测量内创建 10,000 rules；普通 Benchmark 读取进程历史 PeakWorkingSet | 工作负载存在污染，且无 Sync/Async 对照和批准预算 |
| API 门禁 | 阻塞 | `build/api-snapshot-baseline.json` 的 `approvedBy`/`approvedAt` 为空，且只有 net8 快照 | 不得由执行 Agent 自行批准或弱化测试 |
| 文档 | 部分完成 | README 未声明 .NET 6/.NET 8 与 NPOI Async 边界 | 需与最终合同同步 |

相对本任务最终清单，当前整体完成度估计约 `50%`：同步核心、异常/资源基础和部分发布工具已存在；双 TFM、全部 Async API、Async 直接测试、双消费者、Sync/Async 性能证据及正式 API 审批均未完成。

## 4. 设计决策与不变量

1. Abstractions/Core 保持 `netstandard2.0`；NPOI、Unit、Integration 目标为 `net6.0;net8.0`。Benchmarks、ResourceProbe、Docs Tests 默认保留 `net8.0`，除非兼容性验证发现必须多目标。
2. Sync 与 Async 共用 Mapping、Plan、Validation、Conversion、Materialization 和异常分类逻辑；仅拆分 IO 编排，不复制业务管线。
3. 生产代码禁止 `Task.Run`、`.Result`、`.Wait()` 伪异步。真正可异步的 Stream/File 操作使用 `ReadAsync`、`WriteAsync`、`CopyToAsync`、`FlushAsync` 并传递原始 `CancellationToken`。
4. NPOI `WorkbookFactory.Create`、Workbook DOM 操作和 `workbook.Write` 保持同步，并在 API/XML docs/README 明示。Excel Async 的含义是外围文件/Stream IO 真异步，不宣称 Fully Async Engine。
5. Excel Export Async 的 MemoryStream、临时文件或混合 staging 方案必须先通过同机 ResourceProbe 决策；不得仅凭代码风格选型，也不得默认引入大文件双份内存。
6. 文件 `Replace/Move` 没有对应异步文件系统 API时可同步执行，但内容传输与 flush 必须异步；文档明确提交元数据阶段的边界。
7. NPOI 类型直接相关的扩展统一进入 `Bing.Offices.Npoi.Extensions`；provider-neutral 扩展保留 `Bing.Offices.Extensions`。这是 Breaking Change，不保留长期双 namespace 转发层。
8. 基于当前构造函数依赖均为稳定 public contract，`NpoiExcelImporter` 与 `NpoiExcelExporter` 计划升级为 `public sealed` Provider User API，允许 DI 与直接实例化两种路径；执行前用成员级 API 表再次确认。
9. Provider SPI 保持最小 public + `EditorBrowsable(Never)`；真正实现细节 internal/private；禁止新增生产程序集间 `InternalsVisibleTo`。
10. .NET 6 已结束官方支持。此次恢复属于用户明确要求的兼容目标；最终报告必须记录安全/运行时风险，不得以该风险为由跳过 net6 验证。

## 5. 分阶段实施计划

### Phase 0：建立新任务基线与决策台账

#### BO-RC-P0-01 基线冻结与需求矩阵

- **目标**：把当前 commit、工具链、TFM、依赖、测试、API、包内容和历史阻塞转化为可复现的新任务基线。
- **现状/证据**：仓库 clean；旧任务结论为 net8 本地大部分通过，但 API 审批、预算和跨平台证据未闭环。
- **已确认文件**：`framework.props`、`common*.props`、全部 `.csproj`、`.github/workflows/ci.yml`、`build/api-snapshot-baseline.json`、两个旧任务报告目录。
- **候选新增文件**：`ai_docs/tasks/BO-RC-20260907-001/baseline.md`、`decisions.md`、`requirements-matrix.md`。
- **实施步骤**：记录 branch/commit/status、`dotnet --info`、SDK/Runtime/OS；在不修改 baseline 的前提下执行 restore、Release build、net8 Unit/Integration/Docs；生成现有包资产与 API candidate；逐项标记完成/部分/未实现/阻塞。
- **依赖**：无。
- **验证**：`dotnet restore Bing.Offices.sln --locked-mode`；`dotnet build Bing.Offices.sln -c Release --no-restore`；现有 net8 三类测试命令；API Snapshot compare 命令以工具 `--help`/现有报告为准。
- **风险**：锁文件或 NuGet 网络导致 baseline 无法恢复；必须区分环境失败与代码失败，不修改门槛。
- **验收标准**：基线报告含精确命令、退出码、测试数、包资产、已知失败；旧证据未被当作新结果。

#### BO-RC-P0-02 API 与 Async 设计冻结

- **目标**：在改代码前冻结 public/SPI/internal、namespace、接口签名、stream ownership、取消和异常合同。
- **现状/证据**：现有接口同步；NPOI provider internal；扩展 namespace 漂移；baseline 无审批人。
- **已确认文件**：四个 importer/exporter 接口、`IFileExportCommitter`、NPOI importer/exporter、DI 扩展、`PublicApiContractTest.cs`。
- **候选新增文件**：`api-classification.md`、`async-contract.md`、`api-diff-proposal.md`。
- **实施步骤**：输出每个公开类型/成员的分类治理表；列出 Async 签名与 Breaking 影响；定义调用方 stream 保持打开、模板 stream 生命周期、取消原样传播、Observer 同实例/单次、失败工作簿异步输出合同；由维护者确认最终 API proposal 后再更新 baseline。
- **依赖**：P0-01。
- **验证**：治理表覆盖全部 exported types；提案中的每个新增/删除/namespace move 都能映射到测试和迁移文档。
- **风险**：在签名未冻结前并行实施会造成二次 Breaking Change。
- **验收标准**：不存在未分类 public 类型；Provider User API/SPI/Internal 决策和 Async 限制无歧义。

### Phase 1：恢复 net6.0 + net8.0 矩阵

#### BO-RC-P1-01 项目、依赖与锁文件适配

- **目标**：恢复 NPOI、Unit、Integration 的双 TFM，同时保持 Abstractions/Core 为 netstandard2.0。
- **现状/证据**：`framework.props` 仅 `net8.0`，测试项目硬编码 net8；NPOI 和 DI 依赖只有 net8 条件。
- **已确认文件**：`framework.props`、NPOI/Unit/Integration `.csproj`、各 `packages.lock.json`、`common.tests.props`。
- **候选文件**：`.github/workflows/ci.yml`、条件 PackageReference；必要时新增仅用于 TFM 兼容的编译条件文件。
- **实施步骤**：设置 `net6.0;net8.0`；逐包验证 NPOI 2.7.4、Microsoft.Extensions、xUnit/Test SDK 兼容性；优先使用双 TFM 均兼容版本或条件版本；重新生成并审查 lock 文件；禁止为通过 net6 删除功能。
- **依赖**：P0。
- **验证**：`dotnet restore Bing.Offices.sln --locked-mode`；分别对 NPOI、Unit、Integration 执行 `dotnet build -f net6.0/-f net8.0 -c Release --no-restore`。
- **风险**：新 SDK 对 net6 EOL 发出警告、条件依赖造成 TFM 行为漂移。
- **验收标准**：两个 TFM 均有独立编译输出；依赖差异有理由；无功能性 `#if` 分叉或未记录警告。

#### BO-RC-P1-02 CI 双 TFM 门禁

- **目标**：让双 TFM 兼容性成为持续门禁，而非本地一次性结果。
- **现状/证据**：CI 只安装 .NET 8，只运行 net8 Unit/Integration。
- **已确认文件**：`.github/workflows/ci.yml`。
- **候选文件**：无新增 workflow，优先修改现有 CI。
- **实施步骤**：安装所需 6.0/8.0 SDK/runtime；对 Unit/Integration 使用 matrix；保留 build、diff-check、pack；上传 TRX/包清单作为失败诊断。
- **依赖**：P1-01。
- **验证**：本地使用与 CI 相同命令；PR/push CI 的两个 TFM job 均通过。
- **风险**：双矩阵显著增加 CI 时间；通过依赖缓存而非跳过测试控制耗时。
- **验收标准**：任一 TFM 编译或测试失败都会阻止合并。

### Phase 2：API 分层、Provider 可见性与 Namespace 收敛

#### BO-RC-P2-01 public/internal/EditorBrowsable 治理

- **目标**：将每个 API 放入 User API、Provider User API、Provider SPI 或 Internal 的正确层级。
- **现状/证据**：已有治理测试与历史分类，但 NPOI concrete provider 被 internal；多个 DTO/成员含 `EditorBrowsable(Never)`，需逐项复核。
- **已确认文件**：`PublicApiContractTest.cs`、Abstractions `Providers/`、`IFileExportCommitter`、Core 默认实现、NPOI importer/exporter、三个 AssemblyInfo。
- **候选文件**：`api-classification.md`、API snapshot candidate/baseline。
- **实施步骤**：生成全量类型和成员 ledger；公开 `NpoiExcelImporter/Exporter` 并审计构造器 XML docs；保留必要 SPI；internal 化仅有真实调用链证据的实现细节；确认 IVT 只指向测试；修正误导性的治理测试名称。
- **依赖**：P0-02、P1。
- **验证**：分类测试、IVT 测试、provider-neutral 不泄漏 NPOI 类型测试；Package Consumer 同时验证 DI 和直接构造。
- **风险**：现有第三方接口实现会受到新增 Async 成员影响；必须在 Breaking 文档单列。
- **验收标准**：0 Compatibility、0 未批准 Execution detail、0 生产 friend assembly；每个 public 成员有类别、理由和迁移策略。

#### BO-RC-P2-02 NPOI Extension namespace 迁移

- **目标**：消除 `Bing.Offices.Extensions` 与 Provider-specific NPOI 扩展的语义漂移。
- **现状/证据**：六组 NPOI 类型扩展和 NPOI DI 注册均在 `Bing.Offices.Extensions`；Core provider-neutral 扩展也使用同一 namespace。
- **已确认文件**：`src/Bing.Offices.Npoi/Bing/Offices/Extensions/*.cs`、Core Extensions、测试/Benchmark/docs using。
- **候选文件**：Package Consumer、API snapshot、README/`docs/excel/npoi-extensions.md`。
- **实施步骤**：将直接依赖 NPOI 类型的扩展及 NPOI 注册入口迁到 `Bing.Offices.Npoi.Extensions`；保留 provider-neutral API 在原 namespace；全仓替换真实调用点；删除旧 namespace 重复转发；记录 namespace Breaking migration。
- **依赖**：P2-01。
- **验证**：全仓搜索无旧 NPOI namespace 残留；Docs fences、Package Consumer、API snapshot 通过。
- **风险**：namespace 迁移是源码 Breaking；保留双入口会造成扩展解析歧义，因此不作为长期兼容方案。
- **验收标准**：0 namespace drift、0 duplicate extension、0 stale docs/snapshot。

### Phase 3：Async IO 基础设施与文件提交

#### BO-RC-P3-01 IFileExportCommitter/AtomicFileCommitter Async

- **目标**：增加真正异步的原子文件内容写入与 flush，同时保留同步合同。
- **现状/证据**：当前只有 `Commit(Action<Stream>)`，FileStream 未启用 Async，flush 同步。
- **已确认文件**：`IFileExportCommitter.cs`、`DefaultFileExportCommitter.cs`、`AtomicFileCommitter.cs`、`DefaultFileExportCommitterTest.cs`。
- **候选文件**：异步文件系统适配器与直接测试；优先在现有职责文件中扩展，避免无意义抽象。
- **实施步骤**：新增 `CommitAsync(string, Func<Stream,CancellationToken,Task>, CancellationToken, string)`；使用 Async FileStream、`FlushAsync`；复用同目录 temp + replace/move；保持写委托异常原样传播、文件系统异常分类为 FileCommit、取消/致命异常不包装、cleanup 附加诊断不覆盖主异常。
- **依赖**：P1、P2-01。
- **验证**：新目标/替换目标、pre-cancel、mid-write、flush/replace/move/lock 失败、委托失败、cleanup 失败、无临时文件残留；验证 Observer 同实例且一次。
- **风险**：Replace/Move 仍同步；文档需明确，测试不得伪装为异步元数据操作。
- **验收标准**：AsyncOnlyWriteStream 可证明调用 `WriteAsync`；生产无 Task.Run/Wait/Result。

#### BO-RC-P3-02 NpoiStreamCopier CopyAsync

- **目标**：为 Excel 输入外围复制提供带上限和取消的真异步路径。
- **现状/证据**：现有 `Copy` 以 81,920-byte buffer 同步 Read/Write，并在块边界检查取消。
- **已确认文件**：`NpoiStreamCopier.cs`、`StreamPipelineTest.cs`。
- **候选文件**：共享测试 Stream doubles。
- **实施步骤**：新增 `CopyAsync`，使用 `ReadAsync/WriteAsync` 和原 Token；保持 maxBytes、stream ownership 与异常类型一致；根据测量决定 ArrayPool 是否值得引入，禁止先声称 0 GC。
- **依赖**：P1。
- **验证**：AsyncOnlyReadStream、AsyncOnlyWriteStream、非 seekable、精确上限/超限、pre/mid cancellation、stream 保持打开。
- **风险**：netstandard2.0 与 net6/8 Stream overload 差异；使用各目标均真实支持的 API 并由双 TFM 编译证明。
- **验收标准**：同步/异步结果与异常一致，且同步 Read/Write 被禁止的 test double 仍能通过异步复制。

### Phase 4：CSV 真异步管线

#### BO-RC-P4-01 CSV ImportAsync/ExportAsync

- **目标**：提供完整的 CSV Stream 异步读写，不复制映射、校验和转换业务逻辑。
- **现状/证据**：`CsvEntityImporter/Exporter` 使用同步 StreamReader/Writer 和同步 record parser/writer。
- **已确认文件**：`ICsvImporter.cs`、`ICsvExporter.cs`、`CsvEntityImporter*.cs`、`CsvEntityExporter*.cs`、`CsvRecordReader/Writer`、CSV extensions/tests。
- **候选文件**：共享 CSV async reader/writer 或对现有 record 组件的 async 扩展。
- **实施步骤**：增加 `Task<CsvImportResult<T>> ImportAsync<T>`、`Task ExportAsync<T>`、`Task ExportToFileAsync<T>`；以 `Stream.ReadAsync/WriteAsync` 为底层并正确处理 Decoder/BOM/跨 buffer quoted field；同步与异步共享 record state machine 和实体处理函数；异步文件方法接入 `CommitAsync`。
- **依赖**：P3-01。
- **验证**：1/多字节 buffer、UTF-8 BOM、CRLF/LF、quoted newline、非 seekable、AsyncOnly streams、转换/校验/动态列/错误结果 parity、pre/mid read/write cancellation、文件清理和 stream ownership。
- **风险**：直接复制 parser 会造成 Sync/Async 漂移；必须提取共享状态机或共享 record 处理核心。
- **验收标准**：CSV AsyncOnlyRead/Write 全通过；取消抛原始 `OperationCanceledException`；无同步 IO 回退。

### Phase 5：Excel Async IO 与 NPOI 同步边界

#### BO-RC-P5-01 Excel ImportAsync 与 Failure Workbook Async

- **目标**：输入复制和失败工作簿输出使用真实 Async IO，DOM 解析/业务处理复用同步核心。
- **现状/证据**：导入先同步复制到 MemoryStream，再同步 `WorkbookFactory.Create`；Failure Workbook 写出也是同步。
- **已确认文件**：`IExcelImporter.cs`、`NpoiExcelImporter.cs`、`NpoiStreamCopier.cs`、Failure Workbook writer/serialization/filesystem。
- **候选文件**：内部导入执行结果/staging artifact；异步 Failure Workbook serializer。
- **实施步骤**：新增 `ImportAsync<TWorkbook>`；先 `CopyAsync` 到受资源限制的 staging，再同步创建/处理 DOM；重构失败工作簿阶段，使 NPOI 同步序列化与向调用方 sink 的异步传输分离；所有 await 使用正常异步传播，保留 Observer/错误结果合同。
- **依赖**：P3-02、P4 的共享测试基础。
- **验证**：AsyncOnlyReadStream、非 seekable、大输入上限、XLS/XLSX、mapping/validation/converter/dynamic columns、Failure Workbook mid-write cancellation、temp/handle cleanup、sync/async parity。
- **风险**：Failure Workbook 同时持有源/目标 DOM；异步 staging 可能增加峰值内存，必须纳入 Phase 7 资源决策。
- **验收标准**：输入读取和失败输出外层传输真实异步；NPOI 同步阶段被测试和文档明确；取消不包装。

#### BO-RC-P5-02 Excel ExportAsync staging 选型与实现

- **目标**：提供 `ExportAsync`/`ExportToFileAsync`，在 NPOI 同步序列化限制下控制峰值内存和阻塞边界。
- **现状/证据**：当前 `workbook.Write(destination)` 直接同步写入调用方流；MemoryStream 方案可能使大文件峰值翻倍。
- **已确认文件**：`IExcelExporter.cs`、`NpoiExcelExporter.cs`、`NpoiNonDisposingStream.cs`、ResourceProbe/Stream benchmark。
- **候选文件**：内部 staging strategy；不得暴露未经证明的 public option。
- **实施步骤**：在隔离 probe 中比较 A MemoryStream、B temp-file、C 基于规模的混合方案；记录时间、allocated、LOH、PeakWorkingSet、临时磁盘；批准一个默认方案；NPOI 同步写 staging 后以 `CopyToAsync` 写调用方目标；文件导出使用 `CommitAsync`，避免不必要的二次 staging；模板 stream 生命周期与同步 API 一致。
- **依赖**：P3-01、P0-02。
- **验证**：AsyncOnlyWriteStream（外围输出）、模板/图片/样式/图表、XLS/XLSX、非 seekable、pre/mid cancellation、目标保留、文件锁与 cleanup；100K 资源 probe。
- **风险**：同步 NPOI 序列化仍占用调用线程；不得用 Task.Run 隐藏。temp file 有磁盘/权限风险，MemoryStream 有 LOH 风险。
- **验收标准**：选型有原始测量与批准理由；Async API 行为 parity；外层目标使用异步写；文档不宣称 DOM 全异步。

### Phase 6：职责级测试、集成矩阵与 NuGet Consumer

#### BO-RC-P6-01 Unit/Integration Async 与双 TFM 矩阵

- **目标**：以直接测试证明每个新增生产符号、TFM 分支和关键边界。
- **现状/证据**：同步测试较完整；无 Async-only test double；Integration 仅 net8。
- **已确认文件**：现有 Unit/Integration 测试项目及相关 test classes。
- **候选文件**：按职责新增 Async 测试类，避免继续把所有场景堆入 `StreamPipelineTest.cs`。
- **实施步骤**：为四个接口 Async、两个 concrete provider、CommitAsync、CopyAsync、CSV reader/writer、Failure Workbook async 各建直接测试；覆盖 1/4/16/64 并发、共享 mapping cache、租户隔离和同一请求重复渲染；建立“生产符号 -> 测试方法”追踪表。
- **依赖**：P1-P5。
- **验证**：`dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0/-f net8.0 -c Release --no-restore`；Integration 同样双 TFM；失败必须保留 TRX。
- **风险**：仅比较返回值无法证明真异步；必须结合禁止同步调用的 stream doubles 与调用计数。
- **验收标准**：双 TFM Unit/Integration 全绿；0 skip；Sync/Async 矩阵全部有职责级测试；API 审批门禁除外不得存在已知失败。

#### BO-RC-P6-02 Net6/Net8 PackageReference-only Consumer

- **目标**：从最终 nupkg 证明编译资产、API、DI、Provider 和 namespace 可被真实消费者使用。
- **现状/证据**：历史 consumer 仅 net8，NPOI 包仅 net8 asset。
- **已确认文件**：三个生产 `.csproj`、历史 consumer 报告、package props。
- **候选文件**：`Consumer.Net6`、`Consumer.Net8` 隔离项目或同项目双 TFM runner；本地 feed 脚本。
- **实施步骤**：Release pack 到任务专属 feed；消费者只能 PackageReference，禁止 ProjectReference；清空/隔离缓存后 restore/build/run；真实执行 Excel/CSV Sync+Async、DI、直接 provider 构造、mapping、NPOI extensions、文件提交；检查包内 `lib/net6.0`/`lib/net8.0` 与 DLL hash。
- **依赖**：P1-P6-01、P2。
- **验证**：两个 consumer 分别 `dotnet restore/build/run -f net6.0/net8.0`；检查 assets.json 无项目引用。
- **风险**：本机缓存掩盖包缺陷；必须使用任务专属 feed/cache 并记录来源。
- **验收标准**：两个 runtime 都输出成功标记；包内容与 Release 输出一致；namespace 和 concrete provider 可用。

### Phase 7：Benchmark 与 Resource Probe 修正

#### BO-RC-P7-01 清理污染并增加 Sync/Async workload

- **目标**：让 Benchmark 只测目标职责，并提供可比较的 Sync/Async 性能数据。
- **现状/证据**：UniqueJournal 在 Benchmark 内构造插值字符串；Dynamic Cold 在每次工厂创建时构造 10,000 rules；PeakWorkingSet 被作为普通 Benchmark 方法读取。
- **已确认文件**：`MappingValidationBenchmarks.cs`、`StreamPipelineBenchmarks.cs`、Benchmark `Program.cs`。
- **候选文件**：任务专属 BenchmarkDotNet artifacts 与汇总脚本。
- **实施步骤**：Unique 输入移到 GlobalSetup；拆分 FactoryCreation/PlanCompileCold/CacheMiss/CacheHit；从普通 benchmark 移除历史 PeakWorkingSet 判断；增加 Excel/CSV Import/Export Sync/Async 1K/10K/100K，CSV 1M；用 BDN 统计 Mean/Median/P95/P99/Allocated/Gen0/1/2。
- **依赖**：P4-P5。
- **验证**：Benchmark 列表与参数审查；先 smoke 后固定环境正式 run；保留完整原始结果。
- **风险**：Async benchmark 用 MemoryStream 可能只测 CPU；同时增加真实文件/延迟流 probe，区分吞吐和 IO 等待。
- **验收标准**：每个报告行可追溯到原始 BDN 文件；不再声称伪 0 GC；before/candidate 同机同参数。

#### BO-RC-P7-02 资源、并发与 Excel staging 决策

- **目标**：用独立进程测量大文件、Failure Workbook 与 Async 并发的真实资源成本。
- **现状/证据**：已有 ResourceProbe 和 1/4/16/64 尾延迟雏形，但预算未批准、矩阵不完整。
- **已确认文件**：`tests/Bing.Offices.ResourceProbe/Program.cs`、Benchmark child-process probe。
- **候选文件**：新增模式参数、JSONL 原始输出、`resource-report.md`。
- **实施步骤**：独立子进程采集 PeakWorkingSet、GC/LOH、throughput、P95/P99；覆盖 Excel/CSV、100K、CSV 1M、Failure Workbook 双 DOM、模板/图片/样式/验证、并发 1/4/16/64、取消尾延迟；比较 Export staging A/B/C；由维护者批准预算。
- **依赖**：P7-01、P5-02。
- **验证**：重复运行、固定机器/电源/SDK/参数；结果含 before/candidate 与退出码。
- **风险**：进程历史峰值和机器噪声造成误判；使用隔离进程、warmup、重复样本和明确预算。
- **验收标准**：staging 方案有量化结论；所有发布预算有批准人、阈值和 PASS/FAIL，未批准保持 BLOCKED。

### Phase 8：API Contract、文档、报告与发布门禁

#### BO-RC-P8-01 API Snapshot 与 Breaking 审批

- **目标**：得到零漂移、可批准的成员级 API baseline。
- **现状/证据**：当前 baseline 只有 net8，`approvedBy/approvedAt` 为空；新增 Async、public provider、namespace move 必然改变快照。
- **已确认文件**：`build/ApiSnapshot/*`、`build/api-snapshot-baseline.json`、`PublicApiContractTest.cs`。
- **候选文件**：candidate/diff、Breaking migration 文档。
- **实施步骤**：按 net6/net8 生成 canonical snapshot；区分 shared API 与 TFM 差异；审查新增/删除/移动；记录第三方接口实现迁移；维护者批准后才写 approval metadata 和正式 baseline。
- **依赖**：P1-P6。
- **验证**：API compare 退出 0；分类、namespace、IVT、关键 XML docs tests 通过。
- **风险**：Executor 自批 baseline 会失去门禁意义。
- **验收标准**：0 stale snapshot、0 accidental diff；`approvedBy/approvedAt` 来自真实维护者审批。

#### BO-RC-P8-02 文档与 XML Documentation

- **目标**：公开支持范围与真实技术边界一致。
- **现状/证据**：README 未声明 net6/net8 和 Async/NPOI 同步边界；现有 docs tests 可编译 Markdown 示例。
- **已确认文件**：`README.md`、`docs/excel/*.md`、Docs Tests。
- **候选文件**：Async 指南、namespace/API migration、Provider direct-construction 示例。
- **实施步骤**：更新支持 TFM；分别说明 Excel Sync、Async IO、NPOI DOM 同步限制与 CSV 真异步；补 cancellation/stream ownership/atomic commit；更新 namespace using 和 Breaking migration；所有新 public API 补 summary/param/returns/typeparam/exception 或 `inheritdoc`。
- **依赖**：P2-P6、P8-01 proposal。
- **验证**：Docs Tests 全绿；README 示例由最终 nupkg consumer 编译运行；Release build 0 未解释 warning。
- **风险**：文档先于最终 API 会再次漂移；最后同步但从 P0 维护决策草稿。
- **验收标准**：文档不含 Fully Async Excel 等误导描述；所有 public Async API 和迁移路径可发现。

#### BO-RC-P8-03 报告、独立 Review 与最终发布判定

- **目标**：形成可审计的 RC Go/No-Go 证据链。
- **现状/证据**：旧报告属于其他 Task ID，当前 API/预算审批仍阻塞。
- **已确认文件**：历史报告格式、用户指定报告路径。
- **候选文件**：`artifacts/reports/unit-tests.md`、`integration-tests.md`、`benchmark-report.md`、`resource-report.md`、`BO-RC-20260907-001-final.md`；本任务 `execution.md`、`review.md`。
- **实施步骤**：按 TFM 汇总测试数/耗时/失败；保存 BDN 原始结果；列出生产符号到测试方法追踪；由独立 Reviewer 对 plan/execution/diff/结果审查；修复 MUST_FIX/SHOULD_FIX 后全量回归；仅全部门禁满足时标记 Go。
- **依赖**：全部前序 Phase。
- **验证**：最终固定顺序为 clean restore、Release build、net6/net8 Unit、net6/net8 Integration、Docs、pack、两个 consumer、API compare、Benchmark/Resource budget、`git diff --check`、独立 review。
- **风险**：跨平台 runner、API 审批或预算可能需要外部输入；必须保持明确 BLOCKED，不得伪造通过。
- **验收标准**：用户最终清单全部勾选；报告含 Task ID、commit、dirty state、环境、命令、退出码、原始产物链接、未完成项与发布结论。

## 6. 关键调用链验收矩阵

| 调用链 | Sync | Async | 关键证明 |
| --- | --- | --- | --- |
| Excel Stream Import | 保留 | 新增 | AsyncOnlyRead、XLS/XLSX、非 seekable、DOM 边界 |
| Excel Stream Export | 保留 | 新增 | AsyncOnlyWrite 外围输出、staging 资源报告 |
| Excel File Export | 保留 | 新增 | async temp write/flush、replace/move、cleanup |
| Failure Workbook | 保留 | 新增异步输出 | mid-write cancel、双 DOM PeakWorkingSet、目标不污染 |
| CSV Stream Import | 保留 | 新增 | 真 ReadAsync、quoted/BOM/多字节边界 |
| CSV Stream Export | 保留 | 新增 | 真 WriteAsync/FlushAsync、formula policy parity |
| CSV File Export | 保留 | 新增 | CommitAsync、文件锁、目标保留 |
| Mapping/Validation/Converter/Dynamic | 同一核心 | 同一核心 | 完整结果与错误 parity |
| DI/直接 Provider | 支持 | 支持 | Net6/Net8 nupkg consumer |

## 7. Breaking Change 与迁移策略

1. `IExcelImporter`、`IExcelExporter`、`ICsvImporter`、`ICsvExporter`、`IFileExportCommitter` 新增成员会破坏第三方接口实现；版本说明提供完整方法模板，并在最终发布版本按 SemVer 处理。
2. NPOI 扩展 namespace 从 `Bing.Offices.Extensions` 迁到 `Bing.Offices.Npoi.Extensions`；迁移文档给出 using 对照，不保留永久兼容层。
3. `NpoiExcelImporter/Exporter` 从 internal 变 public 是 additive，但其构造器和扩展点进入正式 API baseline，后续不得无治理变更。
4. Sync API 不删除；Async 不是另一套业务语义。已有调用方可以继续使用 Sync。
5. 所有 Breaking 只在 API 提案批准后落地；删除/修改快照不能替代迁移说明。

## 8. 风险与人工门禁

| 风险/门禁 | 责任方 | 解除条件 |
| --- | --- | --- |
| API baseline 未批准 | 维护者 | 审查成员级 diff，提供真实 `approvedBy/approvedAt` |
| 性能/资源预算未批准 | 维护者 + Executor | 固定环境 before/candidate 报告完成，批准阈值 |
| Excel Async 同步 DOM 限制 | Executor | 测试、XML docs、README 明确；禁止 Task.Run |
| Export staging 资源放大 | Executor + 维护者 | A/B/C 独立 probe 后批准默认方案 |
| .NET 6 EOL | 维护者 | 明确仍作为兼容目标，并在发布说明披露风险 |
| 跨平台 runner 不可用 | 维护者/CI | 至少 Ubuntu CI 双 TFM 全绿；若承诺 macOS，提供对应 runner |
| 独立 Review | 独立 Reviewer | `review.md` 无未解决 MUST_FIX/SHOULD_FIX |

## 9. 最终完成条件

- [ ] NPOI、Unit、Integration 支持并验证 net6.0 + net8.0
- [ ] 全量 API 分类表完成，Provider concrete/SPI/Internal 边界正确
- [ ] 无生产程序集间 InternalsVisibleTo
- [ ] NPOI namespace 漂移和重复入口为 0
- [ ] Excel/CSV Sync API 保持，Async API 完整
- [ ] 真正可异步的 IO 使用 ReadAsync/WriteAsync/CopyToAsync/FlushAsync
- [ ] 生产代码无 Task.Run/.Result/.Wait 伪异步
- [ ] CancellationToken 贯穿真实 IO，取消不被包装
- [ ] Atomic File Async、失败清理和文件锁场景通过
- [ ] AsyncOnlyReadStream/AsyncOnlyWriteStream 证明真实异步
- [ ] Sync/Async mapping、validation、converter、dynamic、error parity 通过
- [ ] Unit/Integration 双 TFM 全绿且 0 skip
- [ ] Net6/Net8 PackageReference-only consumer 通过
- [ ] API snapshot 无意外差异并获得维护者审批
- [ ] Benchmark workload 污染已修复，Sync/Async 结果与原始产物保存
- [ ] ResourceProbe 覆盖 PeakWorkingSet/LOH/GC/100K/CSV 1M/Failure/Concurrency
- [ ] 性能和资源预算获得批准且通过
- [ ] README、XML docs、迁移文档与最终 API 一致
- [ ] Unit、Integration、Benchmark、Resource、Final 报告生成
- [ ] 最终生产符号到测试方法的可追溯映射完成
- [ ] 独立 Review 无未解决 MUST_FIX/SHOULD_FIX
- [ ] 未自动 commit、push、PR、tag 或 publish

## 10. 推荐执行顺序

严格按 `P0 -> P1 -> P2 -> P3 -> P4 -> P5 -> P6 -> P7 -> P8` 执行。P1/P2 完成后冻结 API；P3-P5 完成后冻结行为；P6 后再批准 API baseline；P7 的资源结论决定 Excel Export staging；P8 只做文档、审批、Review 和最终全量回归，不在发布收口阶段引入新功能。

## 11. Phase 9：公开扩展方法职责级测试补全（已批准）

### BO-RC-P9-01 覆盖清单门禁

- 新增 `PublicExtensionCoverageTest`，反射枚举 Core/NPOI 程序集中的公开扩展方法。
- 建立完整方法签名到职责级测试方法的映射，覆盖当前 Core 19 个、NPOI 66 个，共 85 个声明及全部重载。
- 断言映射无遗漏、无失效签名，且目标测试真实存在并带 `[Fact]` 或 `[Theory]`。
- `internal` 类型上的扩展不纳入公开 API 门禁。
- 行为测试使用静态扩展类限定调用，避免同名接口或 NPOI 实例方法造成假覆盖。

### BO-RC-P9-02 Core 扩展

- 新增 `CsvStreamExtensionsTest`、`ExcelStreamExtensionsTest`，直接覆盖各 8 个同步/异步入口。
- 通过 tracking importer/exporter 验证参数、options、请求和原始 `CancellationToken` 转发。
- 覆盖正常结果、null、空白路径、异常传播、预取消和异步取消，以及内部流的释放和读写属性。
- 扩充 Mapping Profile 注册测试，覆盖参数校验、非法 Profile、名称回退、链式返回、重复注册、扫描异常及失败原子性。

### BO-RC-P9-03 NPOI 基础扩展

- 按 `ExcelFormat.Xls/Xlsx` 参数化 Workbook、CellStyle、Font、Row 测试。
- 直接覆盖 Workbook 6 个、CellStyle 17 个、Font 4 个、Row 5 个和 Cell 9 个公开扩展声明。
- 固化格式识别、样式属性、单元格读写/转换、缓存、条件格式以及合并行为的当前运行时合同。

### BO-RC-P9-04 NPOI Sheet、合并与图片扩展

- 覆盖 Sheet 行操作、合并区域、图片添加/读取/过滤/移除/移动的全部公开重载。
- HSSF/XSSF 均具备正常场景，失败场景断言当前异常类型、参数名及对象不变性。
- 图片签名覆盖 PNG/JPEG/GIF/未知类型；两个 `TryAddPicture` 重载分别覆盖成功和拒绝/异常路径。

### BO-RC-P9-05 DI 与证据收口

- 补强 `AddBingOfficesNpoi` 的链式返回、重复注册、调用方替换、生命周期和双 TFM 一致性测试。
- 更新 `execution.md`、最终报告和生产符号到测试方法追溯表。
- 完成双 TFM Unit/Integration、Release build、API snapshot compare、`git diff --check`，API baseline 保持零差异。
- 执行阶段不修改 `review.md`；完成后交由独立 Reviewer 复审。

### Phase 9 约束

- 只补测试和证据，不修改公开 API、API baseline、生产实现或依赖。
- 按维护者决策固化当前运行时行为；若文档与实现不一致，只修正文档和执行记录。
- 不新增测试框架或覆盖率包，继续使用现有 xUnit 2.4.2。
- 服务器预算保持 2C4G，逻辑并发等级 1/4/16/64 全部通过，生产最大实际并行度保持 1。
