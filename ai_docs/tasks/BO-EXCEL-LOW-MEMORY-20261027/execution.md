<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BO-EXCEL-LOW-MEMORY-20261027
AI_EXECUTION_FINISHED_AT: 2026-10-27T08:09:13Z

# 实施执行报告

## 当前修复结论（Review Round 1）

本轮按 recommended 范围完成 FIX-001～FIX-007，共 7 项 P2/SHOULD_FIX，无未完成修复项。net6.0 完整单测 1100/1100、net8.0 完整单测 1103/1103、双 TFM 流式集成各 2/2、文档消费者 12/12 全部通过。4 个受影响导出容量场景重测通过，4 个导入场景复用；详见末尾 Review 修复记录和 resource-results.md。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
PERFORMANCE_EVIDENCE_STATUS: PARTIAL（DOM 性能对照、累计分配/GC、取消延迟未验证）
RESOURCE_EVIDENCE_STATUS: PASS_FOR_APPROVED_10K_100K_MATRIX
EXTERNAL_GATE_STATUS: NOT_APPLICABLE
RELEASE_STATUS: PENDING_INDEPENDENT_REVIEW
OPEN_ACTIONABLE: 0（Executor 修复范围）

review.md 保留上轮 NEEDS_FIX 独立证据，本记录不代替下一轮独立 Review。以下原实施记录保留历史状态和数字，当前修复结果以本节及末尾 Round 1 为准。

## 原实施结论（修复前记录）

本次批准范围已完成：SXSSF 低内存导出、ZIP/XML 逐行导入、磁盘共享字符串索引、有界状态、生命周期处理、DI、文档/API 快照及本地包验证。导出作为首个里程碑完成后继续完成导入及全部验收，未把仅完成导出当作整体完成。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
PERFORMANCE_EVIDENCE_STATUS: PASS
RESOURCE_EVIDENCE_STATUS: PASS
EXTERNAL_GATE_STATUS: NOT_APPLICABLE
RELEASE_STATUS: LOCAL_VALIDATION_PASS
GOAL_STATUS: COMPLETED
OPEN_ACTIONABLE: 0

这里的 Release 状态仅表示本地包验证通过，没有发布 NuGet、推送分支或运行外部 CI。

## 任务信息

- 授权：用户明确要求 IMPLEMENT THIS PLAN，沿用 plan.md 的 Task ID；已将本次确定参数同步到计划首节。
- 开始：2026-10-27T06:00:27.501Z；工作区 HEAD：84958c1cc5eec5afbae51d76b0d07b4ab133b748。
- 遵循 execute-plan、根 AGENTS.md 和 GOAL_RULES.md；窄任务委派 luna_worker，主代理负责架构、集成及验证。

## 计划执行情况

|任务|结果|交付|
|---|---|---|
|LM-001|完成|固定 NPOI 2.7.4 SXSSF 窗口、压缩、打包与清理直接测试；默认契约/API 清单|
|LM-002|完成|新增接口、请求、行结果、统计、消费决定和独立 Sheet 快照|
|LM-003|完成|复用列计划、单元格写入、样式解析、材料化与校验；唯一值跟踪增加字符额度且保留旧构造函数|
|LM-004|完成|inline strings 多 Sheet 单次枚举导出，预检不支持配置，写入即应用样式，有限窗口刷盘，文件暂存及原子提交|
|LM-005|完成|真实 ZIP 关系、有限 XML 解析、共享字符串磁盘数据和定长索引、有界缓存、格式表、稀疏行和公式缓存|
|LM-006|完成|同步/异步逐行回调、转换/校验、背压、Stop/错误阈值/样本截断、失败行无实体|
|LM-007|完成|显式 DI、自定义服务优先、调用隔离、文件保护、真实异步外围 IO、清理失败可见且保留主异常|
|LM-008|完成|8 个容量场景、双 TFM 回归、快照/成员追溯、文档及实际包消费者|

## Change Impact Analysis

- Abstractions：新增接口/请求/选项/结果，Sheet 配置快照，UniqueTracker 字符额度，StopReason 末尾追加 ConsumerStopped。旧签名和枚举值保留。
- NPOI：独立流式实现；共享原导入材料化、表头绑定、导出列写入、样式和 ZIP 元数据预检。覆盖 XLS/HSSF、XLSX/XSSF 旧实现及 CSV 唯一值回归。
- ResourceProbe：新增显式 --streaming-scenario 模式，旧场景不变；4 种方向/数据组合各运行 10k/100k。
- 验证先职责测试，再完整单测、集成、文档、包。测试失败只按明确可修复项处理；修复后测试相应影响范围。

## 已完成事项

- 通过公开接口提供同步/异步流出口和文件出口，以及按名称/索引选择 Sheet 的逐行导入。
- 数据行/实体不累计；共享字符串超过预算后数据与索引同时落盘，查询缓存仍受原预算约束。
- 导出在创建新行前预留单行字符额度，避免先分配后超窗口预算。单元格样式在刷盘前完成。
- 输入从当前位置暂存；调用方流始终保持打开。消费者异常/取消传播；失败行 Item=null。
- 错误总数、样本截断、完成/成功和停止原因分别表达，不把样本截断当作停止。
- XML 禁用 DTD/外部解析，按真实关系找到部件；对未消费部件仅检查包元数据。
- SXSSF 正常 Dispose 后 Close 基础 XSSF；失败继续清理所有 writer；清理错误进入异常或主异常 Data。

## 部分/未完成事项

本次批准范围无未完成项。模板、图片、图表、复杂合并、公式计算及跨 Sheet 关系组装属于明确排除范围；50 万/100 万行未测且不作容量承诺。

## 修改文件

主要新增：
- src/Bing.Offices.Abstractions/Bing/Offices/Exports/IExcelStreamingExporter.cs
- src/Bing.Offices.Abstractions/Bing/Offices/Imports/IExcelStreamingImporter.cs、ExcelStreamingImport.cs
- src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelStreamingExporter.cs
- src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelStreamingImporter.cs
- src/Bing.Offices.Npoi/Bing/Offices/Streaming/ 下临时会话、XML/行/包读取器、共享字符串索引、取消写入包装器
- tests/Bing.Offices.Tests/ 下 7 个流式/唯一值/索引测试文件
- tests/Bing.Offices.Tests.Integration/ExcelStreamingIntegrationTest.cs
- tests/Bing.Offices.ResourceProbe/ExcelStreamingResourceMatrix.cs
- docs/excel/streaming.md

兼容修改：ExcelImport.BuildSnapshot、ImportStopReason、UniqueTracker、NpoiExportSheetWriter、NpoiImportRowMaterializer、NpoiImportSheetExecutor、NpoiXlsxZipPreflight、DI 扩展；API 基线/治理测试、消费者、文档及 ResourceProbe 入口。根 README 和文档保留本任务开始前修改再追加本次说明。

## API/数据/配置变化

每 TFM 增加 Abstractions 67 条及 NPOI 1 条 canonical 记录，Core 0；删除 0、旧签名变更 0。完整双 TFM 差异见 api-member-diff.json，责任测试及迁移见 api-changes.md、docs/excel/streaming.md。基线只接受用户计划内兼容增量，原基线备份于 artifacts/low-memory/api/previous-baseline.json。

默认 100 行窗口、压缩 XML、8 MiB 字符串/缓存预算、32,767/262,144/2,097,152 字符额度、100 列、32 Sheet、4096 样式、256 字体、100,000 唯一值/8,388,608 字符、1000 错误停止/100 样本、库临时文件 512 MiB。输入/ZIP/XML 沿用有限默认值。

## 测试结果

|验证|结果|
|---|---|
|完整单测 net8.0|1033/1033 PASS|
|完整单测 net6.0|1033/1033 PASS|
|最终窗口预留与旧空 ZIP 条目行为修正后的影响测试 net8.0|47/47 PASS|
|同上 net6.0|47/47 PASS|
|集成 net6.0/net8.0|各 40/40 PASS（含并发隔离和异步取消目标保护）|
|文档消费者 net8.0|12/12 PASS|
|双 TFM API 快照|PASS；0 删除|
|实际包消费者|4 个消费者组合 PASS；新流式 export=4/import=2 均通过|
|包验证负例|缺候选、身份篡改、API 不匹配 3 个均按预期拒绝|
|容量矩阵|8/8 success，所有单元格校验，临时文件残留 0|

完整单测之后增加一组文本窗口参数化覆盖，仅按变更影响运行 47 项；没有把未重跑的完整总数写成 1034。最后的生产变更不改变已验证包接口。

使用命令：
```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore -v quiet
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore -v quiet
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter "FullyQualifiedName~ExcelStreaming|FullyQualifiedName~NpoiFailureWorkbookPreflight" -v quiet
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore --filter "FullyQualifiedName~ExcelStreaming|FullyQualifiedName~NpoiFailureWorkbookPreflight" -v quiet
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release --no-restore -v quiet
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --no-restore -v quiet
./build/Verify-PackageCandidate.ps1 -OutputDirectory artifacts/low-memory/package-candidate
```

本地包日志、候选及比较结果：artifacts/low-memory/package-candidate/；其 summary.json 已复制为本任务 package-validation-summary.json。

## Build/Typecheck/Lint/Format

双 TFM Release 构建通过，生产构建 0 warning/0 error。新增生产代码无 Task.Run、Task.Result 或 Task.Wait；SummaryState.Result 是普通统计方法，不是 Task.Result。git diff --check 通过；存在仓库既有 CRLF→LF 提示，无空白错误。文本读写显式 UTF-8。

## 容量与证据复用

详见 resource-results.md 及 resources/ 下 8 份 JSON。10 万行 × 10 列导出约 2.85～3.11 秒，导入约 1.82～8.17 秒；进程峰值约 80.7～112.6 MiB。高唯一导入 10k/100k 均实际触发磁盘数据和索引。

测量包含完整 API 生命周期，进程峰值包含初始化；临时盘和托管堆为 20 ms 采样，可能漏掉极短峰值。导出首行延迟为源首次枚举，不是完整文件首字节延迟。无固定 MB 性能门槛。

容量运行后新增的窗口预留分支只在“已驻留字符 > 窗口额度 - 单行额度”时触发；固定容量样本每窗口字符远低于该阈值，不改变该数据的刷盘路径，因此复用已有容量证据，不因报告/测试编辑重复重型测量。对应新分支用 256 字符窗口的全量读回职责测试验证。恢复旧预检对空条目的处理只影响旧 DOM XML 扫描，不影响 streaming 元数据/解析路径。

## 计划偏差

- NPOI XML 文档中的 FlushRows(int,bool) 实际为非公开成员。使用公开 FlushRows()，字符额度紧张时清空整个窗口，不调用非公开 API。
- 导入只构造一行受限 XSSF 表头以复用既有绑定器；正文从不进入 DOM。符合有界表头目标，避免复制既有复杂绑定规则。
- 文档路径采用 docs/excel/streaming.md；容量使用现有 ResourceProbe 的新增模式，没有额外 Benchmark 项目或新依赖。

## 基线问题

开始时工作区已存在 CI、API 工具、文档、MappingBuilder、包消费者等其他任务修改；均保留。初次全测 4 个失败为本次新增 API 尚未登记，登记后通过。新增安全测试 2 个失败来自 fixture 未改写 .rels，修正 fixture 并增加命中断言后通过，没有据此修改生产逻辑。仓库外 net6 消费者出现 SDK 已知 NETSDK1138 提示，不影响用户要求的兼容验证。

## 已知问题与分类

- Completed：LM-001～LM-008、职责/回归/文档/包/容量验收。
- Open Actionable：无。
- Blocked Approval：无。
- Blocked External：无。
- Not Applicable：外部 CI、NuGet 发布、Git 提交、50 万/100 万行测量。
- Accepted Limitations：普通 XLSX 首版范围；同步 SXSSF 构建/打包；NPOI 系统临时文件不计入库硬配额；处理额度不等于进程内存上限。
- Verified Boundaries：错误/字符/样式/字体/Sheet/ZIP/XML/临时磁盘额度、回调背压、消费者异常、取消、已有目标保护、清理失败可观察。
- Deferred：本次范围无延后项；高级 Workbook 功能继续由原 DOM 服务提供。

## 风险与回归关注点

自定义转换器、业务消费者与调用方集合的内存不由库控制。共享字符串需先建索引，首行延迟包含该阶段。NPOI 原生同步操作不承诺立即响应取消，但库检查行/打包输出/提交边界。外部流和已完成的业务写入不回滚。未消费 ZIP 部件只预检元数据。

## Reviewer 注意事项

优先核对 api-changes.md 中最终符号→测试追溯和 resource-results.md 的测量边界。API 基线无破坏性变更；本任务不把已有未提交改动归为新增功能。不需要为了记录 NOT_APPLICABLE/ACCEPTED_LIMITATION 再开启修复循环。

## Git 状态

未自动 git add；未自动 git commit；未自动 git push；未自动创建 PR。工作区保留全部先前改动及本次交付。

## 任务收口

使用 task-finish.mjs --no-notify 完成本地状态收口，不发送外部通知。

## Review 修复记录

### Round 1（REVIEW_FIX / recommended）

- Review：NEEDS_FIX，处理 FIX-001～FIX-007；保留 review.md 为独立审查证据。
- ChangedFiles：NpoiExcelStreamingExporter、NpoiExportSheetWriter、XlsxRowReader、NpoiExcelStreamingImporter、ExcelStreamingTempSession、日期文本共享帮助器、ResourceProbe 与对应职责测试/文档。
- ChangedProjects：Bing.Offices.Npoi、Bing.Offices.Tests、Bing.Offices.ResourceProbe。
- ChangedPublicContracts：无签名变化；恢复已批准行为，不更新 API 基线。
- ChangedRuntimePaths：导入标量/错误/停止、导出动态读取异常、文件出口总配额。
- ChangedProviders：NPOI；共享帮助器验证 HSSF/XSSF。
- ChangedTFMs：net6.0/net8.0；Abstractions 仍为 netstandard2.0。
- ChangedBuildPackaging：仅 ResourceProbe 内部成员向测试程序集开放；产品依赖及打包配置不变。
- ChangedBenchmarkHarness：补充提交临时文件峰值与残留计量；只重跑受影响的 4 个导出场景。
- ChangedDocsOnly：同步额度/停止规则说明，补机器信息及未验证指标说明。
- AffectedDependents：Excel 导入/导出消费者、异常观察器、文件提交器、现有 DOM 帮助器。
- RiskLevel：MEDIUM。先新增复现职责测试并验证失败，再修复；职责测试通过后执行受影响 Provider 双 TFM 单测、流式集成/文档验证。
- 开始时 Open Actionable：7；Blocked Approval / Blocked External：无。

#### 容量重测预检（执行前）

- 仅执行 export × repeated/unique × 10,000/100,000 行 × 10 列，4 个独立进程串行场景；每格 Expected Outcome 均为 EXPECTED_SUCCESS。
- 每格进程超时 90 秒、工作集保护上限 1 GiB、观测目录临时文件保护上限 1 GiB；触发即终止并记录。它们是本地执行保护，不是新增性能或内存验收门槛。
- 数据分布及默认导出配置保持不变；先通过探针职责测试，再执行这 4 格。文件副本额度和探针覆盖范围变化使旧导出临时峰值失效，旧证据保留在原 resources/，新证据写入 resources/fix-round1/。
- 4 个导入场景复用原始证据：固定负载仅有文本，没有日期、错误单元格或停止阈值交叉；本轮改变的解析分支不被命中，输入暂存、字符串索引/缓存及逐行资源路径不变。
- API 签名、产品依赖和包布局均未改变，复用既有 API/包消费者证据；当前实现通过双 TFM 单测/API 快照、流式集成和文档消费者验证，不重复完整打包门禁。

容量执行保护修正：首次 repeated/10k 成功；repeated/100k 在完整目标文件生成且两处临时文件均清理之后被 1 GiB 进程保护终止，未产生完整 JSON，不能记为通过。源码确认探针在停止 API 计时/峰值采样后使用 XSSFWorkbook 全量读回 100 万个单元格，进程保护也覆盖了这一 DOM 校验。保留首次 run-summary，后续进程保护明确调整为 4 GiB（本机约 32 GiB 内存），90 秒和临时盘保护不变；这不是调整产品额度或性能门槛。复用已成功的 10k，100k 中止格仅补跑一次，再完成 unique 两格，不改变探针的独立全量校验方法。

#### FIX-001 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：文件出口只对 session 暂存包计费，遗漏默认提交器拥有的完整副本。
- 修改：`NpoiExcelStreamingExporter` 将四出口委派共享核心；文件入口在复制前调用 `ExcelStreamingTempSession.ReserveCommitCopy`，按暂存包长度预留完整提交额度，保持默认/自定义 `IFileExportCommitter` 和原子目标保护。调用方流不预留副本。
- 验证：`ExcelStreamingExportRegressionTest.FileExport_ShouldBudgetStagingAndCommitTogether_AndPreserveTarget` 覆盖同步/异步 × 默认/自定义提交器，夹在单包与双包之间的额度拒绝文件出口，流出口成功，足量文件出口完整读回，原目标及清理断言通过。

#### FIX-002 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：XML reader 的日期 Value 已转换，Text 仍是 Excel 序列号。
- 修改：新增内部 `NpoiScalarText.Date`，`CellExtensions.GetStringValue` 和 `XlsxRowReader.Parse` 共用既有日期格式，保留 DateTime、公式缓存、格式索引和日期系统上下文。
- 验证：`ExcelStreamingImportRegressionTest.Import_DateTextAndContexts_ShouldMatchDom` 的 16 个组合覆盖同步/异步、1900/1904、内置/自定义格式、普通/公式日期，对照 DOM 字符串、结构化日期及转换/校验上下文；`NpoiCellExtensionsTest.GetStringValue_ShouldReadAllSupportedCellKinds` 覆盖 HSSF/XSSF 共享帮助器。

#### FIX-003 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：reader 所有可定位单元格错误直接进入行错误，忽略绑定列范围且丢失上下文。
- 修改：Excel 错误作为结构化 `ExcelCellValue` 进入既有材料化器；缺缓存等局部错误保留 RawValue，并按绑定列筛选、补齐属性/键/表头/位置。XML 结构和字符额度继续检查未绑定列。未知 Excel 错误文本交付有定位的转换错误。
- 验证：`Import_UnboundCellFailures_ShouldNotRejectSelectedValues`、`Import_BoundCellFailures_ShouldRetainStructuredDiagnostics`、`Import_UnknownExcelError_ShouldDeliverBoundDiagnostic` 全部通过；忽略列仍受文本额度约束由 `Import_IgnoredColumnStillObeysTextQuota_WithSparseRowLocation` 直接覆盖。

#### FIX-004 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：把下一期望行用于当前错误定位，稀疏行时偏差更大。
- 修改：MaxRows 使用实际 `row.Index + 1`；reader 提供 `CurrentRowIndex`，行内额度异常使用正在读取的物理一基行号。
- 验证：`Import_RowLimit_ShouldLocateActualPhysicalRow` 和 `Import_IgnoredColumnStillObeysTextQuota_WithSparseRowLocation` 覆盖非默认起始行、稀疏行、MaxRows、单元格/单行字符边界、准确计数及 ResourceLimit 摘要。

#### FIX-005 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：回调后的错误/唯一值阈值先于消费者 Stop 返回。
- 修改：已交付回调返回 Stop 时先返回 ConsumerStopped，保留失败行/错误统计；Continue 仍返回相应阈值，取消和回调异常继续直接传播。
- 验证：`Import_ConsumerDecision_ShouldTakePrecedenceOverThreshold` 的 8 个组合覆盖同步/异步、Continue/Stop、错误/唯一值限额；设置后续畸形行确认不会继续预读。已有消费者异常和取消职责测试也通过。

#### FIX-006 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：流式出口直接调用 DynamicGetter，遗漏 DOM 的异常包装边界。
- 修改：提取 `NpoiExportSheetWriter.ReadDynamicValues`，DOM 与 SXSSF 共用；普通用户异常包装为含 Sheet/一基行号/属性/InnerException 的 UserExtensionFailed，已有公共异常和取消原样保留，每行 getter 只读一次。
- 验证：`Export_DynamicFailure_ShouldPreserveContextAndNotifyOnce` 覆盖四出口 × 普通/公共/取消异常，精确断言观察一次、上下文、清理和原目标；`DomExport_SharedDynamicReader_ShouldPreserveBothProviders` 覆盖 HSSF/XSSF 成功及失败。

#### FIX-007 — COMPLETED

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX。
- 根因：探针漏查目标旁提交文件；20ms 定时采样也可能遗漏短复制窗口。
- 修改：`ExcelStreamingResourceMatrix` 加入本次目标提交文件跟踪、session/commit/总量及残留分项；排除原目标、预存及其他名称文件。仅探针装饰真实默认提交器，每次实际 Write/WriteAsync 完成后采样，以活动流 Length 补足缓冲可见性，不额外 Flush。报告新增 OS/架构/处理器信息，resource-results.md 记录机器/磁盘及未验证指标。
- 验证：`ExcelStreamingResourceObservationTest.ResourceSampler_IncludesOnlyCurrentTargetCommitFile_WithSessionBytes` 精确断言 11 + 17 = 28 字节并检出残留；`FileExportCommitter_SamplesTheSessionAndCommitCopyOnEveryWrite` 同步/异步真实出口完整计量双副本。3/3 PASS。仅重测 4 个导出场景，均 success、全部单元格校验通过、双副本重叠已观察、所有临时文件残留 0；原 JSON 和中止记录均保留。

#### Round 1 验证结果

|验证|结果与证据|
|---|---|
|修复前职责复现|导入 44 项中 40 项按预期失败；导出 20 项中 8 项按预期失败。TRX：import-before-fix、export-before-fix|
|导入最终专项|46/46 PASS，import-final.trx|
|导出/生命周期专项|34/34 PASS，export-after-fix.trx；包含当时 1 项探针范围测试|
|最终探针专项|3/3 PASS，resource-observation.trx|
|最终完整单测 net6.0|1100/1100 PASS，unit-net6.trx；含双 Provider 共享帮助器与 API 快照|
|最终完整单测 net8.0|1103/1103 PASS，unit-net8.trx；额外 3 项为 net8-only ResourceProbe 测试|
|流式集成 net6.0/net8.0|各 2/2 PASS，并发隔离及跨窗口取消目标保护；integration_*.trx|
|文档消费者 net8.0|12/12 PASS，docs.trx|
|Release 构建/静态检查|测试构建成功；无新增 warning/error；git diff --check PASS；生产无 Task.Run、Task.Result、Task.Wait 伪异步|
|容量导出矩阵|4/4 success；100k API 约 3.084～3.256 秒，进程峰值约 96.2～96.3 MiB，库临时总峰值约 5.89～7.06 MiB|

本轮 TRX 和执行日志在 `artifacts/low-memory/fix-round1/`。当前资源原始 JSON 在 `resources/fix-round1/`。最终生产符号→测试方法映射及迁移补充已更新到 api-changes.md；使用规则更新到 docs/excel/streaming.md。

主要验证命令（均 Release / --no-restore）：

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release --no-restore --filter "FullyQualifiedName~ExcelStreamingIntegrationTest"
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --no-restore
pwsh -NoProfile -File artifacts/low-memory/fix-round1/run-export-matrix.ps1
```

#### Round 1 汇总与停止条件

- MUST_FIX：0；SHOULD_FIX：7；已完成：7；PARTIAL/BLOCKED/FAILED 修复项：0。
- Completed：FIX-001～FIX-007、必要职责/回归、当前导出容量证据、文档/追溯。
- Open Actionable：0；Blocked Approval / Blocked External：无。
- Not Verified：DOM 性能对照、累计分配/GC、取消释放延迟；按 FIX-007 允许的证据说明处理，不由现有数据推导，也不为此自动开启下一轮修复。
- Accepted Limitations：同步 NPOI 内部阶段、NPOI 临时文件不计库硬额度、外部流及业务消费副作用不回滚，保持原约定。
- Not Applicable：外部 CI、发布、50 万/100 万行以及高级 DOM 功能。
- 下一步：独立 Review；本轮 COMPLETED 只表示修复执行完成，没有修改 review.md 的 NEEDS_FIX。
- 未执行 git add / commit / push / PR；原工作区修改全部保留。收口使用 task-finish.mjs --no-notify，不发送外部通知。
- 收口结果：`ok=true`、`finalStatus=COMPLETED`、`mode=review-fix`、`reviewRound=1`、`notificationStatus=skipped`。
