<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260831-001
AI_EXECUTION_FINISHED_AT: 2026-09-01T15:23:43.3424781+08:00

# 实施执行报告

## Review Fix Round 5 收口

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- 执行终态：`PARTIAL`
- 独立复审文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 说明：Round 5 runtime 曾保持 `IN_PROGRESS`，但没有形成新的 Executor 完成记录。独立 Reviewer 已完成源码、Git Diff、测试和证据复验，确认本轮不能记为 `COMPLETED`，因此按实际结果收口为 `PARTIAL`。

### Round 5 FIX 状态

- `FIX-001`：`NOT_RESOLVED`。9 个 lockfile 和当前任务证据仍未进入 Git 交付面，任务 benchmark 证据还受到 `.gitignore` 规则影响；dirty 工作树 locked restore 通过不能替代干净交付验证。
- `FIX-003`：`REGRESSED`。net6/net8 Unit 均为 `371/372`，失败于 API member snapshot；四 TFM 静态 snapshot 仍没有自动生成/比较入口。
- `FIX-004`：`NOT_RESOLVED`。性能结果仍为 `UNAPPROVED`，没有批准预算、批准人或具名 waiver，原始证据未进入 Git。
- `FIX-006`：`PARTIAL`。既有 selector 定向测试 net6/net8 各 `9/9` 通过；越界 `ByIndex`、`VeryHidden` 和客户端互操作/waiver 仍缺失。
- `FIX-007`：`REGRESSION`。Docs consumer 改为生产 `ProjectReference`，当前 `11/11` PASS 不再证明 package-only nupkg 消费。

### Round 5 验证与结论

- Release solution build：PASS，0 errors。
- Unit net6/net8：FAIL，各 `371/372`，API snapshot 失败。
- Integration net6/net8：PASS，各 `15/15`。
- Docs tests net8：PASS，`11/11`，但 package-only 门禁失效。
- 发布结论：`No-Go`。
- 下一步：由 `review-fixer` 按最新 `review.md` 处理 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`、`FIX-007`，完成后重新独立 Review。
- 未执行 Git staging、commit、push、tag、PR、publish 或外部通知。

> Round 2 Review Fix 已完成可执行修复；最终发布结论仍为 `PARTIAL/No-Go`，等待后续独立 Reviewer 验收。

## 当前执行摘要

- 已读取并按 `plan.md` 执行；未修改 `plan.md`，未执行 Git 提交、推送、Tag、PR 或 NuGet 发布。
- P0 selector 已修复：工作簿打开后一次解析请求 selector，结果复用于 plan、materialization 和 Failure Workbook mapping。
- 当前验证：Round 1 的 Release build、net6/net8 Unit、Integration、Docs consumer、资源探针和本地 pack 证据保留；Round 2 新增 net6/net8 API 契约测试均 `7/7`，UniqueJournal 四场景正式 DefaultJob 完成并已归档原始产物；本轮补充 net6/net8 metadata 与 selector 矩阵，并在隔离 NuGet 缓存中完成 Docs package consumer metadata 重开测试。
- 已同步 README、Excel 文档和 NuGet migration；已生成任务要求的 00-11 报告文件。
- 当前发布判定为 **No-Go**，任务保持 `PARTIAL` 终态。限制项见 `10-final-review.md` 和 `11-final-summary.md`：锁文件尚未进入 Git index、部分 API/性能证据仍受执行范围限制、办公客户端互操作 `NOT_VERIFIABLE`、RH31-105 完整追溯矩阵未完成。
- 已完成证据：Release build；net6/net8 Unit 各 `360/360`；net6/net8 Integration 各 `15/15`；隔离缓存 Docs consumer `9/9`；资源探针 `16/16`；三个本地 nupkg pack 成功。
- 未执行 Git commit、push、tag、PR 或 NuGet publish。

## Review 修复记录

### Round 2

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`（处理 `MUST_FIX + SHOULD_FIX`）
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 说明：本轮未修改 `review.md`，其 Reviewer 证据和 `NEEDS_FIX` 状态保持不变。

#### FIX-001

- 严重程度：BLOCKER
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 已完成：保留 `.gitignore` 对 10 个 lockfile 的显式例外；确认 10 个 lockfile 均存在且不再被忽略；`dotnet restore Bing.Offices.sln --locked-mode -v:q` 通过。
- 未完成：lockfile 仍是工作树未跟踪文件。当前用户约束禁止 `git add`，因此无法在本轮把它们纳入 Git index，也无法提供干净克隆的 locked restore 证据；未宣称该项完成。

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED（Round 1 已独立复审为 resolved）
- Round 2 未改变已通过的资源、所有权和原始异常类型修复。

#### FIX-003

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED（可执行范围）
- 修改文件：`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/04-api-breaking-changes.md`。
- 修复：增加逐类型显式分类 manifest；实际导出类型缺失或多余时失败；逐个枚举公开构造函数、属性、字段和方法并回溯到已分类类型；成员快照由 `net8.0` 扩展为 `net6.0` 与 `net8.0`，两目标使用实际程序集输出哈希。
- 验证：`dotnet test ... -f net8.0 --filter "FullyQualifiedName~PublicApiContractTest"`：`7/7`；`dotnet test ... -f net6.0 --filter "FullyQualifiedName~PublicApiContractTest"`：`7/7`。
- 限制：netcoreapp3.1/net7.0 本轮未执行，不据此宣称全部 shipped TFM 已验收。

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL
- 修改/产物：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/08-benchmark-report.md`、任务目录 `benchmarks/` 下六组 BDN 原始 CSV/JSON/Markdown/HTML 产物。
- 修复：完成 `UniqueJournalBenchmarks` 的 1/5 唯一列、10K/100K 行四个正式 DefaultJob 场景；结果为 `3.122 ms`、`52.252 ms`、`26.247 ms`、`313.213 ms`，并保留分配、GC、置信区间等原始数据；将六组 benchmark 产物复制到非忽略任务证据路径。
- 验证：`dotnet run -c Release --project benchmarks/Bing.Offices.Benchmarks -- --filter "*UniqueJournalBenchmarks*" --exporters json markdown`：4/4 场景完成，退出码 0。
- 未完成：没有批准的历史 baseline/预算，未测并发 1/4/16/64 tail latency；因此仍为 `PARTIAL`，不宣称性能回归通过。

#### FIX-005

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED（可执行范围）
- 修改文件：`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`。
- 修复：将 64 路并发导出测试改为每个请求重新打开 XLS/XLSX，并分别断言 Author、Company、Title、Subject、Category、Description/Comments 六个 metadata 字段；期望值独立保存，避免测试自比较。
- 验证：net8.0 和 net6.0 定向测试均通过。
- 残余限制：未将并发测试扩展到 netcoreapp3.1/net7.0，本轮不据此宣称全部 Unit TFM 均已复验。

#### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：`tests/Bing.Offices.Tests/StreamPipelineTest.cs`、`tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`、`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`。
- 修复：增加 XLS/XLSX 按名称 selector、Ordinal/OrdinalIgnoreCase、缺失/隐藏 Sheet、混合名称与索引、同物理 Sheet 冲突，以及按索引失败工作簿实际 Sheet 名称回链测试；Docs consumer 在隔离 NuGet 缓存中通过 `AddNpoi()` 导出并重开 XLS/XLSX，验证六个 metadata 字段。
- 验证：核心 selector/metadata 定向测试 net8.0、net6.0 通过；Docs consumer metadata 测试 net8.0 通过；NPOI 仅新增为 Docs 测试项目依赖，未进入生产公开 API。
- Office 互操作：当前环境未发现 `excel`、`soffice` 或 `libreoffice`，因此记录为 `NOT_VERIFIABLE`，未宣称客户端打开/保存/重开通过。

### Round 2 汇总

- `FIX-001`：PARTIAL/BLOCKED，锁文件存在且 locked restore 通过，但未进入 Git index。
- `FIX-002`：COMPLETED，沿用 Round 1 独立复审结果。
- `FIX-003`：COMPLETED（net6/net8 可执行门禁）；未执行 netcoreapp3.1/net7.0。
- `FIX-004`：PARTIAL，UniqueJournal 和非忽略原始证据已补齐；预算和并发尾延迟仍缺。
- `FIX-005`：COMPLETED（net6/net8 可执行范围）；`FIX-006`：PARTIAL，代码与 package consumer 证据完成，Office 客户端互操作 `NOT_VERIFIABLE`。
- 发布结论：`No-Go`。不得修改 `review.md` 或自行将 Reviewer 状态改为 PASS。

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 说明：Round 1 按当时注册的 `must` 范围只处理 `MUST_FIX`；`review.md` 未被修改。Round 2 已按 `recommended` 范围补处理 `FIX-005` 和 `FIX-006`。

#### FIX-001

- 严重程度：BLOCKER
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `.gitignore`
	- `build/packages.lock.json`
	- `src/Bing.Offices.Abstractions/packages.lock.json`
	- `src/Bing.Offices.Core/packages.lock.json`
	- `src/Bing.Offices.Npoi/packages.lock.json`
	- `tests/Bing.Offices.Docs.Tests/packages.lock.json`
	- `tests/Bing.Offices.ProfileFixtures/packages.lock.json`
	- `tests/Bing.Offices.Tests/packages.lock.json`
	- `tests/Bing.Offices.Tests.Integration/packages.lock.json`
	- `benchmarks/Bing.Offices.Benchmarks/packages.lock.json`
- 根因：锁文件原先被通用 `packages.lock.json` 忽略，且 lock graph 的 requested 范围与当前 PackageReference 表达不一致，导致 `NU1004`。
- 修复：使用 `--force-evaluate -p:RestoreLockedMode=false` 重新生成当前依赖图；在 `.gitignore` 中仅为解决方案项目和 build/benchmark 项目增加明确 negation 例外；没有关闭 locked mode。
- 验证：
	- `dotnet restore Bing.Offices.sln --locked-mode -v:q`：PASS
	- 所有 10 个锁文件在工作树中可见为未跟踪交付文件，未执行 `git add`：PASS
- 残余限制：锁文件尚未提交到 Git，需由后续交付流程纳入版本控制后才具备干净克隆可复现性。

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/StreamPipelineTest.cs`
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/05-unit-test-report.md`
- 根因：已有资源、释放、取消和异常测试分散在多个测试文件，缺少 provider-main-chain 的直接追溯证据。
- 修复：新增 `Import_NonSeekableInputOverLimit_ShouldRejectAndKeepSourceOpen` 和 `Import_MappingFactoryFailure_ShouldPreserveOriginalExceptionType`；补充 `NpoiStreamCopier`、Excel/CSV importer/exporter、parser/writer、Failure Workbook、AtomicFileCommitter、mapping plan、API 和 ResourceProbe 的生产符号到测试方法矩阵。
- 验证：
	- net8.0 两个新增测试：2/2 PASS
	- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore -v:q`：PASS
	- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore -v:q`：PASS
	- 矩阵报告已写入 `05-unit-test-report.md`：PASS
- 残余限制：Office 客户端互操作仍为 `NOT_VERIFIABLE`；selector/consumer 扩展矩阵已在 Round 2 补齐。

#### FIX-003

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/04-api-breaking-changes.md`
- 根因：原 API 基线只记录当前导出面，没有逐组分类、迁移决策和分类自动门禁。
- 修复：建立 `User API`、`Provider SPI`、`Compatibility`、`Execution detail` 四类符号账；记录当前可见性、消费者/测试证据、source/binary 影响、版本迁移策略和批准状态；新增 `PublicApi_ExportedTypes_ShouldHaveGovernedClassification`，并保留 NPOI 泄漏、生产 IVT、NPOI 精确成员和 net8 成员哈希门禁。
- 验证：
	- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter "FullyQualifiedName~PublicApiContractTest"`：7/7 PASS
	- `get_errors` 检查触及 API 测试文件：无错误
	- Release 构建：PASS
- 残余限制：本轮未执行已发布每个 TFM 的独立 API 文件 diff；当前跨 TFM 仍由测试目标矩阵和 net8 完整成员哈希门禁覆盖，后续需在版本负责人批准 breaking 收敛前补充各 shipped TFM 的独立导出快照。

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`
	- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/08-benchmark-report.md`
- 根因：原 benchmark 使用固定三次迭代 smoke 作业，不能作为 1K/10K/100K release evidence。
- 修复：移除所有 `SimpleJob(1,2,3)`，改用 BenchmarkDotNet `DefaultJob`；正式运行 Stream Pipeline、Failure Workbook、MappingValidation、DynamicPlan 和 TenantPlanCache，保留 CSV、JSON、Markdown、HTML 原始产物；报告绑定环境、基线 HEAD、当前工作树 diff、分配/GC 结果和资源探针限制。
- 验证：
	- Stream Pipeline 1K/10K/100K Import/Export/ExportDestinationCapacity：PASS，原始产物位于 `BenchmarkDotNet.Artifacts/results/`
	- Failure Workbook 1K/10K/100K：PASS，原始产物位于 `BenchmarkDotNet.Artifacts/results/`
	- MappingValidation、DynamicPlan、TenantPlanCache：PASS，原始产物位于 `BenchmarkDotNet.Artifacts/results/`
	- ResourceProbe 16 场景：PASS
	- `git diff --check`：无 whitespace error
- 残余限制：`UniqueJournalBenchmarks` 本轮未形成最终摘要文件；没有批准的历史性能预算/基线，不能将单机结果判定为回归通过；并发 1/4/16/64 tail latency 仍不可验证。因此报告和本执行记录保持 `PARTIAL`，不宣称 zero-GC、真实 streaming 或完整压缩炸弹防护。

### Round 1 汇总

- MUST_FIX：`FIX-001` COMPLETED；`FIX-002` COMPLETED；`FIX-003` PARTIAL；`FIX-004` PARTIAL。
- 已完成：严格锁定还原、资源/异常直接回归测试、符号级测试矩阵、API 分类账与分类门禁、主要正式 Benchmark 产物。
- PARTIAL：独立多 TFM API 文件 diff、UniqueJournal 最终摘要、批准性能预算、并发尾延迟。
- BLOCKED：无新的执行命令阻塞；发布结论仍受上述 PARTIAL 证据限制。
- 回归验证：net6/net8 Unit 全量通过；Release 解决方案构建通过；API 定向测试 7/7 通过；严格 locked restore 通过；`git diff --check` 无空白错误。
- 下一步：保持 `review.md` 的 `NEEDS_FIX` 不变，交还下一轮独立 Review；不得由 Executor 自行改写为 PASS。

### Round 6

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 说明：本轮只处理用户指定的 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-007`，明确跳过 `FIX-006`；未修改 `review.md`、`plan.md` 或批准 API baseline，未执行 Git staging、commit、push、PR、发布或通知。

#### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修复：保留 lockfile 的显式 `.gitignore` 例外，并移除任务 benchmark 原始证据的错误 ignore 规则；当前 9 个有效 lockfile 和任务证据仍未进入 Git index。
- 验证：当前工作树 locked restore 既有证据保持有效；本轮未执行 `git add`，因此不能声称 clean clone 或 `git ls-files` 交付验证完成。
- 残余限制：需要获授权的 Git 交付步骤纳管 lockfile、报告、快照和原始 benchmark 证据后，才能完成隔离 worktree/clean clone 验收。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修改文件：`build/ApiSnapshot/ApiSnapshot.csproj`、`build/ApiSnapshot/Program.cs`、`build/ApiSnapshot/PublicApiSnapshot.cs`、`tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、四份 `api-diff-*.md`、`build/api-snapshot-baseline.json`（仅读取/保留批准值）。
- 修复：新增独立 net8 API snapshot executable；使用共享 metadata-only canonicalizer；对 `netcoreapp3.1`、`net6.0`、`net7.0`、`net8.0` 自动生成并比较；纳入 public/protected/protected-internal governed members；删除临时 legacy hash 入口；resolver 仅探测声明的第三方依赖包并排除运行时 System 程序集候选。
- 验证：工具可执行并生成四份 JSON 快照，成员数为 Abstractions `737`、Core `273`、Npoi `1`；Npoi hash `A0DBE980...` 匹配批准值；Abstractions 实际 `5A1B668E...`、Core 实际 `5F68499B...` 与批准值不匹配。net8/net6 `PublicApiContractTest` 各 `6/7`，唯一失败为批准 hash 断言。
- 残余限制：批准 hash 对应的历史产物/生成上下文无法由当前 Release DLL 复现；不得替换 baseline，需版本负责人确认产物和规范后再收敛。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修复：保留并复核已有正式 BenchmarkDotNet 和 tail-latency 原始 JSONL；确认所有性能证据仍为 `budgetStatus=UNAPPROVED`，未伪造预算、批准人或 waiver。
- 验证：Round 4 两份 tail-latency JSONL 可读取，覆盖并发 `1/4/16/64`、预热、5 轮和每档 `1280` 样本；报告保留高并发波动和未批准状态。
- 残余限制：没有版本负责人批准的 baseline、p99/吞吐预算、容差、批准人和日期或具名 waiver；本项不能判定性能 PASS。

#### FIX-007

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`、`tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`、`07-package-consumer-report.md`。
- 修复：恢复三个精确 `2.0.0` Bing `PackageReference`，移除生产 `ProjectReference`；保留 metadata XLS/XLSX、DI、mapping、CSV 和文档 fence 消费测试；明确本地 Bing feed 与 nuget.org 第三方依赖源。
- 验证：全新隔离 NuGet 缓存 restore/build/test 成功；Docs package-only consumer `11/11`；`project.assets.json` 中三个 Bing 包均为 `type=package`、版本 `2.0.0`；三个本地 nupkg 均含 DLL、XML、nuspec。

### Round 6 汇总

- MUST_FIX：`FIX-001`、`FIX-003`、`FIX-004`、`FIX-007` 均已按当前可执行范围处理。
- 已完成：API 四 TFM 自动生成/比较入口；共享 canonicalizer；package-only Docs consumer 隔离 restore/build/test；报告证据同步。
- PARTIAL：`FIX-001` Git 交付/clean clone；`FIX-003` approved hash 与当前产物不一致、net6/net8 runtime gate 各 `6/7`；`FIX-004` 性能预算/waiver。
- COMPLETED：`FIX-007` package-only consumer 恢复并通过 `11/11`。
- 跳过：`FIX-006`，本轮 fixScope 为 `must`，按用户要求不处理 SHOULD_FIX。
- 回归验证：API snapshot 工具编译成功；Docs package-only consumer `11/11`；net6/net8 API 定向测试各 `6/7`，失败仅为批准 hash 不匹配；changed-file `get_errors` 无错误。
- 发布结论：`No-Go`；Executor 本轮终态为 `PARTIAL`，下一步交还独立 `code-reviewer` 再次验收。

## Review 修复记录

### Round 3

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 说明：本轮只处理 `MUST_FIX`，未修改 `review.md`，未执行 Git staging、commit、push、PR 或发布。

#### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修复：使用当前项目评估结果重新生成 9 个项目锁文件；确认 Docs consumer 的 `NPOI 2.7.4` 已写入 `tests/Bing.Offices.Docs.Tests/packages.lock.json`；保留 `.gitignore` 的显式例外。
- 验证：`dotnet restore Bing.Offices.sln --force-evaluate -p:RestoreLockedMode=false -v:q` 和随后 `dotnet restore Bing.Offices.sln --locked-mode -v:q` 均通过。
- 残余限制：锁文件仍为工作树未跟踪文件，因本轮禁止 `git add`，不能声称已进入 Git 交付面或完成干净克隆验证。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修复：将报告中定义的 Excel/CSV 主入口、请求、结果、错误、资源、metadata、图表、样式和接口类型统一为 `User API`；API 清单继续对导出类型执行 exact lookup、对未知/缺失类型失败，并保留 NPOI 泄漏、生产 IVT、Provider SPI 隐藏和多 TFM 快照门禁。
- 验证：`PublicApiContractTest` 源文件 `get_errors` 无错误；API 定向测试的 net8.0 当前通过 `7/7`。
- TFM 证据：net6/net8 为测试 host 运行证据；net7/netcoreapp3.1 已有 Release 目标程序集可做静态输出核对，但本机缺 runtime，未宣称运行 PASS。
- 残余限制：逐成员独立分类、影响和迁移 ledger 仍未入库；本轮未伪造该证据，也未修改 public API。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修复：Benchmark 入口新增固定 1/4/16/64 并发尾延迟矩阵，输出每档 256 个操作的 p50/p95/p99、墙钟吞吐和 `budgetStatus=UNAPPROVED`；报告补充可重跑命令和预算限制。
- 验证：Benchmark 项目源码通过错误检查；未运行无界大规模作业。
- 残余限制：没有具名批准的预算、历史 baseline 或 release waiver，原始 Benchmark 与任务报告仍未进入 Git index，因此不能判定性能验收 PASS。

### Round 3 汇总

- `FIX-001`：锁图已修复并通过严格还原；Git 交付面受禁止 staging 约束，`PARTIAL`。
- `FIX-003`：分类事实已对齐，net6/net8 门禁通过，缺 runtime TFM 仅有静态输出证据，逐成员 ledger 未完成，`PARTIAL`。
- `FIX-004`：补充 1/4/16/64 并发测量入口和报告；预算/waiver/交付跟踪仍缺，`PARTIAL`。
- 跳过 `FIX-006`：本轮 fix scope 为 `must`，不处理 SHOULD_FIX。
- 发布结论：`No-Go`；任务终态保持 `PARTIAL`，交还下一轮独立 Review。

### Round 4

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- `review.md` 未修改；本轮未执行 Git staging、commit、push、PR 或发布。

#### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 根因与现状：当前工作树中的 lockfile 已生成，`dotnet restore Bing.Offices.sln --locked-mode -v:q` 通过；但 lockfile、任务报告和 Benchmark 原始证据仍未进入 Git 交付面。
- 限制：本轮安全约束禁止 `git add`，因此无法完成 `git ls-files` 和干净 clone 的 locked restore 验证；未将当前工作树 PASS 等同于交付 PASS。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修改文件：`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/04-api-breaking-changes.md`。
- 修复：保留逐类型分类和逐 public/protected/protected-internal member 的自动治理记录；保留 `NETCOREAPP3_1`、`NET6_0`、`NET7_0`、`NET8_0`、`NET5_0` 显式快照分支；新增四个 shipped NPOI TFM 的独立静态 API snapshot/diff 文件；清理本轮临时 AssemblyLoadContext、路径和哈希诊断代码。
- 验证：标准 net8 Release 输出的 `PublicApiContractTest` 为 `7/7 PASS`；net6/netcoreapp3.1/net7.0 的目标程序集已完成静态快照核对，其中本机缺少 netcoreapp3.1/net7.0 runtime；net6 VSTest 运行仍返回 `3D3AA5...`，而独立 net6 进程对同一文件和同一私有哈希函数返回批准值 `7B0BA...` 并通过完整方法调用，表明测试宿主存在运行上下文差异。
- 残余限制：不能把 net6 VSTest runtime gate 记为 PASS，也不能修改 expected hash 规避失败；FIX-003 保持 `PARTIAL`，等待后续 Reviewer/运行环境复验。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修复：Benchmark 入口已包含预热、start gate、队列提交到 plan 完成的端到端计时、5 轮重复、每档 1280 个正式样本、1/4/16/64 并发、吞吐、p50/p95/p99、worker 启动耗时和环境身份；Round 4 已归档两次 JSONL 运行结果。
- 验证：`benchmarks/release-hardening-tail-latency-round4-a.jsonl` 和 `benchmarks/release-hardening-tail-latency-round4-b.jsonl` 均记录 .NET 8.0.27、Windows 10.0.19045、24 logical processors、X64、workstation GC 及 `budgetStatus=UNAPPROVED`；高并发波动已原样保留。
- 残余限制：没有具名批准的历史 baseline、环境预算或 release waiver，Benchmark 不能判定性能验收 PASS；原始证据仍未进入 Git index。

#### FIX-006

- 处理要求：`SHOULD_FIX`
- 执行状态：`SKIPPED/OUT_OF_SCOPE`
- 原因：本轮 fixScope 为 `must`，按用户要求只处理 MUST_FIX；不修改本项的索引越界、VeryHidden 和 Office 互操作证据。

### Round 4 汇总

- MUST_FIX：`FIX-001`、`FIX-003`、`FIX-004` 均已执行必要修复和验证，但分别保持 `PARTIAL/BLOCKED`、`PARTIAL`、`PARTIAL`，未宣称全部完成。
- 已完成：API 分类和 member governance 自动检查、四 TFM 静态 snapshot、Benchmark 端到端重复测量方法、当前工作树 strict locked restore。
- PARTIAL：Git 交付面/clean clone、net6 API runtime gate、性能预算和 waiver。
- BLOCKED：`FIX-001` 的 Git tracked/clean clone 验证受禁止 staging 约束。
- FAILED：无业务实现测试失败；存在 net6 VSTest API 快照宿主差异，作为 FIX-003 的未闭环证据保留。
- 回归验证：标准 net8 API 契约 `7/7 PASS`；独立 net6 反射复验 `PASS`；`get_errors` 无错误；`git diff --check` 无空白错误，仅有 CRLF/LF 提示。
- 发布结论：`No-Go`；执行器已完成本轮可执行工作，下一步交还独立 `code-reviewer` 再次验收。

## Review 修复 Round 8

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 说明：本轮仅处理 `MUST_FIX`；未修改 `review.md`、批准 API baseline 或其它 Reviewer 证据，未执行 Git staging、commit、push、PR、发布或通知。

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修复与复验：保留并同步当前项目 lockfile；先执行 `dotnet restore Bing.Offices.sln --force-evaluate -p:RestoreLockedMode=false`，随后 `dotnet restore Bing.Offices.sln --locked-mode` 通过。由于当前约束禁止 `git add`，`git ls-files -- '*packages.lock.json'` 和任务证据跟踪结果仍为空，无法完成干净交付提交/隔离 worktree 的 Git 复现证明。
- 结果：当前 dirty 工作树的严格 locked restore 已恢复，但 Git 交付面仍未闭环，不能将本项标记为完成。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修复与复验：保留共享 metadata-only canonicalizer、四 TFM API snapshot 工具和批准 baseline 原值；修复 Unit 项目文件后，net6/net8 API contract 均可执行，但各为 `6/7`，唯一失败为 Abstractions hash `actual=5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646` 不等于批准值 `7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`。Core actual `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7` 也不等于批准值 `41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`。
- 验证：API 工具 Release build 通过；以 `output/release` 为根运行四 TFM compare，`netcoreapp3.1`、`net6.0`、`net7.0`、`net8.0` 均报告 Abstractions/Core mismatch，Npoi `A0DBE980...` 与批准值一致并返回退出码 `1`。
- 结果：未修改 expected hash/baseline，也未伪造 API PASS；需版本负责人确认批准产物、规范和 migration 后才能继续收敛。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修复与复验：复核现有 BDN、资源探针和 tail-latency 证据；报告与 JSONL 仍为 `budgetStatus=UNAPPROVED`。未发现批准的 baseline、p99/吞吐预算、容差、批准人、日期或具名 release waiver。
- 结果：保留原始测量和高并发波动，不伪造性能批准或 PASS；由于没有授权 Git staging，原始性能证据也尚未进入 Git index。

### FIX-007

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修复：重新使用当前源码本地打包 `Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi` `2.0.0` 到 `artifacts/packages-vnext`；确认新 Abstractions nupkg 包含 `ExcelWorkbookMetadataOptions` 和 `ExcelWorkbookExportBuilder.Metadata`，未恢复生产 `ProjectReference`。
- 验证：Docs package-only consumer 在全新缓存中使用项目声明的本地 Bing feed 和 nuget.org 依赖源还原、构建、测试均通过；`11/11`；`project.assets.json` 三个 Bing 包均为精确 `2.0.0` 且 `type=package`。此前旧全局缓存导致的 metadata 编译失败已定位为缓存包内容过期，不是当前 nupkg 内容缺失。

### FIX-008

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修复：恢复 `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` 为唯一合法 `<Project>` 根节点，保留 net6/net7/net8/net5/netcoreapp3.1 目标、`System.Reflection.MetadataLoadContext`、测试依赖和 `PublicApiSnapshot.cs` link；文件长度为 `2849` 字节，UTF-8 XML parse 的根节点数量为 `1`。
- 验证：`get_errors` 无错误；force-evaluate 后 `dotnet restore Bing.Offices.sln --locked-mode` 通过；统一全新 NuGet 缓存下 Release solution build 通过；Unit net6/net8 均可执行并得到 `371/372`，唯一失败为 FIX-003 API baseline mismatch；selector net6/net8 各 `90/90`；Integration net6/net8 各 `15/15`。

### FIX-006

- 处理要求：`SHOULD_FIX`
- 执行状态：`SKIPPED/OUT_OF_SCOPE`
- 原因：本轮 fixScope 为 `must`，按用户要求只处理 MUST_FIX；未修改 VeryHidden、越界 `ByIndex` 或 Office 客户端互操作材料。

### Round 8 汇总

- `FIX-001`：`PARTIAL/BLOCKED`；当前工作树 locked restore 通过，但 lockfile、API 工具、报告和原始性能证据仍未被 Git 跟踪，禁止 staging 导致 clean-clone 证据无法完成。
- `FIX-003`：`PARTIAL`；Unit/API runtime gate 已恢复可执行，net6/net8 各 `6/7`，四 TFM 自动 compare 对 Abstractions/Core 保持真实 mismatch，Npoi 匹配；未修改批准 baseline。
- `FIX-004`：`PARTIAL/BLOCKED`；性能测量证据保留，但批准预算、批准人、日期和 waiver 仍缺失，状态继续为 `UNAPPROVED`。
- `FIX-007`：`COMPLETED`；当前源码重新打包后，package-only Docs consumer 隔离 restore/build/test `11/11`，三包均为 `type=package`。
- `FIX-008`：`COMPLETED`；Unit csproj 已恢复，solution restore/build 和 Unit/API/selector runtime gate 均可执行。
- 回归验证：Release solution build PASS；Docs package-only `11/11`；selector net6/net8 各 `90/90`；Integration net6/net8 各 `15/15`；Unit net6/net8 各 `371/372`，唯一失败为批准 API hash；`get_errors` 无错误；`git diff --check` 退出码 `0`，仅有 CRLF/LF 提示。
- 发布结论：`No-Go`，执行器终态为 `PARTIAL`；`review.md` 保持 `NEEDS_FIX`，下一步交还独立 `code-reviewer` 再次验收。

## Review 修复 Round 9

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 执行终态：`PARTIAL`
- 说明：本轮仅处理 `MUST_FIX`。未修改 `review.md`、`plan.md` 或批准 API baseline；未执行 Git staging、commit、push、tag、PR、publish 或外部通知。三个 MUST_FIX 均完成了当前权限范围内的复核和可执行准备，但外部批准/交付权限缺失，不能合法记为 `COMPLETED`。

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：无业务代码修改；沿用现有 `.gitignore` lockfile 例外和 package-only consumer 配置。
- 根因复核：当前 9 个有效 `packages.lock.json`、`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、任务报告/快照/benchmark 原始证据均存在于工作树，但仍未被 Git 跟踪；当前执行约束和安全策略不允许本轮擅自执行 `git add`。
- 包输入复核：`artifacts/packages-vnext` 中三个 `2.0.0` nupkg 均存在；当前 package consumer 通过项目声明的本地 Bing feed 加第三方 nuget.org source 在隔离缓存中工作。默认全局缓存中的同版本 `Bing.Offices.Abstractions.2.0.0.nupkg` 与本地包内容不同，已导致 `NU1403`，因此不能把默认缓存状态作为复现证据，也不能覆盖已消费的同版本包。
- 已完成修复准备：确认所有要求文件均未命中 `.gitignore`（除顶层 `artifacts/` 之外的任务证据路径）；确认 Docs consumer 无生产 `ProjectReference`，使用三个精确 `2.0.0` `PackageReference`；保留隔离 `NUGET_PACKAGES` 和显式第三方 source 的复现方式。
- 未完成与原因：不能在无用户授权的情况下纳管文件到 Git index，也不能创建 clean clone/实际待交付提交验证；当前 package 版本仍是可变同版本本地输入，无法由执行器批准其发布身份。
- 验证：当前工作树的隔离 locked restore、Release solution build、Docs package-only consumer `11/11` 已有 PASS；本轮重新核对文件存在/ignore/track 状态和三个 nupkg SHA-256。`git ls-files` 对要求文件仍无输出，因此本项保持 `PARTIAL/BLOCKED`。
- 后续条件：获授权的交付步骤纳管 lockfile、API 工具/baseline、任务报告、快照和需保留的原始性能证据；固定本地 Bing source 与第三方 source/预热 cache；禁止覆盖已消费的 `2.0.0`，改用不可变包内容或唯一版本身份；在实际待交付状态重跑 locked restore、Release build、API、Unit 和 Docs 门禁。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：无。未修改 `build/api-snapshot-baseline.json` 或 `tests/Bing.Offices.Tests/PublicApiContractTest.cs` 的批准 hash。
- 根因复核：当前共享 metadata-only canonicalizer 已使四个 TFM 得到稳定 actual，但批准 baseline 只有 hash，没有对应的历史成员快照、生成上下文或批准记录。Git 历史未找到 `7B0BA279...` 或 `41B6D12...` 的批准来源；现有 Round 6/8 actual snapshots 彼此一致，均为当前 actual。
- 当前证据：四个 TFM 均为 Abstractions `737` members、actual `5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646`，approved `7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`；Core `273` members、actual `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`，approved `41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`；Npoi `1` member 且 actual 与 approved `A0DBE9808D82547601429D8958C7ED283467031A3763EB9037B19D03F19D80BD` 匹配。
- 已完成修复准备：复核 `build/ApiSnapshot/Program.cs` 与测试 link 的同一 `PublicApiSnapshot.cs`；复核 `artifacts/api-snapshot-review8` 四个实际快照；确认 API compare 会对四个 shipped TFM 返回 mismatch 并退出码 `1`；确认 net6/net8 `PublicApiContractTest` 各 `6/7`，唯一失败为批准 hash mismatch。
- 未完成与原因：无法从 hash 反推出成员级 actual-vs-approved diff；无法在无版本负责人批准的情况下恢复旧产物、批准当前 actual、修改 baseline 或编造 migration 决策。不得机械替换 expected hash。
- 验证：API snapshot 工具可构建；四 TFM compare 对 Abstractions/Core 真实失败、Npoi 通过；net6/net8 Unit 的唯一失败与该 mismatch 一致。因此本项保持 `PARTIAL/BLOCKED`。
- 后续条件：版本负责人提供批准 hash 对应的历史 DLL/成员快照和生成规范，或明确批准当前 actual；记录批准人、日期、版本、source/binary 影响和 migration 后再同步 baseline，并纳管工具、baseline、快照和 CI 入口。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：无性能实现修改；未篡改 benchmark 原始结果的 `budgetStatus`。
- 根因复核：BDN 六组正式结果、资源探针和 Round 4 两份 tail-latency JSONL 均可读取，但报告/JSONL 明确为 `budgetStatus=UNAPPROVED`；仓库没有批准 workload、p99/吞吐预算、容差、批准人、日期或具名 release waiver。
- 已完成修复准备：复核 tail-latency 覆盖并发 `1/4/16/64`、预热、5 轮、每档 `1280` 样本、p50/p95/p99、吞吐和环境身份；复核 `08-benchmark-report.md` 未宣称 zero-GC、真实 streaming 或完整压缩炸弹防护；确认任务 benchmark 原始证据路径未命中当前 `.gitignore`，但仍未被 Git 跟踪。
- 未完成与原因：执行器没有权限代表版本负责人批准性能预算或签署 waiver，也不能依据高并发波动自行判定 PASS；不能把测量值改写为已批准状态。
- 验证：正式 benchmark 和 tail-latency 原始证据存在且可复核；状态仍为 `UNAPPROVED`，因此本项保持 `PARTIAL/BLOCKED`。
- 后续条件：版本负责人记录冷/热 workload、环境、p99/吞吐预算、容差、批准人和日期，或提供带范围/风险的具名 release waiver；报告逐项标记 `PASS`/`FAIL`/`WAIVED`，原始 JSONL/BDN 证据进入 Git；再按批准条件重复验证。

### Round 9 验证汇总

- `git ls-files`：要求 lockfile、API 工具/baseline 和任务证据仍无跟踪输出，`FIX-001` 阻塞。
- 文件与 ignore 核验：要求文件均存在；任务证据路径未被忽略；顶层 `artifacts/` 仍为构建产物忽略规则。
- API 四 TFM compare：Abstractions/Core mismatch，Npoi match，退出码 `1`；未修改批准 baseline。
- net6/net8 API contract：各 `6/7`；唯一失败为批准 hash mismatch。
- net6/net8 Unit：各 `371/372`；唯一失败为同一 API hash mismatch。
- 隔离 locked restore：PASS，未改写 lockfile。
- 隔离 Release solution build：PASS，退出码 `0`。
- Docs package-only consumer：PASS，`11/11`；三个 Bing 包精确 `2.0.0` 且 `type=package`。
- 性能证据：正式 BDN、资源探针和两份 tail-latency JSONL 可读取；批准状态仍 `UNAPPROVED`。
- 诊断与差异：相关文件 `get_errors` 无诊断；`git diff --check` 退出码 `0`，仅有 CRLF/LF 提示。

### Round 9 汇总

- MUST_FIX：`FIX-001`、`FIX-003`、`FIX-004`。
- 已完成：三项的当前源码/证据/交付边界复核；API actual 快照、package source/cache 风险和性能未批准状态已明确记录；未伪造批准结论。
- PARTIAL：`FIX-001` Git tracking/clean clone；`FIX-003` approved API baseline 来源和批准决策；`FIX-004` performance budget/waiver 和 Git 证据交付。
- BLOCKED：Git staging/实际待交付提交、版本负责人 API 批准、版本负责人性能批准或 waiver。
- FAILED：无新增业务代码或测试代码失败；已有 API baseline mismatch 按真实结果保留。
- 执行终态：`PARTIAL`。
- 发布结论：`No-Go`。
- 下一步：由获授权交付/版本负责人完成上述外部条件后，重新运行独立 `code-reviewer` Review；不得修改 `review.md` 伪造通过。

## Review 修复 Round 10

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260831-001/review.md`
- 执行终态：`PARTIAL`
- 说明：本轮按用户明确要求仅处理 `MUST_FIX`，不处理 `FIX-006`。未修改 `review.md`、`plan.md` 或业务/测试代码；未执行 Git staging、commit、push、tag、PR、publish 或外部通知。三个 MUST_FIX 的本地复核和可执行准备已完成，但 Git 交付、API 批准和性能批准需要当前执行器无权代行的外部授权，因此不能记为 `COMPLETED`。

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：无；沿用现有 `.gitignore` 对 10 个 lockfile 的显式例外和 package-only consumer 配置。
- 根因复核：工作树存在 10 个 `packages.lock.json`，但 `git ls-files -- '*packages.lock.json'`、任务目录和 API 工具/baseline 的跟踪数仍为 `0`。默认缓存 strict restore 继续对三个 Bing `2.0.0` 包报 `NU1403`，说明既有全局缓存包含不同内容的同版本包。
- 已完成：全新隔离 `NUGET_PACKAGES` 与 HTTP cache 下 `dotnet restore Bing.Offices.sln --locked-mode --no-cache` 和 Release solution build 均成功；package-only Docs consumer 隔离 restore/build/test 为 `11/11`，assets 中三个 Bing 依赖均为精确 `2.0.0`、`type=package`；确认要求文件未命中当前任务路径的 ignore 规则。
- 未完成与原因：当前执行器不能在无用户授权下执行 `git add`，不能创建实际待交付提交或 clean clone，也不能批准继续使用可变同版本本地包身份。Git tracking 和 clean-clone 证据因此保持阻塞。
- 后续验证：获授权交付步骤纳管 lockfile、API 工具/baseline、任务报告、快照和需保留的原始性能证据；使用不可变包内容或唯一版本身份，明确本地 Bing source 与第三方 source/预热 cache；在实际待交付状态重跑 strict restore、Release build、API、Unit 和 Docs 门禁。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：无；未修改 `build/api-snapshot-baseline.json` 或 `tests/Bing.Offices.Tests/PublicApiContractTest.cs` 的批准 hash。
- 根因复核：`build/ApiSnapshot/Program.cs` 与 Unit 测试继续共享 `PublicApiSnapshot.cs`，四 TFM actual 稳定，但批准 baseline 只有 hash，没有可追溯的历史成员快照、生成上下文或版本负责人批准记录。
- 当前证据：net6/net8 Unit 各 `371/372`，唯一失败为 API approved hash mismatch；四 TFM API compare 对 Abstractions/Core 均失败并以退出码 `1` 结束，Npoi 匹配。Abstractions actual 为 `5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646`，approved 为 `7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`；Core actual 为 `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`，approved 为 `41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`。
- 已完成：复核 API 工具可构建、四 TFM 比较可执行且当前 mismatch 可稳定复现；未用替换 expected hash 的方式规避门禁。
- 未完成与原因：无法从 hash 反推出批准成员差异，也不能在无版本负责人批准的情况下恢复历史产物、批准当前 surface、修改 baseline 或编造 migration 决策。
- 后续验证：版本负责人提供批准 hash 对应的历史 DLL/成员快照和生成规范，或明确批准当前 actual；记录批准人、日期、版本、source/binary 影响和 migration 后同步 baseline；再验证 net6/net8 API contract 全绿和四 shipped TFM compare 全绿。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：无；未篡改性能原始数据或将 `UNAPPROVED` 改写为通过。
- 根因复核：正式 BenchmarkDotNet、资源探针和 tail-latency 原始 JSONL 均可读取，但性能报告和 tail-latency 数据仍无批准 workload、p99/吞吐预算、容差、批准人、日期或具名 release waiver；原始证据仍未被 Git 跟踪。
- 已完成：确认 tail-latency 覆盖并发 `1/4/16/64`、预热、5 轮、每档 `1280` 样本，并输出 p50/p95/p99、吞吐和环境身份；确认报告未宣称 zero-GC、真实 streaming 或完整压缩炸弹防护。
- 未完成与原因：执行器无权代表版本负责人批准预算或签署 waiver，不能依据测量值自行判定性能 PASS，也不能改写原始结果状态。
- 后续验证：版本负责人记录冷/热 workload、环境、p99/吞吐预算、允许波动、批准人和日期，或提供带范围与风险的具名 waiver；报告逐项标记 `PASS`/`FAIL`/`WAIVED`，原始 JSONL/BDN 证据进入 Git，再按批准条件重复运行。

### Round 10 验证汇总

- 默认 locked restore：FAIL，三个 Bing `2.0.0` 包报 `NU1403` 内容哈希冲突。
- 隔离 locked restore：PASS。
- 隔离 Release solution build：PASS，退出码 `0`。
- net6/net8 Unit：各 `371/372`，唯一失败为 API baseline mismatch。
- 四 TFM API compare：Abstractions/Core mismatch，Npoi match，退出码 `1`。
- 隔离 package-only Docs consumer：`11/11` PASS，三个依赖均为精确 `2.0.0`、`type=package`。
- `get_errors`：相关 `src`、`tests`、`build` 无诊断错误。
- `git diff --check`：退出码 `0`，仅有 CRLF/LF 提示。

### Round 10 汇总

- MUST_FIX：`FIX-001`、`FIX-003`、`FIX-004`。
- 已完成：本地根因复核、可执行工具/consumer/构建验证和后续验收条件记录。
- PARTIAL：三个 MUST_FIX 均因外部 Git 交付或负责人批准未闭环。
- BLOCKED：Git 纳管/clean clone、API baseline 批准、性能预算或 waiver。
- FAILED：无新增业务代码或测试代码失败。
- 回归验证：隔离 restore/build 和 package-only consumer 通过；既有 API 门禁失败按真实结果保留。
- 发布结论：`No-Go`。
- 下一步：由获授权交付/版本负责人完成阻塞条件后，重新进行独立 Review；不得修改 `review.md` 伪造通过。
