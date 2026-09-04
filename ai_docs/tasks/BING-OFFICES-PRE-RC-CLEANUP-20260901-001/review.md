<!-- AI_REVIEW_STATUS: BLOCKED -->
AI_TASK_ID: BING-OFFICES-PRE-RC-CLEANUP-20260901-001
AI_REVIEWED_AT: 2026-09-03T09:12:51.0712637+08:00

# 当前独立验收（正式 API baseline 收口后）

## 最终结论

`BLOCKED`，RC 继续为 `No-Go`。本次未发现需要实施 Agent 修复的新增代码缺陷，因此不生成 `FIX-xxx`；阻断来自尚未满足的发布和交付门禁，而非用测试失败或未批准 API 差异掩盖实现问题。

本 Reviewer 仅修改本报告，未修改业务代码、测试代码、[plan.md](plan.md) 或 [execution.md](execution.md)，未执行暂存、提交、推送、重置或清理操作。

## 复审边界与 Git 证据

- 已基于计划、当前执行报告、旧 Review、当前 Git Diff、CSV/Mapping/NPOI 主链、相关 Unit/Integration 测试、API 工具及 RC 报告复审。
- 当前受跟踪修改涵盖 Abstractions、Core、NPOI、测试、Benchmark、文档；新增的 CSV Support、Mapping cache key、NPOI sheet executor 与任务证据均在计划范围内。未发现明确可归属为无关的生产行为改动。
- `git diff --check` 通过；输出仅为 CRLF/LF 归一化警告，无空白错误。
- `.agents/runtime/current-task.json` 处于删除状态，任务目录和原始证据仍未被 Git 跟踪。因此当前报告不能证明 clean clone 能取得同一套 artifacts；Reviewer 未擅自恢复或纳入这些文件。
- 已审阅的源码及命令输出中未发现提示注入、密钥泄露、生产 `Task.Run()`、`.Wait()`、`.Result` 或新增生产友元程序集证据。

## 本次独立验证

| 命令/检查 | 结果 | 证据 |
| --- | --- | --- |
| `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release --no-restore -f net8.0` | PASS，`384/384` | `tests/Bing.Offices.Tests/TestResults/independent-review-net8.trx` |
| `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release --no-restore -f net6.0` | PASS，`384/384` | `tests/Bing.Offices.Tests/TestResults/independent-review-net6.trx` |
| `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release --no-restore -f net8.0` | PASS，`15/15` | `tests/Bing.Offices.Tests.Integration/TestResults/independent-review-integration-net8.trx` |
| API Snapshot 对正式 baseline 比较 | PASS，`netcoreapp3.1`、`net6.0`、`net8.0` | `artifacts/independent-review-api-snapshot/` |
| `git diff --check` | PASS | 仅 CRLF/LF 警告 |

测试构建仍出现 netcoreapp3.1 第三方依赖支持提示、nullable、obsolete 和 xUnit analyzer 警告；本次没有错误，亦未通过抑制警告或弱化断言取得通过。

## API 与主流程验收

- 用户已批准正式 API baseline；当前 Release 产物同 [api-diff.md](api-diff.md) 的 formal baseline 一致：Abstractions `723` members、Core `194` members、NPOI `1` member。独立 compare 已重新通过，旧 Round 9 的 baseline mismatch 已解除。
- `CsvEntityImporter`/`CsvEntityExporter` 通过 `IExcelMappingPlanFactory` 接入 mapping 主流程；`CsvPipelineSupport` 负责 RFC4180 reader、受限流与公式防护。对 BOM、控制符、Unicode 空白和带符号数值表达式的安全策略及 options 验证均由全量 Unit 覆盖。
- Mapping factory 的请求配置按方向合并；cache key 和容量传递已接入真实 Factory。此前 `FIX-007` 的容量/日志不一致已由回归、基准和现有实现共同证实为已解决。
- NPOI 公开面仅保留链式 `AddBingOfficesNpoi(IServiceCollection)`；当前 exact-member/API Unit 已通过，未见 Core→NPOI 反向依赖或 NPOI 类型泄漏。
- 计划中的 P3 重构和 public execution detail 治理仍是部分完成，未在本次被误写为完成。

## 计划验收矩阵

| 计划项 | 状态 | 验收依据与剩余项 |
| --- | --- | --- |
| P0-01 / P0-03 | `PARTIAL` | 基线、扫描和决策证据存在；任务 artifacts 尚未形成可 clean-clone 的交付输入。 |
| P0-02 / P4 | `PARTIAL` | net6/net8 Unit 均 `384/384`，本次 net8 Integration `15/15`；缺失 runtime TFM 仍未运行。 |
| P1-01 ～ P1-05 | `PASS`（可运行 TFM） | Mapping、v1 方向、relation 异常、CSV 防护/options、DI/不支持能力契约由当前全量 Unit 覆盖并通过。 |
| P1-06 / P1-07 / P5-03 | `PARTIAL` | Excel child-process Probe 与 Failure Workbook 回归有效；压缩输入、取消和 Failure Workbook 双 DOM 的资源矩阵仍不足。 |
| P2-01 / P2-02 | `PARTIAL` | 已批准删除/internal 化及正式 baseline 已闭环；DataTable 显式兼容和剩余 execution detail 尚待逐符号治理。 |
| P2-03 | `PASS` | 唯一链式 DI 入口、精确成员基线和当前 Unit 已验证。 |
| P3-01 ～ P3-04 | `PARTIAL` | NPOI sheet executor、CSV support 和 cache key 已拆分；Failure Writer 拆分、rule index/dynamic compiler 与完整量化门禁尚未完成。 |
| P5-01 ～ P5-02 | `BLOCKED` | ShortRun 为局部证据，`100k` Import 有 `3` 个异常待复核，性能预算未批准。 |
| P6-01 / P6-02 | `BLOCKED` | 文档未夸大资源限制；但 runtime、资源、性能及 clean-clone 交付门禁未关闭。 |

## 阻断项、风险与解除条件

### HIGH（发布阻断，不是本轮代码修复任务）

1. `netcoreapp3.1` runtime 缺失，且其 Unit 无独立运行结果；安装对应 runtime 后重跑受控 TFM 测试。
2. 性能门禁未完成：统一 Benchmark 矩阵和具名/获批预算缺失，且 StreamPipeline `100k` Import 的 `3` 个异常尚未复核。
3. 资源门禁未完成：现有 Probe 不证明压缩输入、任意 NPOI DOM、取消延迟或 Failure Workbook 双 DOM 的硬上限；应补独立子进程样本并维持文档边界。
4. 任务报告、正式 baseline、TRX/JSONL/BDN artifacts 尚未进入获授权交付输入；必须由授权人员完成 staging/提交后，以 clean clone 验证包消费者和证据可复现性。

### MEDIUM

1. API 分层仍遗留 DataTable 显式兼容和 public execution detail，需按 [plan.md](plan.md) 的逐符号治理流程另行裁决；当前不能声明 API 收敛已经完成。
2. 长路径 NuGet cache 的 `MSB3106` 仍使 package-only consumer 仅为短路径环境的部分验证，不能泛化为交付环境无条件通过。

### LOW

- 构建警告和 CRLF/LF 提示不构成本次失败，但发布前应在不降低分析器/TFM 支持标准的前提下持续治理。

## 最终验收 Checklist

- [x] 已阅读计划、当前执行报告、旧 Review、真实代码、Git Diff、测试和报告。
- [x] 已独立重跑 net8/net6 Unit、net8 Integration 与正式 API compare。
- [x] 正式 API baseline 已获批准，且当前 compare/Unit 通过。
- [x] 旧 `FIX-001`～`FIX-005`、`FIX-007` 无回归；`FIX-006` 保持 OPTIONAL 的交付治理状态。
- [x] 未发现需要新建 `FIX-xxx` 的可复现代码问题。
- [ ] 缺失 TFM runtime 验证完成。
- [ ] 性能预算、100k 异常复核及资源矩阵完成。
- [ ] 任务证据进入授权交付并以 clean clone 复现。
- [ ] 剩余 API 治理与长路径 package consumer 限制关闭或获正式 waiver。

## 当前裁决

正式 API baseline 已从旧 Round 9 的阻断项转为通过。由于其余 RC Go 条件仍未满足，本次最终结论只能是 `BLOCKED`，不生成修复清单，也不进入自动修复。

---

# 历史 Round 9 复审（已被上述当前裁决取代）

## 验收摘要

结论：`BLOCKED`；RC 保持 `No-Go`。

本次按 `plan.md`、Round 8 `execution.md`、上一版 `review.md`、当前 Git Diff、相关生产/测试/Benchmark 源码及本次独立命令复审。上一轮唯一纳入修复范围的 `FIX-007` 已解决，未发现其回归或新的必须修复代码问题；但正式 API 快照、`netcoreapp3.1` runtime、完整性能/资源门禁以及 clean-clone 证据交付仍未闭环，故不能判为 `PASS` 或 `PASS_WITH_ISSUES`。

Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

## Review 边界与 Git 分析

- 计划范围包括 Abstractions、Core、NPOI、Unit/Integration/Docs/ResourceProbe、Benchmark、API 治理和 RC 证据。当前受跟踪源码、测试和文档改动总体位于该范围。
- `git diff --check` 通过；当前仅输出 CRLF/LF 归一化提示，不存在空白错误。此前 `CsvEntityPipeline.cs` 的 EOF 额外空行已消失。
- 当前任务目录仍为未追踪目录，`.agents/runtime/current-task.json` 仍显示删除状态；因此 clean clone 无法取得任务报告、原始 Benchmark/Probe 与收口证据。这是已知交付复现阻断，Reviewer 未擅自恢复、暂存或修改 Git 状态。
- 本次 Benchmark 运行会更新被忽略的 `BenchmarkDotNet.Artifacts`，未发现运行期间有受跟踪业务或测试文件被自动修改。
- 未在已读代码和命令输出中发现提示注入、密钥泄漏、生产 `.Result`、`.Wait()`、`Task.Run()` 或生产友元程序集扩张的新证据。

## 上一轮 FIX 复审

| FIX | 处理要求 | 状态 | 本次证据 |
| --- | --- | --- | --- |
| `FIX-001` | `MUST_FIX` | `RESOLVED` | 完整 net8 Unit 中 CSV 回归未失败；既有精确 CSV 输出断言仍在测试集内。 |
| `FIX-002` | `MUST_FIX` | `RESOLVED` | `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot()` 仍使用旧正式 hash；本次完整 net8 Unit 的唯一失败仍是该真实 mismatch，未通过更新基线或跳过断言掩盖。 |
| `FIX-003` | `MUST_FIX` | `RESOLVED` | 当前资源报告仍将 Excel child-process 与 mapping/unique Probe 分离，并未宣称 NPOI DOM 的硬内存上限。 |
| `FIX-004` | `SHOULD_FIX` | `RESOLVED` | `decisions.md` D-013 和 `progress.md` 保持延期、批准和 RC 影响的明确记录。 |
| `FIX-005` | `SHOULD_FIX` | `RESOLVED` | `execution.md` 为合法 `BLOCKED` 终态，任务运行时已由 `task-finish` 收口。 |
| `FIX-006` | `OPTIONAL` | `NOT_RESOLVED` | 任务证据目录仍未 Git 跟踪；保持 OPTIONAL，不在本轮新增修复任务。 |
| `FIX-007` | `SHOULD_FIX` | `RESOLVED` | Provider 的显式 `cacheCapacity` 已传入 internal Factory；定向 Unit `2/2` 通过；本次独立 TenantPlanCache ShortRun 在 100/1000 tenant 均输出与实际 Factory 一致的 `capacity=50/256` 和 `observed=True`。 |

## 计划验收矩阵

| 计划项 | 状态 | 独立验收结论 |
| --- | --- | --- |
| P0-01 / P0-03 | `PARTIAL` | 环境和候选扫描证据存在；任务报告未进入 Git，clean-clone 交付不可验证。 |
| P0-02 / P4 | `PARTIAL` | 可运行 TFM 的构建和专项回归有效；完整 net8 Unit `383/384`，唯一 API hash mismatch；`netcoreapp3.1` runtime 未运行。 |
| P1-01 至 P1-05 | `PARTIAL` | Mapping、迁移、关系、CSV 和 DI 的既有证据未见回归；完整 Unit 唯一失败不在这些行为。 |
| P1-06 / P1-07 / P5-03 | `PARTIAL` | Failure Workbook 专项与 ResourceProbe 报告存在；真实压缩输入、Failure Workbook 双 DOM 和取消资源矩阵仍缺失。 |
| P2-01 / P2-02 | `PARTIAL` | 已删除和 internal 化的子集仍可构建；formal API baseline 与剩余 execution detail 治理未完成。 |
| P3-01 | `PARTIAL` | 当前 `NpoiExcelImporter` 只保留请求转换、Plan 选择和外层编排；Sheet 列绑定、动态列、表头与行执行已委派至 `NpoiImportSheetExecutor`。完整 API 门禁未绿，故不升级为发布验证通过。 |
| P3-02 | `PARTIAL` | CSV Reader/Writer/LimitedStream 已拆出；Failure Workbook 的产物构建和临时提交尚未拆分。 |
| P3-03 | `PARTIAL` | `FIX-007` 已解决；缓存键提取、显式容量、命中/淘汰回归和真实 Benchmark 已验证。rule-index/dynamic compiler 与完整性能矩阵仍未完成。 |
| P3-04 | `PARTIAL` | 所有权与同步 API 决策已记录；DOM 取消延迟和 Failure Workbook 峰值未形成完整量化门禁。 |
| P5-01 | `PARTIAL` | Tenant cache 与属性访问器 ShortRun 已有真实产物；仍不是计划要求的统一完整 Benchmark 与获批预算。 |
| P6-01 / P6-02 | `BLOCKED` | 文档/发布清单维持 No-Go 表述；API、runtime、性能、资源和可交付 Git 证据未关闭。 |

## FIX-007 代码与运行证据

- `ExcelMappingPlanFactoryProvider.CreateDefault()` 当前公开 `cacheCapacity` 可选参数，并将其传给 internal `ExcelMappingPlanFactory` 构造器；默认 DI 注册仍显式使用 `256`。
- `TenantPlanCacheEviction()` 以同一局部 `capacity` 创建 Provider Factory、输出容量，并在 `TenantCount > capacity` 时直接断言首个计划发生重建。`TenantPlanCache()` 单独使用默认容量 `256`，两类场景的语义没有混写。
- `MappingPlan_ProviderCacheCapacity_ShouldDriveEviction()` 验证容量 `2` 时 hit、租户隔离、淘汰、重建及重建列配置；`MappingPlan_ProviderDefaultCache_ShouldRetainThenEvictPlans()` 验证默认容量的命中与超容量重建。
- 本次独立命令：

	```powershell
	dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter "FullyQualifiedName~ReviewFixRegressionTest.MappingPlan_ProviderCacheCapacity_ShouldDriveEviction|FullyQualifiedName~ReviewFixRegressionTest.MappingPlan_ProviderDefaultCache_ShouldRetainThenEvictPlans" --logger "trx;LogFileName=review-round9-cache.trx"
	```

	结果为 `2/2` 通过，TRX 为 `tests/Bing.Offices.Tests/TestResults/review-round9-cache.trx`。

- 本次独立命令：

	```powershell
	$env:NUGET_PACKAGES = $null
	dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- -j short -m 3 -f "*TenantPlanCacheBenchmarks*"
	```

	4 个 ShortRun 场景完成；最新日志为 `BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.TenantPlanCacheBenchmarks-20260902-215154.log`。`TenantCount=100` 反复输出 `capacity=50`、`observed=True`，`TenantCount=1000` 反复输出 `capacity=256`、`observed=True`，无 `CACHE_EVICTION unexpected` 异常。该结果证明容量/日志/淘汰判断来自同一 Factory 输入；ShortRun 的宽置信区间不作为性能发布预算结论。

## 实际验证

| 命令 | 结果 |
| --- | --- |
| `git diff --check` | PASS；仅 CRLF/LF 提示。 |
| `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore` | PASS；`0 error / 4 warning`，其中 Benchmark 项目保留一个已有 `IgnoreNullValues` obsolete 警告。 |
| FIX-007 定向 Unit net8 | PASS，`2/2`。 |
| TenantPlanCache BenchmarkDotNet ShortRun | PASS，4 场景完成，容量与实际 Factory 一致。 |
| `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build` | FAIL，`383/384`；唯一失败为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。Abstractions expected `7B0BA279...F5F577`，actual `7F9A2AA8...2972943`。 |

## API、架构、性能与文档审查

- `cacheCapacity` 作为 `EditorBrowsable(Never)` Provider SPI 的可选尾参数，未暴露 internal Factory；当前 NPOI、Benchmark 和 ResourceProbe 仍经 SPI 获取 `IExcelMappingPlanFactory`，未见 Core 到 NPOI 的反向依赖。
- `PropertyAccessorBenchmarks` 已注册进现有 BenchmarkSwitcher，并有 MemoryDiagnoser；其独立 ShortRun 可作为计划矩阵的局部证据，但不等价于真实导入导出端到端预算。
- `benchmark-report.md` 明确标注 ShortRun、未批准预算、100k Import 异常计数和资源边界，未将本次缓存修复夸大为 RC 性能通过。
- `release-checklist.md` 仍将正式 API baseline、计划 Benchmark 矩阵、ResourceProbe 边界、任务证据 Git 交付和缺失 runtime 标记为未关闭，表述与实际验证一致。

## 阻断项与兼容风险

### HIGH

- 正式 API baseline 未获批准。完整 net8 Unit 的唯一失败为 Abstractions public member hash mismatch；不得机械更新 hash 或降低断言。
- `netcoreapp3.1` 仍是当前目标 TFM，但本机缺少 runtime；相关 Unit 未验证。
- Release Go 条件要求的 Failure Workbook 双 DOM/压缩输入/取消资源样本、完整统一 Benchmark 矩阵与获批性能预算尚未具备。
- 任务目录、原始证据和执行报告尚未进入 Git，clean clone 无法复现当前 Review 输入或 package/benchmark/resource 结论。

### LOW

- `FIX-006` 继续未处理，属于当前定义的 OPTIONAL 交付治理项。
- `CsvEntityPipeline.cs` 的 EOF 空白已修复；CRLF/LF 提示不构成 `git diff --check` 失败。

## 最终验收 Checklist

- [x] 已读取计划、当前执行报告、上一轮 Review、当前 Diff、相关源码、测试和 Benchmark 报告。
- [x] 已逐项复核上一轮 `FIX-001` 至 `FIX-007`。
- [x] `FIX-007` 的 Provider 容量传递、直接 Unit 和独立 Benchmark 均通过。
- [x] 已复核 `git diff --check` 与 Benchmark 编译。
- [x] 已重跑完整 net8 Unit 并保留 formal API hash 的真实失败。
- [x] Reviewer 未修改业务代码、测试、计划或执行报告。
- [ ] 正式 API baseline 已批准并在可运行 TFM 全绿。
- [ ] `netcoreapp3.1` runtime 验证完成。
- [ ] 完整性能/资源门禁、Failure Workbook 双 DOM 样本及 clean-clone 交付证据关闭。

## 最终结论

`BLOCKED`。`FIX-007` 已解决，不新增 FIX。RC 继续 `No-Go`，直至 API baseline 获批、缺失 TFM 验证完成、性能/资源门禁闭环，且任务证据被纳入可 clean-clone 复现的获授权交付输入。
