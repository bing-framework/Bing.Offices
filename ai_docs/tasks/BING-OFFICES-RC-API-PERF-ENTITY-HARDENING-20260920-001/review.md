<!-- AI_REVIEW_STATUS: BLOCKED -->
AI_TASK_ID: BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001
AI_REVIEWED_AT: 2026-09-22T11:20:00+08:00

## Round 43 当前审查分类

本节是当前机器可读的独立审查结论；下方 Round 1-42 的原始 findings 和总结保留为历史证据，不覆盖本节。

| 维度 | 当前状态 | 依据 |
|---|---|---|
| Implementation | `CLOSED` | Round 42 以前的实现、API、Provider SPI 和公开契约证据保持有效；本轮未改生产源码 |
| Test | `CLOSED` | 既有双 TFM 回归和职责级测试保持通过；本轮 solution/runner Release build 为 `0 warning / 0 error` |
| Performance | `OPEN_ACTIONABLE` | `targeted-performance-budget-v1-5rep.json` 中五个批准目标均 `5/5`，均超过 25% budget |
| Resource | `VERIFIED_BOUNDARY` | 本地 Job Object 2 CPU/4 GiB 资源边界、默认 64 MiB 资源拒绝和清理字段均可复核 |
| External | `BLOCKED_EXTERNAL` | 未取得真实生产机或外部 CI 运行结果；本地 profile 不得冒充外部门禁 |
| Release | `BLOCKED_APPROVAL` | 五项性能回退仍未修复，外部 Build/Test/API/Pack/Consumer Gate 未执行 |
| Goal | `IN_PROGRESS` | `FIX-002C` 是唯一 `OPEN_ACTIONABLE`；Relation 100K、c16/c64 和完整 A-I 按批准边界未重跑 |

当前 Finding 分类：`FIX-002A=VERIFIED_BOUNDARY`（Relation 100K before，禁止重跑）；`FIX-002B=CLOSED`（当前候选 materialization/binding 分离 E2E）；`FIX-002C=OPEN_ACTIONABLE`（五项 100K E2E 回退）；`FIX-003A=VERIFIED_BOUNDARY`（默认 `MaxWorksheetBytes=64 MiB`，Entity/Template 1M 为 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`）；`FIX-003B=VERIFIED_BOUNDARY`（c16/c64 queue/admission 证据足够，禁止同质长实验）；`FIX-003C=CLOSED`（既有本地 staging/cleanup 证据）；`FIX-003D=LOCAL_PROFILE_EVIDENCE`（最终候选本地 100K/500K/1M smoke）；`FIX-003E=BLOCKED_EXTERNAL`（外部 CI 未执行）。`OpenActionable=1`。

### Performance Budget v1 证据

结构化总产物：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/targeted-performance-budget-v1-5rep.json`。当前候选 hash 由原始 header 绑定：Benchmark `419DC8A4F2D7907BEBFFE8351A6AD671FCDD552397CF0C4134F4B824DFC1202F`、NPOI `E87FA5777867FB1D8B22B824EB580373306C75B95F311447F554BE70390F010E`、MiniExcel `3FDDC6AF4EE936A44E783B1208CA6B33C59705419585D8A353110980C4D7DCEB`。

| Target | Before median ms | After median ms | Delta | Samples | Result |
|---|---:|---:|---:|---:|---|
| NPOI async provider E2E | 4,734.42 | 13,784.323 | `+191.15%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| MiniExcel sync provider E2E | 1,423.58 | 7,599.0974 | `+433.80%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| Real IO `excel-file-sync` | 1,586.69 | 6,935.1346 | `+337.08%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| Real IO `excel-file-async` | 2,036.25 | 11,111.6705 | `+445.69%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| Real IO `excel-throttled-async` | 2,387.22 | 5,636.8019 | `+136.12%` | `5/5` | `REGRESSION_OVER_BUDGET` |

两组 Provider probe 是并发启动的顶层进程，Real IO 三场景串行执行；该执行拓扑已保留在 artifact 和报告中。它不改变五项均超预算的当前审查分类，但后续修复后的精准复测应采用串行或明确隔离的资源条件。

### 资源与外部环境证据

`windows-job-production-equivalent-smoke-{100k,500k,1m}.json` 及对应 child JSON 均记录 `passed`、`2/2` 操作、错误 `0`、`assignedToJob=true`、2 CPU 等效、4 GiB、TEMP `0/0/deleted`；Job 峰值分别为 `399,478,784`、`1,224,552,448`、`2,357,735,424` bytes。这是本地生产等价资源 profile 证据，不是真实生产机，也不构成外部 CI 通过。未提高默认 64 MiB 限制，未重跑 Entity/Template 1M、Relation 100K 或 c16/c64。

# 独立代码审查

## 历史审查结论（Round 1-42）

本轮以格式纠正后的批准文件及 capture-v2 candidate identity 为最终 API 验收证据。FIX-001 已形成完整闭环：用户批准文件、approved baseline、双 TFM 空差异、双 TFM Public API 合同测试、identity self-test、错误批准路径拒绝和正式 CI 调用链均指向同一批准工件。FIX-001 保持 `CLOSED`。

历史结论为 `NEEDS_FIX` / `PARTIAL`，仅描述 Round 1-42 的审查快照；当前状态以文档顶部 Round 43 分类为准。Round 10 已关闭上一轮 hotspot 诊断的实现级 `SHOULD_FIX`：Kill 后退出等待和 stdout/stderr 读取均有严格上限，非超时 worker 非零退出会形成结构化失败并停止后续场景，退出码 `2/3` 分别表示超时未验证/worker 失败，最终诊断身份也与中间产物及冻结正式候选明确隔离。当前没有用历史快照覆盖 Round 43 的 `FIX-002C=OPEN_ACTIONABLE`。

## 最终 API 证据

- 批准文件：`ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md`，当前 SHA-256 为 `B6BD00B1C6FFF4730AD37DAE52FCADAE42DEC18C9011C6BA1F583308524667F0`；所有审批范围均为 `APPROVED`。
- 最终 capture identity：`artifacts/api/fix001-20260921-capture-v2/api-candidate.json`。其 `candidateIdentity` 使用 `logical-v3`，`baseCommit=94bb52e84ffc70634067b04541857433ef7af9df`，`candidateSourceManifestSha256=9E552C7553D33E768637DA76FFA51706936B14B074B82F2BA888186AD5F9A1C2`，批准路径与批准哈希均与上项一致。
- approved baseline：`build/api-snapshot-baseline.json` 的完整 candidate identity 与 capture-v2 一致，`approvedBy=user:FIX-001`，`approvedAt=2026-09-21T17:12:16+08:00`。
- 双 TFM compare：`artifacts/api/fix001-20260921-compare-after-baseline-final/api-diff.json` 为 `{"net6.0":{},"net8.0":{}}`。
- Public API 合同：`artifacts/tests/fix001-20260921/net6.0-final2/public-api-net6.0-final2.trx` 与 `net8.0-final2/public-api-net8.0-final2.trx` 各 `9/9 PASS`，失败、错误、超时和未执行均为 `0`。
- identity self-test：本轮独立复跑退出码为 `0`；LF/CRLF checkout、source/assembly/nupkg/approval 篡改和错误 TFM 均按预期拒绝。`build/ApiSnapshot/Program.cs` 的 capture/validate 共享 `--approval` 路径规范化、仓库边界、路径一致性和哈希验证。
- 错误批准路径：`artifacts/api/fix001-20260921-compare-wrong-approval/` 与 `execution.md` 记录旧任务批准路径 compare 退出 `1`，错误为批准路径不匹配；当前实现的 `ValidateApprovalFile` 也明确拒绝配置路径与 identity 路径不一致。
- 正式 CI：`.github/workflows/ci.yml` 的 `release-gates` 在干净 checkout 中重新构建并 pack 四个生产包、上传 `artifacts/packages`；`api-gate` 通过 `needs: release-gates` 下载该产物后执行 identity self-test 和正式 compare。正式 compare 已显式传入本任务 `--approval`，identity self-test 命令未改。CI diff 为 `1 insertion / 1 deletion`。
- 版本文件未变：`version.props`、`version.dev.props`、`common.props`、`framework.props` SHA-256 分别为 `EC5201C9A9EC7F4981879FFA604C5BA2B35F89F064AB88C790ECFB6644DC5306`、`0A2F75F575DCC863D51F256A0D823F400FC33369FB1C7EB823CE6EBA9321872B`、`1C830F653129C28F59094F3F5886EB1BA1640ACBFE8984FF34F24FEDAB10ECA4`、`5B64D45A8FC516449251449C619E8F0434481A295D18A8B727AC4978FF5FDE35`，与计划基线一致。

## 编码与文件格式核对

- 本文件、`api-approval-request.md`、`.github/workflows/ci.yml` 与 `build/api-snapshot-baseline.json` 均为严格 UTF-8、无 BOM、LF、末尾一个真实换行。批准文件最终尾字节为 `0A`，不存在字面量 `\n` 尾缀。
- `git diff --check` 退出成功，仅报告一个既有 ProfileFixtures 行尾提示，不属于本轮文件。

## 计划验收矩阵

| Phase | 结果 | 当前证据 |
|---|---|---|
| P0 | PARTIAL | 当前候选证据可追踪；初始 clean-before 不完整 |
| P1 | PASS | rename/新增 API 已批准；最终 baseline、双 TFM compare 和 Public API 合同通过 |
| P2 | PASS | public-only 第三方 Provider fixture 与 capability SPI 已验证，API 范围已批准 |
| P3 | PASS | MiniExcel importer 已拆责，职责级双 TFM 测试通过 |
| P4-P6 | PASS | Dynamic/List/RawDate/Relation 真实路径与合同测试通过 |
| P7 | PASS | Entity 固定 Cell、mapping/converter/validation 顺序、List Region 边界均有双 TFM直接测试 |
| P8 | PARTIAL | Template 主路径和边界测试已覆盖；完整大规模资源门禁未完成 |
| P9 | PARTIAL | 专项与 Public API 双 TFM通过；最终性能/资源发布门禁仍缺失 |
| P10 | PARTIAL | 有效 after 与部分 before 存在；A-I 完整可比矩阵和批准阈值未完成 |
| P11 | PARTIAL | 本地 Job Object 与若干容量场景已验证；Entity/Template 1M、完整矩阵、生产/外部环境未验证 |
| P12 | PARTIAL | 包、Consumer、最终 API 身份链和 CI 接入已完成；FIX-002/003 尚未关闭 |

## Findings

### FIX-001

严重程度：HIGH  
处理要求：MUST_FIX  
处理状态：CLOSED

问题：已解决。API 批准工件、candidate identity、baseline、双 TFM compare、Public API 合同与正式 CI compare 已统一绑定本任务批准文件。

证据：capture-v2 与 baseline 均记录批准路径 `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md`、批准 SHA-256 `B6BD00B1C6FFF4730AD37DAE52FCADAE42DEC18C9011C6BA1F583308524667F0` 和 source-scope SHA-256 `9E552C7553D33E768637DA76FFA51706936B14B074B82F2BA888186AD5F9A1C2`；final API diff 双 TFM为空；Public API final2 双 TFM各 `9/9 PASS`；identity self-test 退出 `0`；错误批准路径 compare 按预期失败；CI 正式 compare 显式传入相同 `--approval`，并消费 `release-gates` 新鲜打包产物。

影响：成员批准、候选来源、程序集、四个 nupkg、批准文档和 CI 执行已形成一致的可验证身份链。

修复目标：已达到。FIX-001 不因生产/外部 CI 尚未执行而重新打开；该外部环境缺口归入 FIX-003。

验证方式：以本节 capture-v2/B6BD 与 final/final2 证据为准。

### FIX-002（历史 Round 1-42 finding）

严重程度：HIGH  
处理要求：MUST_FIX  
处理状态：历史快照 OPEN（当前已拆分为 Round 43 状态）

问题：A-I 完整性能矩阵和正式阈值尚未闭合。Relation 100K before、真实 materializer/binding before/after 仍缺失；Entity/Template 只有当前 API 下的 after，旧基线不存在相同 API，只能标记 `NOT_APPLICABLE_BEFORE`，不能据此推导全部优化收益。

证据：`benchmark-report.md` 明确保留 `NOT_VERIFIED`，并记录 Relation 100K before、真实 materializer/binding 和阈值批准缺口。100K/500K Entity/Template 仅为单次 after 样本。本轮独立核验 Release Benchmark DLL 物理 SHA-256 为 `81249E5E1247272488B10D5FFA5B7CEB5ABC42AE92AEC118D5F014CA5F28D649`，与三份 final diagnostic header 完全一致：100 ms before 的六个场景均为 `timeout-killed`/`not-verified`，六个 summary 的五项中位数全部为 `null`；60 秒 before 中 RawDate 3/10/30 和 Relation 1K/10K 均为 3 个完整样本，Relation 100K 在 `60061.3724 ms` 超时，`completedRepetitions=0` 且五项中位数全部为 `null`；after 六个场景均 `passed` 且各有 3 个样本。三份文件的四程序集 candidate identity 完全一致，但其 Benchmark DLL 与冻结正式 after 候选 `6CBFFF...` 不同；`execution.md` 和 `benchmark-report.md` 均将 `81249...` final diagnostic、此前 `4CBB...`/`1156...` 中间诊断和 `6CBFFF...` 冻结候选明确分层，未迁移或合并统计。

影响：尚不能用可重复、同环境、同数据集的结果证明全部热点没有性能或分配回退，也不能执行正式 RC 阈值判定。

修复目标：在同一 candidate identity 与固定数据集下补齐适用的 A-I before/after，记录 wall time、吞吐、allocated bytes、GC、峰值工作集和重复次数；无法构造 before 的新增 API 必须保留可审计的 `NOT_APPLICABLE_BEFORE` 理由，并取得阈值批准。

修复要求：不得把方向性 microbenchmark、单次 after 或不同候选历史产物替代正式可比矩阵；不得伪造缺失数据。

本轮附加审查要求（SHOULD_FIX）：`CLOSED`。

1. `RunWorkerProcessAsync` 在 worker timeout 后只等待最多 `5000 ms`；即使未观察到退出，也继续生成 `exitObserved=false` 的结构化诊断。stdout/stderr 各自最多等待 `1000 ms`，未完成时写入固定占位值；父进程不再无限等待 worker 或诊断管道。
2. 非超时 worker 非零退出由 `workerFailed` 分支写入 `hotspot-failure` 和 `status=failed` 的 summary，诊断先截断到 4096 字符；`RunAsync` 随即停止后续场景并设置退出码 `3`。超时场景仍使用退出码 `2`，不会与 worker failure 混为同一结论。
3. 三份 final diagnostic 均绑定最终诊断 DLL `81249...`；报告明确把此前 `4CBB...`/`1156...` 文件作为历史中间诊断，把 `6CBFFF...` 保持为冻结正式候选。最终 100 ms/60 秒 timeout 记录采用 `timeout-killed`，部分 summary 统计全部为 `null`，after 文件六场景全部完整。

残余测试风险：当前 final diagnostic 覆盖了正常完成和成功 Kill 的 timeout 路径，非超时 worker 非零退出、Kill 失败及 `exitObserved=false` 仍主要由源码控制流验证，没有独立的故障注入产物。该缺口不改变本轮实现级 `SHOULD_FIX` 的关闭结论，也不能被用作 FIX-002 正式性能门禁证据；后续若继续演进该探针，宜补充可控 worker 故障注入测试。

验证方式：核对 benchmark manifest、候选程序集 identity、运行时/硬件、数据集、重复次数、统计结果和批准阈值，保证报告可复跑且不混用历史候选；增加可控 worker 非零失败、超时前已完成部分样本及超时/退出竞态测试，断言错误原因不丢失、不完整 summary 的统计字段为 `null`、总进程保持非零退出。

### FIX-003（历史 Round 1-42 finding）

严重程度：HIGH  
处理要求：MUST_FIX  
处理状态：历史快照 OPEN（当前已拆分为 Round 43 状态）

问题：完整资源与外部环境发布门禁未完成。Entity/Template 1M 在默认 `MaxWorksheetBytes=64 MiB` 的 ZIP 预检阶段被结构化拒绝，没有 roundtrip；完整取消/失败提交矩阵、生产机器、真实容器和外部 CI 尚未形成可审计通过证据。

证据：`resource-report.md` 记录本地 2 CPU 等效/4 GiB Windows Job Object、简单列表 SXSSF 100K/500K/1M、Entity/Template 100K/500K 和 staging 矩阵结果；同时明确 Entity/Template 1M 为 `NOT_VERIFIED`，生产/外部环境与正式容量阈值未验证。当前 CI YAML 接入正确，但没有外部运行结果。

影响：不能证明 RC 在目标生产约束下满足容量、取消、原子提交、临时文件清理和稳定性要求。

修复目标：按已批准阈值完成适用的 100K/500K/1M 资源矩阵、取消/失败提交、生产机器和外部 CI；保留默认资源上限，若 1M 不属于支持范围，应形成明确的产品能力决策与发布验收，而不是静默放宽限制。

修复要求：每个未执行环境继续标记 `NOT_VERIFIED`；不得用本地单次样本、超时演练或结构化拒绝替代成功容量证据。

验证方式：检查受限运行原始 JSON/log、退出码、峰值内存、CPU 限制、输出完整性、目标文件保护、临时文件清理、重复次数、环境身份和外部 CI 链接。

## 历史审查总结（Round 1-42）

- FIX-001：`CLOSED`，最终 API 治理和 CI 身份链已闭合。
- FIX-002：`OPEN / MUST_FIX`，完整性能矩阵与阈值未闭合；Round 10 已关闭 Kill 后有界退出、非超时 worker 结构化失败、部分统计置空和诊断 identity 隔离的实现级 `SHOULD_FIX`，本轮未发现新的 `SHOULD_FIX`。
- FIX-003：`OPEN / MUST_FIX`，完整资源、生产和外部 CI 门禁未闭合。
- 历史最终状态：`NEEDS_FIX`，仅适用于 Round 1-42；当前机器状态以文档顶部和下方 Round 43 总结为准。

## Round 43 当前审查总结

- `FIX-001=CLOSED`；Round 42 以前的实现、API、测试和候选身份链没有被本轮源码变更破坏。
- `FIX-002A=VERIFIED_BOUNDARY`、`FIX-002B=CLOSED`、`FIX-002C=OPEN_ACTIONABLE`。五个批准的 100K E2E 目标各有 `5/5` 样本，delta 均超过 `25%`，应先定位并修复性能回退；本轮没有重跑 Relation 100K、c16/c64 或完整 A-I。
- `FIX-003A=VERIFIED_BOUNDARY`、`FIX-003B=VERIFIED_BOUNDARY`、`FIX-003C=CLOSED`、`FIX-003D=LOCAL_PROFILE_EVIDENCE`、`FIX-003E=BLOCKED_EXTERNAL`。100K/500K/1M smoke 的本地 2 CPU/4 GiB Job Object 证据完整，但没有真实生产机或外部 CI 结果；Entity/Template 1M 继续默认 64 MiB 的 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`。
- 当前没有新增实现级 `MUST_FIX`/`SHOULD_FIX`；唯一 OPEN_ACTIONABLE 是由批准预算确认直接产生的 `FIX-002C` 性能修复任务。外部环境缺口保持 `BLOCKED_EXTERNAL`，不通过本地结果伪造关闭。
- 当前审查状态保持 `BLOCKED`，Progress 约 `99.5%`，Goal=`IN_PROGRESS`，OpenActionable=`1`，不能标记 `COMPLETED`。
