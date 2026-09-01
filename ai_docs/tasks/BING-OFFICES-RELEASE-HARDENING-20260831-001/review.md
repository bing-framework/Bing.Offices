<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260831-001
AI_REVIEWED_AT: 2026-09-01T15:30:46.6539289+08:00

# Review Fix Round 10 独立复审

## 验收摘要

本次复审以 `plan.md`、Round 10 `execution.md`、上一轮 `review.md`、当前源码、Git Diff、锁图、API baseline、测试、包和性能原始证据为依据，优先验证 `FIX-001`、`FIX-003`、`FIX-004`，并检查 `FIX-006`、`FIX-007`、`FIX-008` 是否变化。Round 10 没有业务或测试代码修改，只记录了外部授权阻塞；Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

最终结论：`NEEDS_FIX`，发布判定保持 `No-Go`。

- `FIX-001`：`PARTIAL/NOT_RESOLVED`。默认缓存和全新隔离缓存的 strict restore 当前均通过，Release solution build 也通过；但 10 个 lockfile、API 工具/baseline 和任务证据目录的 Git 跟踪数仍为 `0`，没有实际待交付提交/clean clone 复现，也没有不可变包版本身份。此前同版本不同内容导致的 `NU1403` 本轮不再复现，但仍证明覆盖已消费版本会破坏复现。
- `FIX-003`：`PARTIAL/NOT_RESOLVED`。net6/net8 Unit 各 `371/372`；四 TFM compare 对 Abstractions/Core 全部失败，Npoi 匹配。没有新增批准产物、成员快照或版本负责人决策。
- `FIX-004`：`NOT_RESOLVED`。两份 tail-latency JSONL 各 25 行，批准标记均为 `0`；报告仍为 `PARTIAL/UNAPPROVED`，没有预算、批准人、日期或具名 waiver，原始证据也未进入 Git。
- `FIX-006`：`NOT_RESOLVED`。Round 10 按 `fixScope=must` 未处理；仍无明确 XLS/XLSX 越界 `ByIndex`、`SheetVisibility.VeryHidden` 测试或客户端互操作/waiver。
- `FIX-007`：`RESOLVED`。全新缓存 package-only Docs restore/build/test 为 `11/11`；三个 Bing 依赖均为精确 `2.0.0`、`type=package`，SHA-512 与上一轮一致。
- `FIX-008`：`RESOLVED`。隔离 solution build 和 Unit 测试入口保持可执行；当前唯一 Unit 失败仍归属于 `FIX-003`。

## 上一轮 FIX 复审矩阵

| FIX | Round 10 执行状态 | 本轮复审状态 | 结论 |
| --- | --- | --- | --- |
| `FIX-001` | PARTIAL/BLOCKED | PARTIAL/NOT_RESOLVED | 当前默认与隔离 restore 均通过；Git 交付、clean clone 和不可变包身份仍未成立。 |
| `FIX-003` | PARTIAL/BLOCKED | PARTIAL/NOT_RESOLVED | 自动门禁稳定失败，没有批准依据。 |
| `FIX-004` | PARTIAL/BLOCKED | NOT_RESOLVED | 测量存在，批准预算/waiver 和 Git 交付仍缺。 |
| `FIX-006` | SKIPPED/OUT_OF_SCOPE | NOT_RESOLVED | selector 明确边界和客户端政策没有变化。 |
| `FIX-007` | 历史 COMPLETED | RESOLVED | package-only consumer 在全新缓存中再次通过。 |
| `FIX-008` | 历史 COMPLETED | RESOLVED | Unit 项目、构建和运行入口保持有效。 |

## 主要发现

### BLOCKER-001：Git 交付和不可变包身份仍未成立

- 分支为 `master`，HEAD 为 `1968b24a3ab07b44c3b386a3f761fcdff2fc4315`。
- 工作树存在 10 个 `packages.lock.json`，但 lockfile 跟踪数为 `0`；任务目录和 `build/ApiSnapshot/**`/baseline 的跟踪数也均为 `0`。
- Git status 继续将 lockfile、API 工具/baseline、00-11 报告、API 快照和 benchmark 原始证据列为未跟踪。
- 本轮默认 `dotnet restore Bing.Offices.sln --locked-mode --no-cache` 已通过；全新隔离 `NUGET_PACKAGES`/HTTP cache 的 strict restore 与 Release build 也通过。因此 Round 9 的“当前默认 restore 失败”不再是现状。
- 同一批 `2.0.0` 包曾因全局缓存内容不同产生 `NU1403`；本轮缓存更新后恢复通过，不代表版本身份已经不可变。发布模型仍不能依赖覆盖同版本包或清理/更新本机缓存。

当前工作树可构建是有效进展，但计划要求的是可由 Git 取得的待交付状态和不可变输入。`FIX-001` 继续为 BLOCKER。

### HIGH-001：批准 API 门禁仍失败

- net6/net8 全量 Unit 各 `372` 个，均为 `371` 成功、`1` 失败；唯一失败为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。
- Abstractions actual `5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646`，approved `7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`。
- Core actual `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`，approved `41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`。
- 四 TFM compare 对 netcoreapp3.1/net6/net7/net8 的 Abstractions/Core 均失败并返回退出码 `1`；Npoi 匹配。
- `04-api-breaking-changes.md` 仍为 `BLOCKED`；没有批准 hash 对应历史 DLL/成员快照、生成上下文或批准当前 actual 的记录。

`FIX-003` 继续阻止 RH31-203、RH31-401、RH31-404 和最终 Go。

### HIGH-002：性能没有批准判定合同

- 正式 BenchmarkDotNet、资源探针和两份 Round 4 tail-latency JSONL 均存在。
- 两份 tail-latency JSONL 各 25 行，对 `APPROVED/PASS/WAIVED` 的 `budgetStatus` 标记计数均为 `0`。
- `08-benchmark-report.md` 状态仍为 `PARTIAL`，明确缺少批准 baseline、环境预算、容差、批准人和 release waiver。
- 原始性能证据和报告仍未被 Git 跟踪。

测量结果可用于观察，但不能转化为发布 PASS。`FIX-004` 保持 HIGH/MUST_FIX。

### MEDIUM-001：selector 边界与客户端政策仍未闭环

- Round 10 未修改 selector 测试；搜索仍无明确 XLS/XLSX 越界 `ByIndex` 或 `SheetVisibility.VeryHidden` 用例。
- Integration net6/net8 各 `15/15` 通过，说明现有主链没有新回归，但不能替代缺失边界测试。
- 没有本轮 Excel/WPS/LibreOffice 版本、fixture hash、打开/保存/重开结果或具名 waiver。

`FIX-006` 继续为 SHOULD_FIX。

## 计划验收矩阵

| 计划项 | 状态 | 实际证据 |
| --- | --- | --- |
| RH31-000 / RH31-001 / RH31-602 | FAIL | 默认与隔离 strict restore/build 通过；要求文件仍未进入 Git，无实际待交付状态和不可变包身份复现。 |
| RH31-101 | PARTIAL | selector 单次解析主链未见回归；越界与 VeryHidden 直接测试仍缺。 |
| RH31-102 至 RH31-105 | PASS/PARTIAL | Release build 和 Integration 通过；最终发布仍被 API 门禁阻断。 |
| RH31-201 / RH31-202 | PARTIAL | public surface 可生成和分类；批准来源与版本决策未闭环。 |
| RH31-203 | FAIL | 四 TFM compare 可执行但 Abstractions/Core 不匹配，工具/baseline 未进入 Git。 |
| RH31-401 | FAIL | net6/net8 Unit 各 `371/372`，强制全绿条件未满足。 |
| RH31-402 | PASS | Integration net6/net8 各 `15/15`。 |
| RH31-403 | PASS/PARTIAL | package-only consumer `11/11`；Git 交付和不可变版本身份仍属于复现缺口。 |
| RH31-404 | FAIL | 多目标 Release build 成功；net6/net8 Unit/API 强制门禁失败。 |
| RH31-501 / RH31-502 | PARTIAL | 正式测量存在；批准预算/waiver 和证据交付缺失。 |
| RH31-601 | PARTIAL | 报告如实保持 No-Go；包输入不可变性仍需交付说明。 |
| RH31-603 | NOT_VERIFIABLE | 无办公客户端结果或具名 waiver。 |
| RH31-604 | FAIL | 仍有一个 BLOCKER、两个 HIGH 和一个 SHOULD_FIX。 |

## Git、功能与架构 Review

- Round 10 只修改 `execution.md` 运行记录和任务 runtime 状态，没有新增业务/测试行为变化。
- 当前已跟踪修改仍集中于 selector、API/benchmark 门禁、Docs consumer、文档和测试，属于原计划范围；未发现新的无关行为变化。
- 生产依赖方向保持 `Abstractions <- Core <- Npoi`；未发现第二套 API canonicalizer、NPOI public 类型泄漏或 production IVT 回归。
- API 工具与 Unit contract 继续共享 `PublicApiSnapshot.cs`；问题是批准合同和 Git 交付未闭环，不是工具不可执行。
- `get_errors` 对 `src`、`tests`、`build` 返回 0 个诊断；`git diff --check` 退出码 `0`，仅有 CRLF/LF 提示。

## 测试与包 Review

| 验证 | 本轮结果 |
| --- | --- |
| 默认缓存 locked restore | PASS |
| 全新隔离缓存 locked restore | PASS |
| 全新隔离缓存 Release solution build | PASS，314 个既有兼容/obsolete/analyzer 警告 |
| Unit net6 | FAIL：`371/372`，唯一失败为 API approved hash mismatch |
| Unit net8 | FAIL：`371/372`，唯一失败为 API approved hash mismatch |
| 四 TFM API compare | FAIL：Abstractions/Core mismatch，Npoi match，退出码 `1` |
| Integration net6 | PASS：`15/15` |
| Integration net8 | PASS：`15/15` |
| Docs package-only net8 | PASS：`11/11` |
| Docs assets | PASS：三个 Bing 依赖均为精确 `2.0.0`、`type=package`，SHA-512 与上一轮一致 |
| 性能批准 | FAIL/UNAPPROVED |
| Git tracking/clean clone | FAIL：要求文件跟踪数为 `0` |
| 编辑器诊断 | PASS：0 个错误 |
| `git diff --check` | PASS，仅 CRLF/LF 提示 |

本机仍缺 net7/netcoreapp3.1 runtime，因此这两个 TFM 只有 metadata-only API compare，不应声明 runtime test PASS。

## 文档与残余风险

- `execution.md` 将 Round 10 如实标为 `PARTIAL`，没有伪造 Git、API 或性能批准。
- 默认 restore 当前恢复通过，应在后续报告中与历史 `NU1403` 区分：当前命令成功，但同版本包曾被覆盖，版本身份风险仍存在。
- `04-api-breaking-changes.md` 和 `08-benchmark-report.md` 均保持 BLOCKED/PARTIAL，没有把缺失负责人决策写成 PASS。
- package-only consumer 证据有效；它证明当前本地 feed 内容可消费，不证明同一版本未来不会被替换。

## 结构化修复任务

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL/NOT_RESOLVED`
- 对应计划项：RH31-000、RH31-001、RH31-403、RH31-602、RH31-604。
- 涉及文件：10 个 `packages.lock.json`、`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、任务报告/快照/benchmark 原始证据、package consumer 复现说明。
- 问题：当前工作树可 strict restore/build，但要求文件不属于 Git 交付内容，也没有不可变 package version identity 或实际待交付提交/clean clone 证明。
- 证据：lockfile、任务目录和 API 文件跟踪数均为 `0`；默认与隔离 restore/build 当前通过；历史上相同 `2.0.0` 不同内容曾导致 `NU1403`。
- 影响：待交付提交无法取得锁图、API 门禁和发布证据；覆盖已消费版本会使缓存和 lock hash 失效。
- 修复目标：使要求文件和包身份成为可交付、不可变且可重复验证的输入。
- 明确修复要求：由获授权交付步骤纳管 lockfile、API 工具/baseline、00-11 报告、四 TFM 快照和批准保留的性能原始证据；禁止覆盖已消费的 `2.0.0`，使用唯一版本身份或受控不可变 feed；明确 Bing 本地 source 与第三方 source/预热 cache；在实际待交付状态执行 strict restore、Release build、API、Unit 和 Docs 门禁。
- 修复后的验证方式：`git ls-files` 列出全部要求文件；实际待交付提交或隔离 worktree 在全新缓存中 strict restore/build 成功；同一版本包哈希固定；Docs assets 精确解析该不可变包；报告和原始证据可由 Git 取得。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL/NOT_RESOLVED`
- 对应计划项：RH31-201、RH31-202、RH31-203、RH31-401、RH31-404。
- 涉及文件：`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、`04-api-breaking-changes.md`、四 TFM 快照/diff。
- 问题：共享 canonicalizer 和自动门禁可执行，但批准 baseline 与当前 Abstractions/Core 产物不一致。
- 证据：Unit net6/net8 各 `371/372`；四 TFM compare 对 Abstractions/Core 失败且退出码 `1`；Npoi 匹配；无批准历史成员快照或当前 actual 决策。
- 影响：无法确认当前 API 漂移是预期 2.0.0 surface 还是未批准 breaking，强制发布门禁失败。
- 修复目标：基于可追溯批准产物/规范使 API 门禁全绿，而不是机械替换 expected hash。
- 明确修复要求：版本负责人提供批准 hash 对应历史 DLL/成员快照与生成规范，或明确批准当前 actual；记录批准人、日期、版本、source/binary 影响和 migration 后同步 baseline；将工具、baseline、快照和 CI/验收入口纳入 Git。
- 修复后的验证方式：net6/net8 `PublicApiContractTest` 全绿；四 shipped TFM compare 通过；受治理签名变化会导致门禁失败；批准记录可追溯。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-501、RH31-502、RH31-604。
- 涉及文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`08-benchmark-report.md`、BDN/JSONL 原始产物和批准记录。
- 问题：性能测量没有批准 workload、预算、容差、批准人或具名 waiver，原始证据也未进入 Git。
- 证据：两份 tail-latency JSONL 的批准状态计数均为 `0`；报告仍为 `PARTIAL/UNAPPROVED`；任务证据跟踪数为 `0`。
- 影响：不能把性能结果判为 PASS，不满足发布 Go 条件。
- 修复目标：将现有测量绑定到具名批准的环境/workload/预算，或取得完整 release waiver，并交付原始证据。
- 明确修复要求：版本负责人记录冷/热 workload、p99/吞吐预算、允许波动、环境、批准人和日期，或签署包含范围与风险的具名 waiver；报告逐项标记 `PASS`/`FAIL`/`WAIVED`；原始 JSONL/BDN 结果进入 Git。
- 修复后的验证方式：相同环境重复运行可按批准条件判定；报告和原始数据包含批准身份、范围、日期和结论，Git 可取得证据。

### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-101、RH31-403、RH31-603。
- 涉及文件：`tests/Bing.Offices.Tests/StreamPipelineTest.cs`、互操作报告与 fixture 证据。
- 问题：selector 主链未回归，但越界索引、VeryHidden 和客户端政策仍未闭环。
- 证据：无明确 XLS/XLSX 越界 `ByIndex`、`SheetVisibility.VeryHidden` 测试；无客户端结果或批准 waiver。
- 影响：明确边界分支缺少直接自动回归，P1 客户端兼容风险未按发布政策处理。
- 修复目标：补齐 XLS/XLSX selector 边界，并提供客户端证据或具名 waiver。
- 明确修复要求：增加 XLS/XLSX 越界 `ByIndex` 和 `SheetVisibility.VeryHidden` 用例，断言 `InvalidHeader`、稳定 `SheetName` 且 plan 不执行；运行可用客户端打开/保存/重开，或取得包含批准人、范围、日期和风险的 waiver。
- 修复后的验证方式：新增测试在 net6/net8 通过；互操作报告包含客户端版本、generation/fixture/roundtrip hash 和结果，或完整 waiver。

## 最终 Checklist

- [x] 已优先逐项复验上一轮 FIX。
- [x] 默认与隔离 strict restore、Release build、Unit、API compare、Integration 和 package-only Docs 均已独立运行。
- [x] 修正默认 restore 状态：本轮当前为 PASS，历史 `NU1403` 仅作为不可变版本风险证据。
- [x] `FIX-007` 和 `FIX-008` 保持 RESOLVED，未发现新回归。
- [x] `get_errors` 与 `git diff --check` 通过。
- [ ] lockfile、API 工具和发布证据进入 Git并完成实际待交付状态复现。
- [ ] net6/net8 API contract、Unit 和四 TFM compare 全绿并有批准依据。
- [ ] 性能获得具名预算/批准或 waiver，原始结果可交付。
- [ ] 越界 `ByIndex`、VeryHidden 和客户端政策闭环。
- [x] 未执行 Git staging、commit、push、tag、publish、PR 或外部通知。

## 结论

`NEEDS_FIX`。Round 10 没有实现变更或新增回归，默认和隔离 restore/build 当前均通过；但 `FIX-001` Git 交付与不可变包身份 BLOCKER、`FIX-003` API HIGH、`FIX-004` 性能 HIGH 和 `FIX-006` SHOULD_FIX 仍未关闭。`FIX-007`、`FIX-008` 继续保持解决状态。后续应先完成获授权的 Git 纳管和版本负责人批准，再重新独立 Review。

---

# Review Fix Round 9 独立复审

## 验收摘要

本次复审以 `plan.md`、Round 9 `execution.md`、上一轮 `review.md`、当前源码、Git Diff、测试、包、API 和性能原始证据为依据，优先逐项验证上一轮 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`，并确认已解决的 `FIX-007`、`FIX-008` 未回归。Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

最终结论：`NEEDS_FIX`，发布判定保持 `No-Go`。

- `FIX-001`：`NOT_RESOLVED`。当前 10 个 lockfile、API 工具/baseline 和整个任务证据目录的 Git 跟踪数仍为 `0`；默认 locked restore 稳定失败于三个同版本 `2.0.0` 包的 `NU1403`。全新隔离缓存 restore/build 成功，只能证明 dirty 工作树在明确包输入下可用，不能证明待交付提交或 clean clone 可复现。
- `FIX-003`：`PARTIAL/NOT_RESOLVED`。net6/net8 Unit 各 `371/372`；四 TFM API compare 对 Abstractions/Core 全部失败，Npoi 匹配。Round 9 没有取得批准 baseline 来源或版本负责人决策。
- `FIX-004`：`NOT_RESOLVED`。正式 BDN、资源探针和 tail-latency 数据存在，但尾延迟证据仍无批准标记；报告仍缺批准预算、批准人、日期或具名 waiver，证据也未进入 Git。
- `FIX-006`：`NOT_RESOLVED`。Round 9 按 `fixScope=must` 未处理；仍无 XLS/XLSX 越界 `ByIndex`、`SheetVisibility.VeryHidden` 直接测试，也无客户端互操作结果或具名 waiver。
- `FIX-007`：`RESOLVED`。全新隔离缓存下 package-only Docs restore/build/test 为 `11/11`；三个 Bing 依赖均为精确 `2.0.0` 且 `type=package`。
- `FIX-008`：`RESOLVED`。Unit csproj 仍是合法非空 XML；隔离 solution build 成功，Unit/API 门禁可执行。当前唯一 Unit 失败属于 `FIX-003`。

## 上一轮 FIX 复审矩阵

| FIX | Round 9 执行状态 | 本轮复审状态 | 结论 |
| --- | --- | --- | --- |
| `FIX-001` | PARTIAL/BLOCKED | NOT_RESOLVED | 隔离 dirty 工作树可构建；Git 交付、clean clone 和不可变包身份仍未成立。 |
| `FIX-003` | PARTIAL/BLOCKED | PARTIAL/NOT_RESOLVED | 自动门禁稳定可执行，但批准 API 比较仍失败。 |
| `FIX-004` | PARTIAL/BLOCKED | NOT_RESOLVED | 测量存在，批准预算/waiver 和 Git 交付仍缺。 |
| `FIX-006` | 未处理（must scope） | NOT_RESOLVED | selector 明确边界与客户端政策没有变化。 |
| `FIX-007` | 历史 COMPLETED | RESOLVED | package-only consumer 在全新缓存中独立通过。 |
| `FIX-008` | 历史 COMPLETED | RESOLVED | Unit 项目、构建和运行入口保持有效。 |

## 主要发现

### BLOCKER-001：要求文件仍不属于 Git 交付面，包输入仍非不可变

- 当前分支为 `master`，HEAD 为 `1968b24a3ab07b44c3b386a3f761fcdff2fc4315`。
- 工作树存在 10 个 `packages.lock.json`，但 `git ls-files -- '*packages.lock.json'` 数量为 `0`；任务目录和 `build/ApiSnapshot/**`/baseline 的跟踪数也均为 `0`。
- `git status --short --untracked-files=all` 继续将任务报告、API 快照、benchmark 原始证据、lockfile 和 API 工具列为 `??`。
- 默认 `dotnet restore Bing.Offices.sln --locked-mode --no-cache` 失败，Docs consumer 的 Abstractions/Core/Npoi 三个 `2.0.0` 包均报 `NU1403` 内容哈希不同。
- 全新隔离 `NUGET_PACKAGES` 与 HTTP cache 下，strict restore 和 Release solution build 成功；该结果证明锁图与本轮本地包可以配套工作，但不能证明同版本包被覆盖后的默认缓存可重复，也不能替代实际待交付提交验证。

`FIX-001` 继续是发布 BLOCKER。Reviewer 无 Git staging/提交授权不是验收豁免。

### HIGH-001：API 批准门禁仍失败

- net6/net8 全量 Unit 各运行 `372` 个，均为 `371` 成功、`1` 失败；唯一失败是 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。
- Abstractions actual 为 `5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646`，approved 为 `7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`。
- 四 TFM compare 中 Core actual 为 `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`，approved 为 `41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`；Abstractions/Core 在 netcoreapp3.1、net6、net7、net8 均失败，工具退出码 `1`。
- Npoi 保持匹配；实际 hash 在四个 TFM 间稳定，因此当前失败不是随机 runtime 差异。
- 现有 baseline 只有批准 hash，没有对应历史成员快照、生成上下文或批准记录。Round 9 正确地没有机械替换 expected hash，但也未关闭发布门禁。

`FIX-003` 继续阻止 RH31-203、RH31-401、RH31-404 和最终 Go。

### HIGH-002：性能测量仍没有批准判定合同

- 正式 benchmark、资源探针和两份 Round 4 tail-latency JSONL 仍存在；tail-latency 覆盖并发 `1/4/16/64`、预热、5 轮及 p50/p95/p99/吞吐。
- 两份 tail-latency JSONL 没有批准标记；`08-benchmark-report.md` 仍未提供获批准的 workload、p99/吞吐预算、容差、批准人和日期，也没有具名 release waiver。
- 原始 BDN/JSONL 和报告仍未被 Git 跟踪。

现有数据可用于观察，不能据此把性能门禁判为 PASS。`FIX-004` 保持 HIGH/MUST_FIX。

### MEDIUM-001：selector 边界与客户端政策仍未闭环

- 生产主链仍复用一次解析后的 physical sheet 信息，Integration net6/net8 各 `15/15` 通过，未发现 Round 9 引入主链回归。
- 当前测试搜索仍无明确 XLS/XLSX 越界 `ByIndex` 或 `SheetVisibility.VeryHidden` 用例。
- 没有 Excel/WPS/LibreOffice 本轮版本、fixture hash、打开/保存/重开结果，也没有具名批准 waiver。

`FIX-006` 继续为 SHOULD_FIX；本轮 must scope 不改变发布计划对该项的要求。

## 计划验收矩阵

| 计划项 | 状态 | 实际证据 |
| --- | --- | --- |
| RH31-000 / RH31-001 / RH31-602 | FAIL | 隔离工作树可 restore/build，但要求文件未进入 Git，默认缓存因同版本包冲突报 `NU1403`，无实际交付状态复现。 |
| RH31-101 | PARTIAL | selector 单次解析主链未回归；越界与 VeryHidden 直接测试仍缺。 |
| RH31-102 至 RH31-105 | PASS/PARTIAL | Release solution build 与 Integration 通过；最终发布仍被 API 门禁阻断。 |
| RH31-201 / RH31-202 | PARTIAL | 当前 public surface 可生成，但批准来源和版本决策未闭环。 |
| RH31-203 | FAIL | 四 TFM compare 可执行但 Abstractions/Core 不匹配，工具/baseline 未进入 Git。 |
| RH31-401 | FAIL | net6/net8 Unit 各 `371/372`，强制全绿条件未满足。 |
| RH31-402 | PASS | Integration net6/net8 各 `15/15`。 |
| RH31-403 | PASS/PARTIAL | package-only consumer 在全新缓存中 `11/11`；同版本可变包和 Git 交付仍是复现风险。 |
| RH31-404 | FAIL | 多目标 Release build 成功；net6/net8 Unit/API 强制门禁失败。 |
| RH31-501 / RH31-502 | PARTIAL | 正式测量存在；批准预算/waiver 和证据交付缺失。 |
| RH31-601 | PARTIAL | 报告保持 No-Go；发布复现说明仍需约束不可变包输入。 |
| RH31-603 | NOT_VERIFIABLE | 无办公客户端结果或具名 waiver。 |
| RH31-604 | FAIL | 仍有一个 BLOCKER、两个 HIGH 和一个 SHOULD_FIX。 |

## Git、功能与架构 Review

- 当前已跟踪修改集中于 selector、API/benchmark 门禁、Docs consumer、文档和测试，属于原计划范围；未发现 Round 9 新增的业务实现变更。
- lockfile、API 工具/baseline、任务报告、快照和性能证据全部未跟踪。这是交付完整性问题，不是可忽略的 Reviewer 临时文件状态。
- 生产依赖方向仍为 `Abstractions <- Core <- Npoi`；未发现新的 NPOI public 类型泄漏或第二套 API canonicalizer。
- API 工具与 Unit contract 继续共享 `PublicApiSnapshot.cs`；自动化入口本身可用，失败点是批准合同未闭环。
- `get_errors` 对 `src`、`tests`、`build` 返回 0 个诊断；`git diff --check` 退出码 `0`，只有 CRLF/LF 提示。

## 测试与包 Review

| 验证 | 本轮结果 |
| --- | --- |
| 默认缓存 locked restore | FAIL：三个 Bing `2.0.0` 包 `NU1403` |
| 全新隔离缓存 locked restore | PASS |
| 全新隔离缓存 Release solution build | PASS，314 个既有兼容/obsolete/analyzer 警告 |
| Unit net6 | FAIL：`371/372`，唯一失败为 API approved hash mismatch |
| Unit net8 | FAIL：`371/372`，唯一失败为 API approved hash mismatch |
| 四 TFM API compare | FAIL：Abstractions/Core mismatch，Npoi match，退出码 `1` |
| Integration net6 | PASS：`15/15` |
| Integration net8 | PASS：`15/15` |
| Docs package-only net8 | PASS：`11/11` |
| Docs assets | PASS：三个 Bing 依赖均为精确 `2.0.0`、`type=package` |
| 性能批准 | FAIL/UNAPPROVED |
| Git tracking/clean clone | FAIL：要求文件跟踪数为 `0` |
| 编辑器诊断 | PASS：0 个错误 |
| `git diff --check` | PASS，仅 CRLF/LF 提示 |

本机仍没有 net7/netcoreapp3.1 runtime，因此这两个 TFM 只有 metadata-only API 比较证据，不应声明 runtime test PASS。

## 文档与残余风险

- `execution.md` 对 Round 9 的 `PARTIAL/BLOCKED` 判定与实际验证一致，没有伪造 Git、API 或性能批准。
- package-only consumer 的隔离验证有效，但相同 `2.0.0` 在本地 feed 和既有全局缓存中内容不同。版本身份不可变性必须在交付流程中解决，不能依赖清缓存作为发布合同。
- API baseline 若没有历史成员快照，只能由版本负责人选择提供原批准产物或批准当前 surface；Reviewer 无法从 hash 反推出批准差异。
- 性能和办公客户端互操作均涉及外部批准/环境证据。缺失时必须保持 No-Go/NOT_VERIFIABLE，不能转换成隐式 PASS。

## 结构化修复任务

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-000、RH31-001、RH31-403、RH31-602、RH31-604。
- 涉及文件：10 个 `packages.lock.json`、`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、当前任务报告/快照/benchmark 原始证据、package consumer 复现说明。
- 问题：要求文件未进入 Git，默认缓存因相同版本不同包内容报 `NU1403`，没有实际待交付提交/clean clone 的不可变输入证明。
- 证据：lockfile、任务目录和 API 文件 Git 跟踪数均为 `0`；默认 locked restore 失败；隔离 restore/build 通过。
- 影响：待交付提交无法取得锁图、API 门禁和发布证据，依赖复现受本机缓存与可变同版本包影响。
- 修复目标：使要求文件和包身份成为可交付、可重复验证的输入。
- 明确修复要求：由获授权交付步骤纳管 lockfile、API 工具/baseline、00-11 报告、四 TFM 快照和批准保留的原始性能证据；禁止覆盖已消费的 `2.0.0`，使用不可变包内容或唯一版本身份；明确本地 Bing source 与第三方 source/预热 cache；在实际待交付状态执行 strict restore、Release build、API、Unit 和 Docs 门禁。
- 修复后的验证方式：`git ls-files` 可列出全部要求文件；实际待交付提交或隔离 worktree 在全新缓存中 strict restore/build 成功；Docs assets 精确解析不可变包；报告和原始证据可由 Git 取得。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL/NOT_RESOLVED`
- 对应计划项：RH31-201、RH31-202、RH31-203、RH31-401、RH31-404。
- 涉及文件：`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、`04-api-breaking-changes.md`、四 TFM 快照/diff。
- 问题：共享 canonicalizer 和自动门禁可执行，但批准 baseline 与当前 Abstractions/Core 产物不一致。
- 证据：Unit net6/net8 各 `371/372`；四 TFM compare 对 Abstractions/Core 失败且退出码 `1`；Npoi 匹配。
- 影响：无法判断当前 API 漂移是预期 2.0.0 surface 还是未批准 breaking，强制发布门禁失败。
- 修复目标：基于可追溯批准产物/规范使 API 门禁全绿，而不是机械替换 expected hash。
- 明确修复要求：版本负责人提供批准 hash 对应历史 DLL/成员快照与生成规范，或明确批准当前 actual；记录批准人、日期、版本、source/binary 影响和 migration 后同步 baseline；将工具、baseline、快照和 CI/验收入口纳入 Git。
- 修复后的验证方式：net6/net8 `PublicApiContractTest` 全绿；四 shipped TFM 自动 compare 通过；受治理签名变化会导致门禁失败；批准记录可追溯。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-501、RH31-502、RH31-604。
- 涉及文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`08-benchmark-report.md`、BDN/JSONL 原始产物和批准记录。
- 问题：性能测量没有批准 workload、预算、容差、批准人或具名 waiver，原始证据也未进入 Git。
- 证据：tail-latency 证据无批准标记；报告仍为 `UNAPPROVED`；要求文件跟踪数为 `0`。
- 影响：不能把性能结果判为 PASS，不满足发布 Go 条件。
- 修复目标：将现有测量绑定到具名批准的环境/workload/预算，或取得完整 release waiver，并交付原始证据。
- 明确修复要求：版本负责人记录冷/热 workload、p99/吞吐预算、允许波动、环境、批准人和日期，或签署包含范围与风险的具名 waiver；报告逐项标记 `PASS`/`FAIL`/`WAIVED`；原始 JSONL/BDN 结果进入 Git。
- 修复后的验证方式：相同环境重复运行可按批准条件判定；报告和原始数据包含批准身份、范围、日期和结论，Git 可取得证据。

### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-101、RH31-403、RH31-603。
- 涉及文件：`tests/Bing.Offices.Tests/StreamPipelineTest.cs`、互操作报告与 fixture 证据。
- 问题：selector 主链未回归，但越界索引、VeryHidden 和客户端政策仍未闭环。
- 证据：无明确 XLS/XLSX 越界 `ByIndex`、`SheetVisibility.VeryHidden` 测试；无客户端结果或批准 waiver。
- 影响：明确边界分支缺少直接自动回归，P1 客户端兼容风险未按发布政策处理。
- 修复目标：补齐 XLS/XLSX selector 边界，并提供客户端证据或具名 waiver。
- 明确修复要求：增加 XLS/XLSX 越界 `ByIndex` 和 `SheetVisibility.VeryHidden` 用例，断言 `InvalidHeader`、稳定 `SheetName` 且 plan 不执行；运行可用客户端打开/保存/重开，或取得包含批准人、范围、日期和风险的 waiver。
- 修复后的验证方式：新增测试在 net6/net8 通过；互操作报告包含客户端版本、generation/fixture/roundtrip hash 和结果，或完整 waiver。

## 最终 Checklist

- [x] 已优先逐项复验上一轮 FIX。
- [x] 已独立运行默认及隔离 locked restore、Release build、Unit、API compare、Integration 和 package-only Docs。
- [x] `FIX-007` 和 `FIX-008` 保持 RESOLVED，未发现新回归。
- [x] `get_errors` 与 `git diff --check` 通过。
- [ ] lockfile、API 工具和发布证据进入 Git并完成实际待交付状态复现。
- [ ] net6/net8 API contract、Unit 和四 TFM compare 全绿并有批准依据。
- [ ] 性能获得具名预算/批准或 waiver，原始结果可交付。
- [ ] 越界 `ByIndex`、VeryHidden 和客户端政策闭环。
- [x] 未执行 Git staging、commit、push、tag、publish、PR 或外部通知。

## 结论

`NEEDS_FIX`。Round 9 没有引入新的业务或测试回归，并正确保留了需要外部授权/批准的阻塞项；但 `FIX-001` Git 交付 BLOCKER、`FIX-003` API HIGH、`FIX-004` 性能 HIGH 和 `FIX-006` SHOULD_FIX 均未关闭。`FIX-007`、`FIX-008` 继续保持解决状态。下一轮应由获授权交付步骤和版本负责人先完成不可变包/Git 纳管、API 批准及性能批准，再交由独立 Reviewer 复验。

---

# Review Fix Round 8 独立复审

## Round 8 验收摘要

本次复审以当前 `plan.md`、Round 8 `execution.md`、上一轮 `review.md`、实际源码、Git Diff、锁图、包、测试和性能证据为依据，优先逐项验证上一轮 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`、`FIX-007`、`FIX-008`。Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

最终结论：`NEEDS_FIX`，发布判定继续为 `No-Go`。

- `FIX-001`：`NOT_RESOLVED`。全新隔离缓存下 locked restore 和 Release solution build 通过，但 10 个 lockfile、API 工具/baseline、任务报告、快照和性能原始证据仍未被 Git 跟踪，不能形成实际待交付提交或 clean-clone 证明。默认全局缓存还因同版本旧包内容不同而报 `NU1403`，进一步说明交付验证必须使用明确隔离缓存和不可变包输入。
- `FIX-003`：`PARTIAL/NOT_RESOLVED`。Unit/API runtime gate 已恢复可执行，但 net6/net8 API contract 均为 `6/7`；四 TFM 自动 compare 对 Abstractions/Core 全部失败，Npoi 匹配。没有批准当前 actual hash 的版本决策。
- `FIX-004`：`NOT_RESOLVED`。BDN 和 tail-latency 原始结果存在，但报告与 JSONL 仍为 `budgetStatus=UNAPPROVED`；没有批准预算、批准人、日期或具名 waiver，证据也未进入 Git。
- `FIX-006`：`NOT_RESOLVED`。本轮按 `fixScope=must` 跳过；搜索仍无 XLS/XLSX 越界 `ByIndex` 或 `SheetVisibility.VeryHidden` 直接测试，也没有办公客户端结果或具名 waiver。
- `FIX-007`：`RESOLVED`。Docs consumer 无生产 `ProjectReference`；全新缓存 restore/build/test 通过 `11/11`，三个 Bing 依赖均为精确 `2.0.0`、`type=package`，本轮 nupkg 含 DLL、XML、nuspec、LICENSE 和 README；Abstractions 包实际包含 metadata 类型与方法。
- `FIX-008`：`RESOLVED`。Unit csproj 当前为 2849 字节、唯一合法 `Project` 根节点；locked restore/build 在隔离缓存中通过，net6/net8 Unit、API、selector 均已恢复为可执行门禁。全量 Unit 的唯一失败属于 `FIX-003`，不是项目文件损坏。

## Round 8 上一轮 FIX 复审矩阵

| FIX | Round 8 执行状态 | 本轮复审状态 | 结论 |
| --- | --- | --- | --- |
| `FIX-001` | PARTIAL/BLOCKED | NOT_RESOLVED | dirty 工作树的隔离 locked restore 成立；Git 交付、clean clone 和不可变包输入仍未成立。 |
| `FIX-003` | PARTIAL | PARTIAL/NOT_RESOLVED | runtime gate 已恢复，但批准 API 比较仍失败。 |
| `FIX-004` | PARTIAL/BLOCKED | NOT_RESOLVED | 测量存在，批准预算/waiver 和 Git 交付仍缺。 |
| `FIX-006` | SKIPPED/OUT_OF_SCOPE | NOT_RESOLVED | VeryHidden、越界索引和客户端政策仍未闭环。 |
| `FIX-007` | COMPLETED | RESOLVED | package-only consumer、assets 和 nupkg 内容均独立通过。 |
| `FIX-008` | COMPLETED | RESOLVED | Unit 项目恢复，构建和测试入口已恢复。 |

## Round 8 主要发现

### BLOCKER-001：Git 交付和 clean-clone 复现仍未成立

- 分支为 `master`，HEAD 为 `1968b24a3ab07b44c3b386a3f761fcdff2fc4315`。
- `git ls-files -- '*packages.lock.json'` 无输出；当前任务目录被跟踪文件数为 `0`；`build/ApiSnapshot/**` 和 `build/api-snapshot-baseline.json` 也未被跟踪。
- Git status 仍将 lockfile、API 工具/baseline、00-11 报告、四 TFM 快照和 BDN/JSONL 证据列为 `??`。
- `artifacts/benchmarks/*` 仍被顶层 `artifacts/` ignore；任务目录中的副本不再被忽略，但同样未被 Git 跟踪。
- 全新 `NUGET_PACKAGES` 与 HTTP cache 中，`dotnet restore Bing.Offices.sln --locked-mode --no-cache` 和 Release solution build 均通过；但默认全局缓存还原报 `NU1403`，因为缓存中的 `Bing.Offices.Abstractions.2.0.0.nupkg` SHA-256 为 `024B34F1...`，本轮本地 feed 同版本包为 `FB603C46...`。

当前结果证明本轮包和锁图可在明确隔离输入下工作，但不能证明待交付提交可取得这些文件，也不能把同版本包被覆盖后的本机缓存当作可重复发布模型。`FIX-001` 继续为 BLOCKER。

### HIGH-001：API 自动门禁可执行，但批准比较仍失败

- API snapshot 工具 Release build 通过，并自动处理 `netcoreapp3.1`、`net6.0`、`net7.0`、`net8.0`。
- 四个 TFM 均得到相同 actual：Abstractions `5A1B668E14A3A2689A0CC88BB95F3396025BABC1AF18A6A5A64C0BD0DF290646`，批准值 `7B0BA2792AE1DB91BB281C1719B0B35671091CA981C659FDC89B3771B7F5F577`；Core `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`，批准值 `41B6D12CD58A988E84701902E0F58476B33903583A51F39DF7544B436504DF54`。
- Npoi `A0DBE980...` 匹配批准值；工具退出码为 `1`。
- net6/net8 `PublicApiContractTest` 各 `6/7`，全量 Unit 各 `371/372`；唯一失败均为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。
- `04-api-breaking-changes.md` 如实标记 `BLOCKED`，未发现批准当前 actual、批准人、日期、版本或 migration 决策。

`FIX-003` 仍阻止 RH31-203、RH31-401、RH31-404 和最终 Go；不得机械替换 baseline。

### HIGH-002：性能仍无批准预算或 waiver

- 六组 BenchmarkDotNet 正式结果、资源探针和两次 Round 4 tail-latency JSONL 均存在。
- 两次 tail-latency 覆盖并发 `1/4/16/64`、预热、5 轮、每档 `1280` 样本，并保留 p50/p95/p99、吞吐和环境身份。
- 报告和 JSONL 均明确 `budgetStatus=UNAPPROVED`；没有批准 workload、baseline、p99/吞吐预算、容差、批准人、日期或 release waiver。
- 任务目录证据仍未进入 Git；`artifacts/benchmarks` 原始位置还受 ignore 规则影响。

测量技术证据可用于观察，但不能转换为发布 PASS。`FIX-004` 保持 HIGH/MUST_FIX。

### MEDIUM-001：selector 明确边界和客户端政策仍缺失

- 生产主链仍只解析 selector 一次，并将 `NpoiResolvedSheet` 复用于 mapping plan、materialization 和 Failure Workbook mapping；未发现第二套主流程回归。
- `StreamPipelineTest` net6/net8 各 `90/90` 通过，说明现有 selector/导入场景可执行。
- 测试搜索仍只有合法 `ByIndex(0/1)` 和普通 `SheetVisibility.Hidden`；没有明确越界索引或 `SheetVisibility.VeryHidden` 用例。
- 没有 Excel/WPS/LibreOffice 本轮版本、fixture hash、打开/保存/重开结果，也没有具名批准 waiver。

`FIX-006` 未因 must scope 跳过而获得发布豁免，继续为 SHOULD_FIX。

## Round 8 实际验证

| 验证 | 结果 |
| --- | --- |
| Unit csproj 磁盘/XML | PASS：2849 字节，唯一 `Project` 根节点 |
| 默认缓存 locked restore | FAIL：Docs 三个 `2.0.0` 包内容哈希与全局旧缓存不同，`NU1403` |
| 全新隔离缓存 locked restore | PASS |
| 全新隔离缓存 Release solution build | PASS，退出码 0 |
| Unit net6/net8 | FAIL，各 `371/372`；唯一失败为批准 API hash mismatch |
| API contract net6/net8 | FAIL，各 `6/7`；同一 Abstractions mismatch |
| StreamPipeline net6/net8 | PASS，各 `90/90` |
| Integration net6/net8 | PASS，各 `15/15` |
| Docs package-only net8 | PASS，`11/11` |
| Docs assets | PASS：三个 Bing 包均为精确 `2.0.0`、`type=package` |
| nupkg 内容 | PASS：DLL、XML、nuspec、LICENSE、README；metadata 类型/方法存在 |
| API snapshot tool build | PASS |
| 四 TFM API compare | FAIL：Abstractions/Core mismatch；Npoi match；退出码 1 |
| selector 越界/VeryHidden | FAIL/NOT_VERIFIABLE：无直接测试 |
| 性能批准 | FAIL/UNAPPROVED |
| Git tracking/clean clone | FAIL：要求文件未跟踪，未形成待交付提交复现 |
| `get_errors` | PASS：相关项目、API 工具和生产文件无诊断错误 |
| `git diff --check` | PASS，仅有 CRLF/LF 提示 |

## Round 8 计划验收矩阵

| 计划项 | 状态 | 说明 |
| --- | --- | --- |
| RH31-000 / RH31-001 / RH31-602 | FAIL | 隔离工作树可构建，但要求文件不属于 Git 交付内容，clean clone 未验证。 |
| RH31-101 | PARTIAL | 单次解析主链和现有测试通过；越界与 VeryHidden 直接测试仍缺。 |
| RH31-102 至 RH31-105 | PASS/PARTIAL | Unit/Integration 主链恢复可执行，未发现新的 IO/资源/metadata 回归；完整发布仍被 API 门禁阻断。 |
| RH31-201 / RH31-202 | PARTIAL | 类型/成员治理和迁移账存在；批准 public surface 尚未确认。 |
| RH31-203 | FAIL | 四 TFM 自动入口存在，但批准比较不通过且工具/baseline 未进入 Git。 |
| RH31-401 | FAIL | Unit 可执行但 net6/net8 均非全绿。 |
| RH31-402 | PASS | Integration net6/net8 各 `15/15`。 |
| RH31-403 | PASS/PARTIAL | package-only consumer 在全新缓存中 `11/11`；默认旧缓存存在同版本内容冲突，交付说明必须坚持隔离/不可变输入。 |
| RH31-404 | FAIL | 多目标构建可用；强制 net6/net8 Unit 因 API 门禁失败。 |
| RH31-501 / RH31-502 | PARTIAL | 正式测量存在；预算、批准/waiver 和 Git 交付缺失。 |
| RH31-601 | PARTIAL | 报告总体如实保持 No-Go；package 报告未记录默认缓存 `NU1403` 风险。 |
| RH31-603 | NOT_VERIFIABLE | 无客户端结果或具名 waiver。 |
| RH31-604 | FAIL | 仍有一个 BLOCKER、两个 HIGH 和一个 SHOULD_FIX。 |

## Round 8 Git、功能与架构 Review

- 当前 17 个已跟踪文件有修改，主要集中于 selector 单次解析、API/benchmark 门禁、Docs package consumer、测试和发布文档，属于计划范围；未发现新的无关行为变更。
- 任务报告、API 工具/baseline、lockfile 和原始证据仍全部未跟踪，属于交付完整性缺口，而不是可忽略的 Reviewer 工作区状态。
- `NpoiExcelImporter` 在 Workbook 打开后一次解析全部 selector，先拒绝同物理 Sheet 冲突，再把 resolved index/name 传给 `NpoiImportPlanBuilder`、导入循环和 Failure Workbook；未发现重新调用 selector 的第二条主链。
- 生产依赖方向保持 `Abstractions <- Core <- Npoi`；未发现新增 NPOI public 类型泄漏或 production IVT。
- API 工具和 Unit contract 链接同一个 `PublicApiSnapshot.cs`，没有第二套 canonicalizer；当前问题是批准产物/规范未闭环，不是自动入口缺失。
- Unit csproj 的空文件回归已消失；`FIX-008` 不再阻断其它门禁。

## Round 8 测试、性能与文档 Review

- net6/net8 可用 runtime 均已实际运行；本机仍没有 net7/netcoreapp3.1 runtime，因此这两个 TFM 只有静态 API 工具证据，不应声明 runtime PASS。
- Docs consumer 覆盖 DI、provider-neutral API、mapping、CSV、metadata XLS/XLSX 重开和 docs fence；当前 `11/11` 是有效 package-only 证据。
- 本轮本地 `2.0.0` 包与全局缓存中旧 `2.0.0` 内容不同。隔离缓存验证方法正确，但发布流程应保持包版本不可变，避免以覆盖同版本包作为常规恢复策略。
- 性能报告没有夸大 zero-GC、真实 streaming 或完整压缩炸弹防护，并如实保留 100K 高分配和高并发波动；缺口仍是批准政策而不是测量代码。
- `07-package-consumer-report.md` 状态为 `VERIFIED`，但 Round 6 命令只显式列出 nuget.org，依赖项目的 `RestoreAdditionalProjectSources` 隐式加入本地 feed；报告还未记录默认全局缓存 `NU1403`。不影响本轮全新缓存 PASS，但降低复现说明的清晰度，纳入 `FIX-001` 的交付文档要求。

## Round 8 结构化修复任务

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-000、RH31-001、RH31-403、RH31-602、RH31-604。
- 涉及文件：10 个 `packages.lock.json`、`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、当前任务报告/快照/benchmark 证据、`.gitignore`、package consumer 复现说明。
- 问题：当前隔离 dirty 工作树可还原和构建，但要求文件仍不属于 Git 交付内容；默认缓存还因被覆盖的同版本本地包产生 `NU1403`，没有实际待交付提交/clean clone 的不可变输入证明。
- 证据：`git ls-files` 对 lockfile、任务目录和 API 工具无输出；Git status 全部为 `??`；默认 restore 报三个包 content hash mismatch；全新缓存 restore/build 通过。
- 影响：交付提交无法获得锁图、API 门禁、报告和性能证据；复现结果依赖未声明的本地包缓存状态。
- 修复目标：使全部要求文件可由 Git 交付，并在实际待交付提交或隔离 worktree 中使用明确、不可变的包输入完成复现。
- 明确修复要求：由获授权交付步骤纳管 lockfile、API 工具/baseline、00-11 报告、四 TFM 快照和批准保留的 BDN/JSONL 证据；明确本地 Bing feed 与第三方 source；不要覆盖已被消费的同版本包，或在本地验证流程中始终使用唯一隔离缓存/唯一版本身份；在实际待交付状态执行 locked restore、Release build、API/Unit/Docs 门禁。
- 修复后的验证方式：`git ls-files` 列出全部要求文件；隔离 worktree/实际待交付提交在全新缓存中 locked restore 与 Release build 退出码 0；Docs assets 精确解析本轮不可变包；报告和原始证据可由 Git 取得。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL/NOT_RESOLVED`
- 对应计划项：RH31-201、RH31-202、RH31-203、RH31-401、RH31-404。
- 涉及文件：`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、`04-api-breaking-changes.md`、四 TFM 快照/diff。
- 问题：自动 canonical 门禁和 runtime gate 已恢复，但批准 baseline 与当前 Abstractions/Core 产物不一致，四 TFM compare 和 net6/net8 Unit 均失败。
- 证据：Abstractions actual `5A1B668E...` vs approved `7B0BA279...`；Core actual `5F68499B...` vs approved `41B6D12...`；Npoi match；API contract 各 `6/7`；Unit 各 `371/372`；工具退出码 1。
- 影响：无法确认当前 API 漂移是预期 2.0.0 surface 还是未批准 breaking，强制发布门禁不通过。
- 修复目标：由可追溯、获批准的产物和规范得到全绿 API 门禁，而不是机械替换 expected hash。
- 明确修复要求：定位批准 hash 对应的历史产物/规范，生成 actual-vs-approved member diff；交由版本负责人决定恢复产物或批准新 baseline。若批准变更，记录批准人、日期、版本、source/binary 影响和 migration 后再同步 baseline；将工具、baseline、快照和 CI/验收入口纳入 Git。
- 修复后的验证方式：net6/net8 `PublicApiContractTest` 全绿；自动命令对四 shipped TFM 比较通过；任意受治理签名变化会失败；隔离交付状态可重复运行。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-501、RH31-502、RH31-604。
- 涉及文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`08-benchmark-report.md`、BDN/JSONL 原始产物和批准记录。
- 问题：可靠测量仍没有批准 workload、预算、容差、批准人或具名 waiver，原始证据也未进入 Git。
- 证据：报告和 JSONL 均为 `budgetStatus=UNAPPROVED`；未发现新的批准身份或 waiver；任务证据均为未跟踪文件。
- 影响：不能将性能结果判定为 PASS，不满足发布 Go 条件。
- 修复目标：将现有测量绑定到具名批准的环境/workload/预算，或取得完整 release waiver，并交付原始证据。
- 明确修复要求：由版本负责人记录冷/热 workload、p99/吞吐预算、允许波动、环境、批准人和日期，或签署带范围和风险的具名 waiver；报告逐项标记 PASS/FAIL/WAIVED；原始 JSONL/BDN 结果进入 Git。无需重写已经满足技术口径的测量实现。
- 修复后的验证方式：相同环境重复运行可按批准条件判定；报告和原始数据包含批准身份、范围、日期和结论，Git 可取得证据。

### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-101、RH31-403、RH31-603。
- 涉及文件：`tests/Bing.Offices.Tests/StreamPipelineTest.cs`、互操作报告与 fixture 证据。
- 问题：现有 selector 主链和 90 个 StreamPipeline 场景通过，但越界索引、VeryHidden 和客户端政策仍未闭环。
- 证据：搜索只有合法 `ByIndex(0/1)` 和普通 Hidden；无 VeryHidden、明确越界、客户端结果或批准 waiver。
- 影响：明确边界分支缺少直接自动回归，P1 客户端兼容风险未按发布政策处理。
- 修复目标：补齐 XLS/XLSX selector 边界，并提供客户端证据或具名 waiver。
- 明确修复要求：增加 XLS/XLSX 越界 `ByIndex` 和 `SheetVisibility.VeryHidden` 用例，断言 `InvalidHeader`、稳定 `SheetName` 且 plan 不执行；运行可用客户端打开/保存/重开，或取得包含批准人、范围、日期和风险的 waiver。
- 修复后的验证方式：新增测试在 net6/net8 通过；互操作报告包含客户端版本、generation/fixture/roundtrip hash 和结果，或完整 waiver。

## Round 8 最终 Checklist

- [x] 已优先逐项复验上一轮 FIX。
- [x] `FIX-008` Unit 项目文件和运行门禁已恢复。
- [x] `FIX-007` package-only consumer、assets 和 nupkg 内容已独立复验。
- [x] selector 单次解析主链未发现回归，现有 StreamPipeline/Integration 通过。
- [x] API 工具和 net6/net8 runtime contract 已实际运行。
- [x] `git diff --check` 和相关源码诊断通过。
- [ ] lockfile、API 工具和发布证据进入 Git并完成实际待交付状态复现。
- [ ] net6/net8 API contract、Unit 和四 TFM compare 全绿并有批准依据。
- [ ] 性能获得具名预算/批准或 waiver，原始结果可交付。
- [ ] 越界 `ByIndex`、VeryHidden 和客户端政策闭环。
- [x] 未执行 commit、push、tag、publish、PR 或外部通知。

## Round 8 结论

`NEEDS_FIX`。`FIX-007` 和 `FIX-008` 已解决，Round 7 的 Unit 项目构建 BLOCKER 不再存在；但 `FIX-001` Git 交付 BLOCKER、`FIX-003` API HIGH、`FIX-004` 性能 HIGH 和 `FIX-006` SHOULD_FIX 仍未关闭。下一轮修复应继续处理上述四项，完成后再进行独立 Review。

---

# Review Fix Round 7 独立复审

## Round 7 验收摘要

本次复审以当前 `plan.md`、Round 6 `execution.md`、上一轮 `review.md`、实际源码、Git Diff、任务证据和独立命令为依据，优先逐项验证上一轮 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`、`FIX-007`。Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

最终结论：`NEEDS_FIX`，继续保持 `No-Go`。

- `FIX-001`：`PARTIAL/NOT_RESOLVED`。`.gitignore` 已不再忽略任务 benchmark 证据，但 9 个 lockfile、API 工具、baseline、任务报告、API 快照和原始性能证据仍未被 Git 跟踪，无法完成隔离交付复现。
- `FIX-003`：`PARTIAL`。共享 metadata-only canonicalizer 和四 TFM 自动生成/比较入口已建立；独立工具可构建运行，但 Abstractions/Core 在四个 TFM 均与批准 baseline 不匹配，API compare 退出码为 1。当前 Unit 项目文件落盘为空，使 net6/net8 runtime contract 无法复验。
- `FIX-004`：`NOT_RESOLVED`。实现、JSONL 和报告仍明确为 `budgetStatus=UNAPPROVED`；没有批准预算、批准人、日期或具名 waiver，原始证据也未进入 Git。
- `FIX-006`：`NOT_RESOLVED`。本轮按 `fixScope=must` 跳过；仍没有 XLS/XLSX 越界 `ByIndex`、`SheetVisibility.VeryHidden` 直接测试，也没有客户端结果或具名 waiver。
- `FIX-007`：`RESOLVED`。Docs consumer 无生产 `ProjectReference`；隔离 NuGet 缓存 restore/build/test 通过 `11/11`，三个 Bing 依赖均为精确 `2.0.0` 的 `type=package`。
- 新增 `FIX-008`：`REGRESSION`。`tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` 当前磁盘长度为 0，Git diff 显示文件被清空，导致 locked restore、solution build 和 Unit/API runtime gate 报 `MSB4025 Root element is missing`。

## Round 7 上一轮 FIX 复审矩阵

| FIX | Round 6 执行状态 | Round 7 复审状态 | 结论 |
| --- | --- | --- | --- |
| `FIX-001` | PARTIAL/BLOCKED | PARTIAL/NOT_RESOLVED | ignore 子问题已修复；Git 交付和 clean-clone 复现仍未成立。 |
| `FIX-003` | PARTIAL | PARTIAL | 自动入口和共享规范已补；批准比较仍失败，runtime gate 又被空 csproj 阻断。 |
| `FIX-004` | PARTIAL/BLOCKED | NOT_RESOLVED | 测量保持可信但仍无批准预算/waiver，不能判定性能 PASS。 |
| `FIX-006` | SKIPPED/OUT_OF_SCOPE | NOT_RESOLVED | 明确要求的边界测试和互操作政策仍缺失。 |
| `FIX-007` | COMPLETED | RESOLVED | package-only consumer 已由隔离缓存和 assets 独立复验。 |

## Round 7 主要发现

### BLOCKER-001：Unit 测试项目文件落盘为空

- `Get-Item` 和 `File.ReadAllBytes` 均确认 `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` 长度为 `0`。
- `git diff` 将该文件显示为从 64 行完整 XML 删除到空 blob `e69de29`。
- `dotnet restore Bing.Offices.sln --locked-mode` 立即失败：`MSB4025 Root element is missing`。
- VS Code 编辑器读取工具仍能看到未落盘 XML 缓冲区，但构建和 Git 读取的是空磁盘文件；未保存缓冲区不能作为发布证据。
- 该回归阻断 solution restore/build、net6/net8 Unit、API runtime contract 和 selector 专项测试，违反 RH31-001、RH31-401、RH31-404、RH31-604。

### BLOCKER-002：Git 交付与隔离复现仍未成立

- `git ls-files -- '*packages.lock.json'` 无输出；任务目录跟踪文件也为 0。
- 9 个有效 lockfile、`build/ApiSnapshot`、`build/api-snapshot-baseline.json`、任务报告、快照和 benchmark 原始产物仍全部为 `??`。
- `.gitignore` 当前不再命中任务 benchmark JSONL，说明上一轮 ignore 回归已修复；但文件可见不等于已进入交付内容。
- 当前工作树 locked restore 又被空 Unit csproj 阻断，因此连 dirty-worktree 严格还原也不再成立。

### HIGH-001：API 自动化已建立，但批准门禁仍失败

- `build/ApiSnapshot/ApiSnapshot.csproj` 可独立 Release 构建。
- 工具自动处理 `netcoreapp3.1`、`net6.0`、`net7.0`、`net8.0`，测试项目通过 link 编译同一个 `PublicApiSnapshot.cs`，满足“同一 canonical 规范”的结构要求。
- 独立执行结果：Abstractions 实际 `5A1B668E...`，批准 `7B0BA279...`；Core 实际 `5F68499B...`，批准 `41B6D12...`；四 TFM 均失败。Npoi `A0DBE980...` 匹配。
- 未发现批准当前实际值、变更 baseline 或接受 API 漂移的版本决策；不能通过替换 expected hash 收口。
- 因 Unit csproj 为空，net6/net8 `PublicApiContractTest` 当前无法执行，不能把 Round 6 的 `6/7` 当作本轮当前证据。

### HIGH-002：性能结果仍没有批准预算或 waiver

- `Program.TailLatency` 仍在 header、scenario 和控制台输出中写入 `UNAPPROVED`。
- `08-benchmark-report.md` 和两份 Round 4 JSONL 均明确没有批准 baseline、p99/吞吐预算、容差、批准人或 release waiver。
- 技术测量方法和原始结果仍可作为观测证据，但不能转换为发布 PASS。
- 原始 BDN/JSONL 证据仍未被 Git 跟踪，与 `FIX-001` 共同阻断可交付性能签名。

### MEDIUM-001：selector 边界和客户端政策仍未闭环

- 生产主链仍只在 `NpoiExcelImporter.ResolveSheet` 解析一次，并将 `NpoiResolvedSheet` 复用于 plan、materialization 和 failure mapping；未发现第二套主流程回归。
- 当前测试仅覆盖普通 `SheetVisibility.Hidden` 和合法索引 `0/1`；没有 `SheetVisibility.VeryHidden` 或明确越界 `ByIndex` 用例。
- 没有 Excel/WPS/LibreOffice 本轮客户端版本、fixture hash、保存重开结果，也没有具名批准 waiver。
- 当前 Unit csproj 为空，已有 selector Unit 也无法作为当前可执行门禁。

## Round 7 实际验证

| 验证 | 结果 |
| --- | --- |
| Git status / ls-files | FAIL：9 个 lockfile、API 工具、baseline 和整个任务证据目录未跟踪 |
| `git diff --check` | PASS，仅有 CRLF/LF 提示 |
| locked restore | FAIL：Unit csproj 为 0 字节，`MSB4025 Root element is missing` |
| Release solution build | BLOCKED：restore 前置失败，Unit 项目不可解析 |
| Npoi 多 TFM Release build | PASS；netcoreapp3.1/net6/net7/net8 均构建，保留旧 TFM 兼容警告 |
| Integration net6/net8 | PASS，各 `15/15` |
| API snapshot tool build | PASS |
| 四 TFM API compare | FAIL：Abstractions/Core baseline mismatch；工具退出码 1 |
| Unit/API contract net6/net8 | BLOCKED：Unit csproj 落盘为空 |
| Docs package-only consumer | PASS：隔离缓存 restore/build/test `11/11` |
| Docs assets | PASS：三个 Bing 包均为 `type=package`、版本 `2.0.0` |
| selector 越界/VeryHidden | FAIL/NOT_VERIFIABLE：无直接测试，且 Unit 项目不可执行 |
| 性能批准 | FAIL/UNAPPROVED |
| Office/WPS/LibreOffice | NOT_VERIFIABLE，无批准 waiver |

## Round 7 计划验收矩阵

| 计划项 | 状态 | 说明 |
| --- | --- | --- |
| RH31-000 / RH31-001 / RH31-602 | FAIL | Git 交付未成立；locked restore 因空 Unit csproj 失败。 |
| RH31-101 | PARTIAL | 单次解析主链存在；越界/VeryHidden 测试缺失且 Unit 门禁不可执行。 |
| RH31-102 至 RH31-105 | PARTIAL | Npoi 构建与 Integration 通过；Unit 项目损坏，不能完成当前全量复验。 |
| RH31-201 / RH31-202 | PARTIAL | 分类和成员治理代码存在；批准 API 门禁失败。 |
| RH31-203 | PARTIAL/FAIL | 自动四 TFM 工具已补，但批准比较不通过且未进入 Git。 |
| RH31-401 | FAIL | Unit 项目文件为空，P0 Unit 无法运行。 |
| RH31-402 | PASS | net6/net8 Integration 各 `15/15`。 |
| RH31-403 | PASS | package-only consumer 隔离测试 `11/11`，assets 精确指向本轮三个包。 |
| RH31-404 | FAIL | 生产多 TFM build 通过；强制 net6/net8 Unit/runtime gate 不可执行。 |
| RH31-501 / RH31-502 | PARTIAL | 测量存在；批准预算/waiver 和 Git 交付缺失。 |
| RH31-601 | PARTIAL | 报告已如实同步 API/package 状态；发布门禁仍未闭环。 |
| RH31-603 | NOT_VERIFIABLE | 无客户端结果或具名 waiver。 |
| RH31-604 | FAIL | 存在两个 BLOCKER、两个 HIGH 和一个 SHOULD_FIX，维持 No-Go。 |

## Round 7 Git 变更分析

- 分支 `master`，HEAD `1968b24a3ab07b44c3b386a3f761fcdff2fc4315`。
- 当前 17 个已跟踪文件有修改，任务目录、API 工具、baseline 和 lockfile 均未跟踪。
- 生产行为改动集中于 selector 单次解析和计划复用，属于计划范围；API/benchmark/docs/package consumer 也属于发布硬化范围。
- `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj` 被清空不是计划允许的行为，且是本轮直接构建回归。
- Reviewer 未执行 `git add`、commit、push、reset、restore、clean、publish 或 PR。

## Round 7 功能与架构 Review

- `NpoiExcelImporter` 先解析全部 selector，拒绝同一物理 Sheet 冲突，再由 `NpoiImportPlanBuilder.Create(existingSheets)` 建 plan；后续导入和 Failure Workbook 使用解析后的 index/name，符合单次解析目标。
- 生产依赖方向仍为 `Abstractions <- Core <- Npoi`，未发现新增反向依赖或公开 NPOI 类型泄漏。
- API snapshot 使用 MetadataLoadContext，不执行目标程序集代码；固定依赖探针避免全 NuGet 缓存递归冲突。测试和工具链接同一源码，未形成第二套 canonicalizer。
- `BuildScript.csproj` 已排除 `ApiSnapshot/**/*.cs`，独立构建验证该项目边界修复有效。
- 但 API baseline 的来源和批准产物仍不可复现，且所有新增工具/证据未进入 Git，架构改进尚未成为可交付门禁。

## Round 7 测试与文档 Review

- Docs consumer 实际覆盖 DI、provider-neutral API、mapping、CSV、metadata XLS/XLSX 重开和文档 fence；本轮 package-only 证据有效。
- Integration net6/net8 当前均 `15/15`，说明生产 Npoi 主链未出现可见集成回归。
- Unit、selector 和 API runtime contract 因空 csproj 完全不可运行，这是发布阻断，不可用旧 bin 或历史报告替代。
- `04-api-breaking-changes.md`、四份 API diff、`10-final-review.md` 和 `11-final-summary.md` 已如实记录 API mismatch/No-Go；性能报告也未误报 PASS。
- 报告和原始证据仍未被 Git 跟踪，文档真实性不能替代交付完整性。

## Round 7 Findings 与结构化修复任务

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL/NOT_RESOLVED`
- 对应计划项：RH31-000、RH31-001、RH31-403、RH31-602。
- 涉及文件：9 个有效 `packages.lock.json`、`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、当前任务报告/快照/benchmark 证据、`.gitignore`。
- 问题：ignore 规则已修复，但锁图、API 工具和全部发布证据仍不属于 Git 交付内容，且当前 strict restore 已因 `FIX-008` 失败。
- 证据：`git ls-files` 对 lockfile 和任务目录无输出；Git status 将上述内容全部显示为 `??`。
- 影响：待交付提交无法复现依赖图、API 门禁、性能证据或 Review 结论。
- 修复目标：使全部要求文件可由 Git 交付，并在实际待交付状态完成隔离复现。
- 明确修复要求：在修复 `FIX-008` 后，由获授权交付步骤纳管 9 个 lockfile、API 工具/baseline、任务报告、四 TFM 快照和批准保留的原始 benchmark 证据；在隔离 worktree 或实际待交付提交上执行 locked restore、build 和必要门禁。不得仅验证 dirty 工作树。
- 修复后的验证方式：`git ls-files` 列出所有要求文件；隔离工作树 `dotnet restore Bing.Offices.sln --locked-mode`、Release build 和关键门禁退出码均为 0，报告/原始证据可读取。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL`
- 对应计划项：RH31-201、RH31-202、RH31-203、RH31-401、RH31-404。
- 涉及文件：`build/ApiSnapshot/**`、`build/api-snapshot-baseline.json`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、四份 `api-diff-*.md`。
- 问题：自动生成/比较入口和共享规范已建立，但批准 baseline 与当前 Abstractions/Core 产物不一致，四 TFM compare 失败；runtime contract 当前又被 `FIX-008` 阻断。
- 证据：Abstractions actual `5A1B668E...` vs approved `7B0BA279...`；Core actual `5F68499B...` vs approved `41B6D12...`；Npoi 匹配；工具退出码 1。
- 影响：无法确认当前 API 是预期兼容面还是未批准漂移，强制 Unit/API 门禁不通过。
- 修复目标：由可追溯、获批准的产物和 canonical 规范得到全绿 API 门禁，而不是机械替换 expected hash。
- 明确修复要求：先修复 `FIX-008` 并恢复 net6/net8 runtime contract；定位批准 hash 对应的历史产物/规范，生成 actual-vs-approved member diff，交由版本负责人决定恢复产物还是批准新 baseline；若批准变更，记录批准人、日期、版本和 migration，再同步 baseline。将工具、baseline 和快照纳入 Git/CI 验收。
- 修复后的验证方式：net6/net8 `PublicApiContractTest` 全绿；自动命令对四 shipped TFM 比较通过；任意治理签名变化会失败；隔离工作树可重复执行。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-501、RH31-502、RH31-604。
- 涉及文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`08-benchmark-report.md`、JSONL/BDN 原始产物和批准记录。
- 问题：可靠测量仍没有批准 workload、预算、容差、批准人或具名 waiver，原始证据也未进入 Git。
- 证据：源码、报告和 JSONL 均为 `UNAPPROVED`；未发现新的批准记录。
- 影响：不能判定性能回归 PASS，不满足发布 Go 条件。
- 修复目标：将现有测量绑定到具名批准的环境/workload/预算，或取得完整 release waiver，并交付原始证据。
- 明确修复要求：由版本负责人记录冷/热 workload、p99/吞吐预算、允许波动、环境、批准人和日期，或签署带范围和风险的具名 waiver；报告逐项标记 PASS/FAIL/WAIVED；原始 JSONL/BDN 结果进入 Git。无需重写已满足技术口径的测量代码。
- 修复后的验证方式：相同环境重复运行可按批准条件判定；报告和原始数据包含批准身份、范围、日期和结论，Git 可取得证据。

### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-101、RH31-403、RH31-603。
- 涉及文件：`tests/Bing.Offices.Tests/StreamPipelineTest.cs`、互操作报告与 fixture 证据。
- 问题：越界索引、VeryHidden 和客户端政策仍未闭环；本轮按 must scope 跳过不构成发布豁免。
- 证据：测试搜索只有普通 Hidden 和合法 `ByIndex(0/1)`；无客户端结果或批准 waiver。
- 影响：明确边界分支没有自动回归，P1 客户端兼容风险未按发布政策处理。
- 修复目标：补齐 XLS/XLSX selector 边界，并提供客户端证据或具名 waiver。
- 明确修复要求：修复 `FIX-008` 后增加 XLS/XLSX 越界 `ByIndex` 和 `SheetVisibility.VeryHidden` 用例，断言 `InvalidHeader`、稳定 `SheetName` 且 plan 不执行；运行可用客户端打开/保存/重开，或取得包含批准人、范围、日期和风险的 waiver。
- 修复后的验证方式：新增测试在 net6/net8 通过；互操作报告包含客户端版本、generation/fixture/roundtrip hash 和结果，或完整 waiver。

### FIX-008

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`REGRESSION`
- 对应计划项：RH31-001、RH31-401、RH31-404、RH31-604。
- 涉及文件：`tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj`、对应 `packages.lock.json`。
- 问题：Unit 测试项目文件当前磁盘内容为空，导致解决方案无法 restore/build，所有 Unit/API/selector runtime 门禁不可执行。
- 证据：文件长度和 byte count 均为 0；Git diff 指向空 blob；MSBuild 报 `MSB4025 Root element is missing`。编辑器中的未落盘缓冲区不能替代磁盘项目文件。
- 影响：发布最基本的 build 和强制 net6/net8 Unit 门禁失效，并使 Round 6 的 API/selector 测试结论不再是当前证据。
- 修复目标：恢复合法、可解析且包含 Round 6 必需 API snapshot 引用的测试项目配置，不丢失用户当前未落盘内容。
- 明确修复要求：先比较 VS Code 未保存缓冲区、Git 基线和 Round 6 预期配置；将正确 XML 以 UTF-8 落盘，保留 `System.Reflection.MetadataLoadContext` 引用和 `PublicApiSnapshot.cs` link；同步/验证 lockfile，不得仅恢复旧基线而丢掉 Round 6 必需项。
- 修复后的验证方式：项目文件长度非 0 且 MSBuild 可解析；locked restore、Release solution build、net6/net8 全量 Unit、API contract 和 selector 定向测试均可执行，并按各自合同通过或给出真实失败。

## Round 7 最终 Checklist

- [x] 已优先逐项复验上一轮 FIX。
- [x] 已检查 plan、execution、旧 review、源码、Git Diff、报告和原始证据。
- [x] `FIX-007` package-only consumer 已独立复验为 resolved。
- [x] Npoi 多 TFM build 和 net6/net8 Integration 已实际运行通过。
- [ ] Unit 项目文件恢复且 solution locked restore/build 通过。
- [ ] lockfile、API 工具和发布证据进入 Git并完成隔离复现。
- [ ] net6/net8 API contract 与四 TFM compare 全绿并有批准依据。
- [ ] 性能获得具名预算/批准或 waiver，原始结果可交付。
- [ ] 越界 `ByIndex`、VeryHidden 和客户端政策闭环。
- [x] 未执行 commit、push、tag、publish 或 PR。

## Round 7 结论

`NEEDS_FIX`。当前新增一个直接构建 `BLOCKER`，既有 Git 交付 `BLOCKER` 仍未关闭；API 和性能两个 HIGH 仍未通过，selector/互操作 SHOULD_FIX 未处理。`FIX-007` 已解决，但不足以满足 Go 条件。Reviewer 不进行修复，下一步交由 `review-fixer` 处理结构化 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`、`FIX-008`。

---

# 历史记录：Review Fix Round 5 独立复审

## 验收摘要

本次复审以当前 `plan.md`、Round 1-4 `execution.md`、上一轮 Review、实际源码、Git Diff、任务证据和独立命令为依据，优先逐项验证 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`。当前 runtime 仍显示 Round 5 Review Fix 为 `IN_PROGRESS`；Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

最终结论：`NEEDS_FIX`，继续保持 `No-Go`。

- `FIX-001`：`NOT_RESOLVED`。当前 dirty 工作树 locked restore 通过，但 9 个 lockfile 和整个任务证据目录仍未被 Git 跟踪；新增 `.gitignore` 规则还会忽略任务目录下的 benchmark 原始证据。
- `FIX-003`：`REGRESSED`。net6 API 契约仍为 6/7，net8 从上一轮 7/7 回归为 6/7；四 TFM Markdown 仍没有自动生成/比较入口。
- `FIX-004`：`NOT_RESOLVED`。源码和两份 JSONL 仍将 `budgetStatus` 固定为 `UNAPPROVED`，没有批准预算、批准人或具名 waiver，证据也未进入 Git。
- `FIX-006`：`PARTIAL`。已有 selector 场景在 net6/net8 各 9/9 通过，但越界 `ByIndex`、`VeryHidden` 和客户端互操作/waiver 仍不存在。
- 新增 `FIX-007`：Docs 测试从三个 Bing nupkg `PackageReference` 改为生产项目 `ProjectReference`。测试 11/11 通过，但已不再证明 package-only consumer，直接回归 RH31-403。

## 上一轮 FIX 复审矩阵

| FIX | Round 4 状态 | Round 5 复审状态 | 结论 |
| --- | --- | --- | --- |
| `FIX-001` | NOT_RESOLVED | NOT_RESOLVED | 锁图只在 dirty 工作树可用，Git 交付与隔离复现仍不成立。 |
| `FIX-003` | PARTIAL | REGRESSED | net6 继续失败，net8 新增失败；静态快照仍不可自动重建。 |
| `FIX-004` | PARTIAL | NOT_RESOLVED | 测量方法保留，但批准预算/waiver 和 Git 交付均无变化。 |
| `FIX-006` | NOT_RESOLVED | PARTIAL | 既有 selector 用例通过，明确要求的边界与互操作政策仍缺。 |

## 主要发现

### BLOCKER-001：锁文件和发布证据仍不属于 Git 交付面

- `git ls-files` 对 9 个当前 lockfile 和任务目录均无输出。
- `git status --short --untracked-files=all` 将 9 个 lockfile、`00-11` 报告、四份 API 快照、`plan.md`、`execution.md` 和 `review.md` 全部列为未跟踪。
- 当前 `.gitignore` 新增 `/ai_docs/tasks/**/benchmarks/*`，会忽略任务目录下的 benchmark JSONL/原始结果，与 Reviewer 要求“原始证据可交付”相反。
- `dotnet restore Bing.Offices.sln --locked-mode -v:q` 在当前工作树通过，但不能证明待交付提交或干净克隆可复现。

因此 `FIX-001` 保持 `NOT_RESOLVED`。Reviewer 禁止 staging 不构成发布豁免，必须由获授权交付步骤完成 Git 纳管和隔离复验。

### HIGH-001：API 契约从单 TFM 失败扩大为 net6/net8 双目标失败

- net6 `PublicApiContractTest`：6/7，`Bing.Offices.Abstractions` expected `7B0BA279...`，actual `3D3AA5DD...`。
- net8 `PublicApiContractTest`：6/7，expected `7B0BA279...`，actual `8DB5790D...`。
- net6/net8 全量 Unit 均为 371/372，唯一失败均为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。
- 当前哈希实现已从 public-only 扩展为 public/protected/protected-internal 成员，但批准哈希没有通过稳定生成流程同步；不同 runtime 仍产生不同结果。
- 仓库搜索只发现测试私有哈希函数和四份静态 Markdown，没有可执行的四 TFM snapshot 生成/比较脚本、项目 target 或 CI 入口。

这不是可接受的“环境差异”：net8 在同一当前环境也已失败。`FIX-003` 标记为 `REGRESSED`，RH31-203 和强制 Unit 门禁均未通过。

### HIGH-002：性能结果仍无批准预算或 waiver

- `Program.TailLatency` 的 header、repetition、scenario 和控制台输出仍硬编码 `UNAPPROVED`。
- 两份 Round 4 JSONL 仍全部为 `budgetStatus=UNAPPROVED`；报告明确没有历史 baseline、环境预算、批准人或 release waiver。
- 未发现 Round 5 新批准记录或重新判定产物。
- benchmark 原始证据目录仍未被 Git 跟踪，且当前 `.gitignore` 会忽略任务 benchmark 目录。

测量技术口径仍然成立，但不能据此判定 PASS。`FIX-004` 保持 `NOT_RESOLVED`。

### MEDIUM-001：selector 边界矩阵和客户端政策仅部分处理

- 名称大小写、正常索引、混合 selector、同物理 Sheet 冲突和 failure mapping 等现有测试在 net6/net8 各 9/9 通过。
- 仓库搜索仍没有 `SheetVisibility.VeryHidden` 直接测试。
- 没有明确越界 `ByIndex` 测试；现有索引场景只使用 0/1 合法索引。
- 没有 Excel/WPS/LibreOffice 本轮客户端版本、fixture hash、保存重开结果，也没有具名批准 waiver。

因此 `FIX-006` 从 `NOT_RESOLVED` 改为 `PARTIAL`，但仍是未关闭的 SHOULD_FIX。

### HIGH-003：Docs consumer 回退为项目引用，package-only 门禁失效

- `Bing.Offices.Docs.Tests.csproj` 删除三个 Bing `PackageReference`，新增 `ProjectReference` 指向 `src/Bing.Offices.Npoi`。
- Docs 测试当前 11/11 通过，但实际构建日志直接构建 `Bing.Offices.Abstractions/Core/Npoi` 项目，不能证明本轮 nupkg 的依赖、资产、XML 文档或 restore 条件。
- 该变化与计划 RH31-403“验证 nupkg 而不是项目引用”直接冲突，也使此前 package consumer 报告失真。

这是本轮修复直接引入的发布门禁回归，新增 `FIX-007`。

## 实际验证

| 验证 | 结果 |
| --- | --- |
| `git status --short --untracked-files=all` | FAIL：9 个 lockfile 和全部任务证据仍未跟踪 |
| `git ls-files` lockfile/任务目录 | FAIL：无输出 |
| `git diff --check` | PASS，仅有 CRLF/LF 提示 |
| locked restore | PASS，仅限当前 dirty 工作树 |
| Release solution build | PASS，0 errors；保留 legacy/旧 TFM 警告 |
| API 契约 net6 | FAIL，6/7；expected `7B0BA279...`，actual `3D3AA5DD...` |
| API 契约 net8 | FAIL，6/7；expected `7B0BA279...`，actual `8DB5790D...` |
| Unit net6 | FAIL，371/372；API snapshot |
| Unit net8 | FAIL，371/372；API snapshot |
| selector 定向 net6/net8 | PASS，各 9/9；不含越界和 VeryHidden |
| Integration net6/net8 | PASS，各 15/15 |
| Docs tests net8 | PASS，11/11；但使用 ProjectReference，不是 package-only consumer |
| changed-file diagnostics | PASS，无编译诊断错误 |
| 性能批准 | FAIL/UNAPPROVED |
| Office/WPS/LibreOffice | NOT_VERIFIABLE，无批准 waiver |

## 计划验收矩阵

| 计划项 | 状态 | 说明 |
| --- | --- | --- |
| RH31-000 / RH31-001 / RH31-602 | FAIL | 锁文件与发布证据未进入 Git，干净交付不可复现。 |
| RH31-101 | PARTIAL | 单次解析主链和已有 selector 场景有效；越界与 VeryHidden 缺测试。 |
| RH31-102 至 RH31-105 | PASS/PARTIAL | 本轮未发现既有 IO/资源合同回归；metadata 测试通过但 package consumer 证据回退。 |
| RH31-201 / RH31-202 | PARTIAL | 分类和成员治理代码存在；API runtime gate 失败。 |
| RH31-203 | FAIL | net6/net8 snapshot 均失败，四 TFM 静态比较不可自动执行。 |
| RH31-401 / RH31-404 | FAIL | net6/net8 强制 Unit 均非全绿。 |
| RH31-402 | PASS | net6/net8 Integration 各 15/15。 |
| RH31-403 | REGRESSED/FAIL | Docs 项目改为生产 ProjectReference，不再验证 nupkg。 |
| RH31-501 / RH31-502 | PARTIAL | 测量方法可复现，但预算、waiver 和证据交付未闭环。 |
| RH31-601 | PARTIAL | README/迁移文档有修正，package consumer 结构与报告不一致。 |
| RH31-603 | NOT_VERIFIABLE | 无客户端结果或具名 waiver。 |
| RH31-604 | FAIL | 存在 BLOCKER/HIGH/SHOULD_FIX，维持 No-Go。 |

## Findings 与结构化修复任务

### FIX-001

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-000、RH31-001、RH31-403、RH31-602。
- 涉及文件：9 个项目 `packages.lock.json`、当前任务报告/证据目录、`.gitignore`。
- 问题：锁文件和发布证据仍不属于 Git 交付内容，且 benchmark 证据被新增 ignore 规则排除。
- 证据：`git ls-files` 无 lockfile/任务证据；`git status` 全部显示 `??`；locked restore 只在 dirty 工作树通过。
- 影响：待交付提交和干净克隆无法复现依赖图、API 快照、benchmark 或 Review 结论。
- 修复目标：使全部有效锁图和批准保留的发布证据可由 Git 交付，并完成隔离复现。
- 明确修复要求：由获授权交付步骤纳入 9 个 lockfile、任务报告、四 TFM 快照及需要保留的原始 benchmark 证据；调整 `.gitignore`，不得忽略要求交付的原始证据；在隔离 worktree 或实际待交付提交上执行 locked restore 和必要门禁。
- 修复后的验证方式：`git ls-files` 列出所有要求文件；隔离工作树 `dotnet restore Bing.Offices.sln --locked-mode` 退出码 0，且所需报告/原始证据可读取。

### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`REGRESSED`
- 对应计划项：RH31-201、RH31-202、RH31-203、RH31-401、RH31-403、RH31-404。
- 涉及文件：`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、四份 `api-diff-*.md`、API snapshot 生成/比较入口。
- 问题：成员 snapshot 在 net6/net8 均失败，且静态快照仍不可自动重建。
- 证据：net6 actual `3D3AA5DD...`，net8 actual `8DB5790D...`，批准值均为 `7B0BA279...`；两目标 Unit 均 371/372；无生成/比较脚本或 CI target。
- 影响：强制 Unit 门禁失败，API 漂移无法可靠区分预期变化、宿主差异和算法变化。
- 修复目标：建立跨 runtime 确定、自动生成、可比较且可交付的 canonical API 门禁。
- 明确修复要求：定位 net6/net8 差异及 net8 回归根因；不要仅替换 expected hash；将 canonical line 生成与批准 snapshot 分离为可执行工具/target，对 netcoreapp3.1/net6/net7/net8 目标程序集自动生成并比较；测试与静态工具必须使用同一规范；同步报告真实状态。
- 修复后的验证方式：net6/net8 `PublicApiContractTest` 全绿；自动命令对四 shipped TFM 比较通过，修改任一受治理签名会失败；隔离工作树可重复执行。

### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`NOT_RESOLVED`
- 对应计划项：RH31-501、RH31-502、RH31-604。
- 涉及文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`08-benchmark-report.md`、JSONL/BDN 原始产物和批准记录。
- 问题：可靠测量仍没有批准预算、批准人或具名 waiver，原始证据也未交付。
- 证据：源码、报告和 JSONL 均为 `UNAPPROVED`；未发现 Round 5 批准记录。
- 影响：不能将性能结果判定为 PASS，不满足发布 Go 条件。
- 修复目标：将测量绑定到批准 workload/预算/容差，或取得完整 release waiver，并交付原始证据。
- 明确修复要求：由版本负责人记录环境、冷/热 workload、p99/吞吐预算、允许波动、批准人和日期，或签署具名 waiver；报告逐项标记 PASS/FAIL/WAIVED；原始 JSONL/BDN 结果进入 Git 交付。无需再次重写已满足口径的测量代码。
- 修复后的验证方式：同环境重复运行按批准条件自动或人工判定；报告和原始数据包含批准身份、范围、日期和结论，Git 可取得证据。

### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 复审状态：`PARTIAL`
- 对应计划项：RH31-101、RH31-403、RH31-603。
- 涉及文件：`tests/Bing.Offices.Tests/StreamPipelineTest.cs`、互操作报告与 fixture 证据。
- 问题：已有 selector 场景通过，但越界索引、VeryHidden 和客户端政策仍未闭环。
- 证据：定向测试 net6/net8 各 9/9；搜索无 `SheetVisibility.VeryHidden` 和越界 `ByIndex`；无客户端结果或批准 waiver。
- 影响：明确边界分支缺少回归保护，P1 客户端兼容风险未按发布政策处理。
- 修复目标：补齐 XLS/XLSX selector 边界，并提供客户端证据或具名 waiver。
- 明确修复要求：增加 XLS/XLSX 越界 `ByIndex` 和 `SheetVisibility.VeryHidden` 用例，断言 `InvalidHeader`、稳定 `SheetName` 且 plan 不执行；运行可用客户端打开/保存/重开，或取得包含批准人、范围、日期和风险的 waiver。
- 修复后的验证方式：新增测试在 net6/net8 通过；互操作报告包含客户端版本、generation/fixture/roundtrip hash 和结果，或完整 waiver。

### FIX-007

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 复审状态：`REGRESSION`
- 对应计划项：RH31-403、RH31-602、RH31-604。
- 涉及文件：`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`、`DocsConsumerTest.cs`、`07-package-consumer-report.md`。
- 问题：Docs consumer 已从本地 Bing nupkg 引用回退为生产项目引用。
- 证据：csproj 删除三个 Bing `PackageReference`，新增 `ProjectReference`；测试日志构建生产项目；11/11 PASS 不能证明包消费。
- 影响：无法验证 nupkg 依赖、资产选择、XML 文档、restore source/cache 前提和外部消费者编译，发布包门禁失效。
- 修复目标：恢复真正 package-only 的独立消费者验证，同时保留 metadata XLS/XLSX 重开测试。
- 明确修复要求：Docs consumer 不得引用生产项目；使用本轮本地 nupkg 的三个 Bing `PackageReference`，明确本地 Bing source 和第三方 source/预热缓存；断言 assets 精确解析本轮版本；运行文档 fence、DI、mapping、CSV、metadata XLS/XLSX 重开测试；同步 package consumer 报告。
- 修复后的验证方式：csproj 无生产 `ProjectReference`；隔离 NuGet 缓存按报告命令 restore/test 通过；`project.assets.json` 指向本轮三个 Bing 包；nupkg 内容和依赖元数据复核通过。

## 最终 Checklist

- [x] 已优先逐项复验上一轮 FIX。
- [x] Release build、selector 定向、Integration 和 Docs 当前测试已实际运行。
- [x] 已确认 selector 单次解析主链未发现现有场景回归。
- [ ] 锁文件和发布证据进入 Git 并完成隔离复现。
- [ ] net6/net8 API 契约与 Unit 全绿，四 TFM snapshot 自动化。
- [ ] 性能获得具名预算/批准或 waiver，原始结果可交付。
- [ ] 越界 `ByIndex`、VeryHidden 和客户端政策闭环。
- [ ] 恢复 package-only Docs consumer。
- [x] 未执行 commit、push、tag、publish 或 PR。

## 结论

`NEEDS_FIX`。当前存在一个 BLOCKER、三个 HIGH（其中 API 和 package consumer 为本轮回归）及一个未关闭 MEDIUM。发布结论继续为 `No-Go`，Reviewer 不进行修复。

---

# 历史记录：Review Fix Round 4 复审

## 验收摘要

本次独立复审优先验证上一轮 `FIX-001`、`FIX-003`、`FIX-004`、`FIX-006`，证据来自 `plan.md`、最新 `execution.md`、上一轮 Review、当前源码与 Git Diff，以及本轮实际 restore/test 检查。未修改业务代码、测试代码、`plan.md` 或 `execution.md`。

最终结论：`NEEDS_FIX`。Round 4 已完成 API 类型分类对齐、自动成员治理账、四份静态 TFM 快照，以及具备预热、重复、端到端排队延迟和环境身份的并发测量；但锁文件和全部任务证据仍未进入 Git 交付面，net6 运行时 API 契约实际失败，四 TFM 静态快照没有自动比较入口，性能预算或具名 waiver 仍缺失，`FIX-006` 仍未处理。计划要求 P0/Blocker 为零且 P1 已修复或获批准 waiver，因此当前不能发布。

## 上一轮 FIX 验证

| FIX | 复审状态 | 证据与结论 |
| --- | --- | --- |
| `FIX-001` | `NOT_RESOLVED` | 当前工作树 `dotnet restore Bing.Offices.sln --locked-mode -v:q` 通过；但 9 个有效 lockfile 全部未被 Git 跟踪，任务目录的报告与原始证据跟踪数为 0，无法从交付提交或干净克隆复现。 |
| `FIX-002` | `RESOLVED` | 本轮未发现已验收的资源、所有权、异常和 IO 行为回归。 |
| `FIX-003` | `PARTIAL` | 类型分类冲突和逐成员治理账已修复，net8 API 契约 7/7 通过；但 net6 API 契约 6/7，`Bing.Offices.Abstractions` 实际哈希为 `3D3AA5...`，不等于批准值 `7B0BA2...`。四份静态快照仅为未跟踪 Markdown，仓库没有自动生成或比较它们的入口。 |
| `FIX-004` | `PARTIAL` | 测量方法已补足 64 次预热、5 轮、每档 1280 样本、队列提交到完成延迟、worker 启动和环境身份；两次 JSONL 也如实保留波动。但所有结果仍为 `budgetStatus=UNAPPROVED`，没有批准预算或具名 waiver，产物也未进入 Git 交付面。 |
| `FIX-005` | `RESOLVED` | 本轮未发现 metadata 隔离与并发证据回归。 |
| `FIX-006` | `NOT_RESOLVED` | 测试搜索仍未发现越界 `ByIndex` 或 `SheetVisibility.VeryHidden` 直接用例；执行报告明确因 must scope 跳过，Office 互操作仍只有 `NOT_VERIFIABLE`，没有具名批准 waiver。 |

## 实际验证

| 验证 | 结果 |
| --- | --- |
| `git status --short --untracked-files=all` | FAIL：9 个 lockfile、任务报告、四份 API 快照和 benchmark 原始证据均为未跟踪文件 |
| `git diff --check` | PASS，仅有 CRLF/LF 提示 |
| `dotnet restore Bing.Offices.sln --locked-mode -v:q` | PASS，仅证明当前 dirty 工作树可还原 |
| API 契约 net8.0，`--no-build --no-restore` | PASS，7/7 |
| API 契约 net6.0，`--no-build --no-restore` | FAIL，6/7；Abstractions expected `7B0BA2...`，actual `3D3AA5...` |
| API 契约普通构建执行 | BLOCKED BY ENVIRONMENT：PowerShell 进程持有 net6/net8 测试输出 DLL；不能替代上述断言结果 |
| 四 TFM 静态 API 快照 | PARTIAL：文件内容存在且哈希一致，但无自动生成/比较入口，且未被 Git 跟踪 |
| tail-latency 两次归档 | MEASURED：方法口径完整；预算状态均为 `UNAPPROVED` |
| selector 越界/VeryHidden 与客户端互操作 | NOT_VERIFIABLE：无直接测试、无客户端证据、无批准 waiver |

## 计划验收矩阵

| 计划项 | 状态 | 说明 |
| --- | --- | --- |
| RH31-000 / RH31-001 / RH31-403 / RH31-602 | FAIL | locked restore 在当前工作树通过，但锁图与任务证据不可由 Git 交付，干净克隆验收未完成。 |
| RH31-101 | PARTIAL | selector 单次解析和已有名称/索引/冲突场景存在；越界索引和 VeryHidden 缺少直接回归。 |
| RH31-102 / RH31-103 / RH31-104 / RH31-105 | PASS | 本轮未发现资源、IO、metadata、异常合同回归。 |
| RH31-201 / RH31-202 / RH31-203 | PARTIAL | 类型与成员治理已明显完善；net6 运行时契约失败，四 TFM 静态比较不可自动执行。 |
| RH31-401 至 RH31-404 | PARTIAL | net8 API 门禁通过；net6 API 门禁失败，net7/netcoreapp3.1 仅有静态报告，交付锁图也未闭环。 |
| RH31-501 / RH31-502 | PARTIAL | 正式 BDN 与并发测量方法成立；批准预算/waiver 和证据交付仍缺。 |
| RH31-601 至 RH31-604 | PARTIAL | 报告如实保持 No-Go；Git 交付、性能批准和客户端互操作政策未闭环。 |

## Git 变更分析

- 分支 `master`，HEAD `1968b24a3ab07b44c3b86a3f761fcdff2fc4315`。
- 14 个已跟踪文件存在修改；9 个 lockfile 和整个当前任务目录仍为未跟踪内容。
- 当前生产行为变更集中于 selector 解析与导入链，测试/API/benchmark/docs 变更均属于计划范围；未发现需要新增 FIX 的无关行为变化。
- Reviewer 的禁止 `git add` 约束不构成发布豁免。只有授权交付步骤将文件纳入提交后，才能验证干净克隆。

## 功能与 API Review

- selector 解析结果已在 plan、materialization 和 failure mapping 中复用，没有发现第二套主流程实现。
- `ApiTypeCategories` 与 `04-api-breaking-changes.md` 的 Round 3 分类冲突已消除；未知或缺失导出类型会失败。
- `PublicApi_ExportedTypes_ShouldHaveGovernedClassification` 已覆盖 public/protected/protected-internal 构造函数、属性、字段和方法，并绑定分类、source/binary 影响和迁移策略，上一轮“只有 `Assert.NotNull`”的问题已解决。
- `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 已为 shipped TFM 使用显式条件分支，没有通用 `#else`。但当前 net6 VSTest 真实执行仍失败，因此报告中“net6 test-host 通过”的结论不成立。
- 四份 `api-diff-*.md` 没有被任何测试、脚本、项目或 CI 配置引用。它们是静态结果记录，不是可持续门禁，且当前不属于 Git 交付内容。

## 架构与维护性 Review

- `Abstractions <- Core <- Npoi` 依赖方向未见新增反向依赖，provider-neutral 公共成员泄漏检查在 net8 通过。
- NPOI 程序集仍只公开 DI 注册入口，未发现新增 NPOI 类型泄漏。
- API canonical line 依赖运行时 reflection 的 `Type.FullName`。net6 与独立静态分析产生不同哈希，说明当前算法或宿主装载上下文尚未达到跨运行时确定性；在根因未闭环前不能将静态 PASS 替代 runtime gate。

## 性能与资源 Review

- Round 4 的并发入口已纠正上一轮方法缺陷：worker ready/start gate、正式运行前预热、5 次重复、每档 1280 样本、提交前时间戳到 plan 完成的端到端延迟、墙钟吞吐、worker 启动和运行环境身份均已记录。
- 两次运行的并发 64 p99 为 `9431-10136 us`，吞吐为 `36141-40983 ops/s`；波动被如实记录，没有误报 PASS。
- 两份原始结果和报告均明确 `UNAPPROVED`。计划 Go 条件要求 P1 已修复或有具名批准 waiver，故性能项仍不能验收通过。
- BDN 与资源报告如实说明 NPOI DOM、高分配和压缩/DOM 峰值限制，未发现新的夸大表述。

## 测试与文档 Review

- net8 API 契约 7/7 通过；net6 API 契约稳定复现为 6/7，失败点是 Abstractions 成员快照哈希。
- 当前进程持有测试 DLL 导致普通 build/test 不能覆盖输出，这是环境阻塞；已有输出的 `--no-build` 断言失败仍是有效运行时证据。
- 文档已更新为自动成员治理账、四份静态快照和新 benchmark 方法；但 `04-api-breaking-changes.md` 声称 net6 test-host 通过，与本轮实际测试冲突。
- `10-final-review.md`、`11-final-summary.md` 和 `execution.md` 对剩余 No-Go 风险总体表述诚实。

## Findings 与修复任务

### 历史 FIX-001（Round 4）

- 严重程度：`BLOCKER`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH31-000、RH31-001、RH31-403、RH31-602。
- 涉及文件：9 个项目 `packages.lock.json`、当前任务报告与证据目录、`.gitignore`。
- 问题：严格还原只在当前未提交工作树成立，锁文件和发布证据不属于 Git 交付内容。
- 证据：9 个有效 lockfile 的 `git ls-files --error-unmatch` 均为 false；当前任务目录被跟踪文件数为 0。
- 影响：交付提交和干净克隆无法获得锁图、API 快照、benchmark 原始证据或验收报告，发布可复现性不成立。
- 修复目标：使全部有效锁图和必须保留的发布证据可由 Git 提交交付，并完成隔离复现。
- 明确修复要求：由获授权的交付步骤纳入 9 个 lockfile 及批准保留的任务证据；基于实际待交付提交或隔离 worktree 执行严格还原和必要测试。不得只验证当前 dirty 工作树。
- 修复后的验证方式：`git ls-files` 列出全部有效 lockfile 和要求的报告/原始证据；隔离工作树执行 `dotnet restore Bing.Offices.sln --locked-mode` 退出码 0。

### 历史 FIX-003（Round 4）

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH31-201、RH31-202、RH31-203、RH31-403。
- 涉及文件：`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、四份 `api-diff-*.md`、API 报告及 CI/验证入口。
- 问题：类型/成员治理账已完成，但 net6 运行时 API 契约失败，四 TFM 静态快照没有可复现的自动生成和比较门禁。
- 证据：net8 7/7；net6 6/7，Abstractions expected `7B0BA279...`、actual `3D3AA5DD...`。仓库搜索仅发现报告链接和测试私有哈希函数，没有静态快照生成/比较入口。
- 影响：shipped TFM 的 API 门禁不一致；手工 Markdown `PASS` 可能在 API 漂移后继续保持绿色，不能支撑发布或版本审批。
- 修复目标：使 canonical API 比较在 net6/net8 运行时及 netcoreapp3.1/net6/net7/net8 静态产物上确定、自动且可交付。
- 明确修复要求：定位并消除 net6 reflection 哈希不确定性的根因，不得仅改期望哈希适配污染宿主；提供可执行的四 TFM 静态快照生成/比较命令并接入验收或 CI；同步修正文档中的 net6 状态；快照和生成逻辑纳入 Git 交付。
- 修复后的验证方式：net6/net8 `PublicApiContractTest` 全部通过；自动命令对四个 shipped NPOI TFM 生成并比较 canonical surface，任意新增/删除/改签名均失败；隔离工作树复验通过。

### 历史 FIX-004（Round 4）

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH31-501、RH31-502、RH31-604。
- 涉及文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`08-benchmark-report.md`、benchmark 原始产物和发布批准记录。
- 问题：测量方法已满足上一轮技术要求，但没有批准的环境限定预算、历史 baseline 或具名 release waiver，且原始证据未进入 Git 交付面。
- 证据：两次 JSONL 均记录 `budgetStatus=UNAPPROVED`；报告明确高并发波动和无批准人；相关文件均未跟踪。
- 影响：无法将测量结果判定为性能回归 PASS，也不满足计划中的发布 Go 条件。
- 修复目标：将现有可靠测量绑定到具名批准的 workload、预算/容差，或取得带范围和风险的 release waiver，并使证据可交付。
- 明确修复要求：由版本负责人明确环境、冷/热 workload、p99/吞吐预算和允许波动，或签署具名 waiver；报告逐项给出批准人和 PASS/FAIL/WAIVED；原始 JSONL 与报告纳入交付。无需再次重写已满足口径的测量实现，除非批准要求改变。
- 修复后的验证方式：同环境重复运行按批准容差判定；报告包含批准人、日期、范围、预算和结论，或完整 waiver；Git 可取得对应原始证据。

### 历史 FIX-006（Round 4）

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH31-101、RH31-403、RH31-603。
- 涉及文件：selector 测试、互操作报告与 fixture 证据。
- 问题：索引越界和 VeryHidden 直接测试仍不存在，Office/WPS/LibreOffice 互操作没有执行或具名批准 waiver。
- 证据：测试搜索仅发现正常 `ByIndex`、普通 `Hidden` 和 failure mapping 场景；没有 `VeryHidden` 或明确越界用例。执行报告确认本轮跳过。
- 影响：计划明确要求的边界分支缺少自动回归保护，客户端互操作 P1 风险未按发布政策闭环。
- 修复目标：完成 XLS/XLSX selector 边界矩阵，并按政策提供客户端证据或批准 waiver。
- 明确修复要求：增加越界 `ByIndex` 和 `SheetVisibility.VeryHidden` 用例，断言稳定错误码、SheetName 和 plan 不执行；运行可用客户端的打开/保存/重开流程，或取得包含批准人、范围和风险的 `NOT_VERIFIABLE` waiver。
- 修复后的验证方式：新增测试在 net6/net8 通过；互操作报告包含客户端版本、fixture/产物哈希和结果，或完整具名 waiver。

## 未完成、风险与问题分级

- `BLOCKER`：`FIX-001`，Git 交付与干净克隆复现未成立。
- `HIGH`：`FIX-003`，net6 API runtime gate 失败且四 TFM 静态门禁不可自动执行。
- `HIGH`：`FIX-004`，性能没有批准预算/waiver，不能判定发布 PASS。
- `MEDIUM`：`FIX-006`，selector 边界与客户端互操作政策未闭环。
- `LOW`：现有 legacy API/TFM/analyzer 警告属于已知残余风险，本轮不新增风格型 FIX。

## 最终 Checklist

- [x] 已优先逐项复验上一轮 FIX。
- [x] 已检查计划、执行报告、源码、Git Diff、测试和原始证据。
- [x] 当前工作树 strict locked restore 通过。
- [x] API 类型分类和成员治理账已复核。
- [x] tail-latency 方法与两次原始结果已复核。
- [ ] 锁文件和发布证据进入 Git，并在隔离工作树完成复现。
- [ ] net6 API 契约通过，四 shipped TFM 静态比较自动化并可交付。
- [ ] 性能预算获得具名批准，或具名 waiver 完成。
- [ ] selector 越界/VeryHidden 和客户端互操作政策闭环。

## 结论

`NEEDS_FIX`。Round 4 已实质解决成员治理账和并发测量方法，但仍存在一个 BLOCKER、两个 HIGH 和一个 MEDIUM 的未关闭修复项。Reviewer 不进行修复；后续应按本文件执行 Review Fix。
