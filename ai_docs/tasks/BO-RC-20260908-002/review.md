<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BO-RC-20260908-002
AI_REVIEWED_AT: 2026-09-11T21:54:57.7861767+08:00

# BO-RC-20260908-002 独立代码审查

## 1. 结论

当前状态：`NEEDS_FIX`。

- 未解决 `MUST_FIX`：0。
- 未解决 `SHOULD_FIX`：1，`FIX-003`。
- `FIX-017` 已通过独立复核并关闭：候选源码身份改用 Git 规范化 blob，文本证据按 LF 规范化散列，identity self-test 已覆盖真实 LF/CRLF fresh clone、证据缺失及源码/程序集/nupkg/artifact 篡改。
- 本地 locked restore、Release build、双 TFM Unit/Integration、Docs、API compare 和资源矩阵均通过，未发现新的业务实现回归。

当前不能批准发布。剩余阻塞不是仓库内算法缺陷，而是计划明确要求的外部门禁：目标生产 `2C4G` 入口实际共享 gate 的容量必须为 `1`，并且最终候选提交必须取得 GitHub Actions 全流程绿灯。

本轮只更新本 `review.md`；未修改生产源码、测试、配置、`plan.md` 或 `execution.md`，未执行 commit、push、PR、部署或发布。

## 2. 计划阶段验收矩阵

| 阶段 | 状态 | 结论 |
| --- | --- | --- |
| P0 基线、locked restore、Release build、pack | `PASS（本地）` | locked restore 在可访问 NuGet 签名服务的环境通过；Release solution build 为 0 warning / 0 error。 |
| P1 Async File Integration、Template 边界 | `PASS` | CSV/Excel Integration 双 TFM 各 `38/38`，0 failed / 0 skipped。 |
| P2 PackageReference-only consumer | `PASS（已有本地证据）` | 仓库和执行记录包含 net6/net8 package-only consumer；最终远端运行仍归 P9 外部门禁。 |
| P3 API governance、breaking、snapshot | `PASS（本地）` | 四个扩展删除有成员级批准；identity self-test 和双 TFM API compare 通过；`FIX-017` 已关闭。 |
| P4 Resource Matrix | `PASS（本地受控）` | Round 9 原始 JSONL 复算为 36/36；submitted/completed=`1530/1530`，maxActive=`1`，64 档 maxQueued=`63`，错误/残留=`0/0`。该结果不等价于生产入口配置。 |
| P5/P6 CSV、UniqueTracker、cache/staging | `PASS（本地）` | 职责测试和正式对照证据保持有效，双 TFM 回归通过。 |
| P7 Benchmark 与正式对照 | `PASS（本地测量）` | before/candidate 证据已形成；目标 2C4G 生产环境仍由 `FIX-003` 验收。 |
| P8 文档和工程治理 | `PASS` | README、Async 边界、迁移说明和 Office 规则与当前实现一致。 |
| P9 报告、最终门禁、独立 Review | `FAIL（外部门禁）` | 无最终候选 GitHub green run，且生产 2C4G 单槽位入口尚无部署配置或运行采样。 |

## 3. 本轮独立验证

- `dotnet restore Bing.Offices.sln --locked-mode --verbosity minimal`：沙箱内因 NuGet TLS/凭证失败；允许访问外部签名服务后通过，属于环境访问限制而非锁文件错误。
- `dotnet build Bing.Offices.sln -c Release --no-restore /m:1`：通过，0 warning / 0 error。
- Unit：net6.0、net8.0 各 `729 passed / 0 failed / 0 skipped`。
- Integration：net6.0、net8.0 各 `38 passed / 0 failed / 0 skipped`。
- Docs：net8.0 `10 passed / 0 failed / 0 skipped`。
- Candidate identity self-test：通过真实 `clone --no-local` 的 LF/CRLF clean checkout；源码、程序集、nupkg、artifact 缺失或篡改均按预期失败。
- API snapshot compare：net6.0、net8.0 均通过，零未批准差异。
- Resource Matrix：按 `result` 实际字段复算为 `36/36`，submitted/completed=`1530/1530`，queue entry events=`1458`，maxActive=`1`，maxQueued=`63`，errors/leftovers=`0/0`。
- 生产源码搜索未发现 `Task.Run`、`.Result` 或 `.Wait()` 伪异步调用。
- `git diff --check`：退出码 0；仅提示 baseline JSON 和 ProfileFixtures XML 将由 CRLF 转为 LF。
- 当前四份 baseline 必需证据已不再被 `.gitignore` 排除，但仍显示为 untracked；CI 新增的 tracked-file 门禁会在最终提交遗漏任一文件时失败。
- 当前环境没有 `gh` CLI，也没有仓库内可验证的最终候选 run URL/head SHA/job 结果，因此未能确认远端绿灯。

## 4. 未解决问题

### FIX-003

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 状态：`OPEN / EXTERNAL_GATE`
- 对应要求：P4/P9 的生产 `2C4G`、最大实际并行度 `1`、最终候选证据提交和 GitHub CI 绿灯。

#### 问题

仓库内 ResourceProbe 已证明受控 `SemaphoreSlim(1,1)` 在 1/4/16/64 请求档位下保持 `maxActiveOperations=1`，但代码搜索确认该 gate 位于测试探针；生产部署入口的实现或配置不在当前仓库中。当前也没有最终候选提交对应的 GitHub Actions 全流程绿色运行证据。

此外，API baseline 绑定的四份证据文件目前虽然可跟踪，但仍为 untracked。计划允许执行阶段不自动 `git add`/commit；发布候选提交必须显式包含它们，否则 CI 会在 API compare 前按设计失败。

#### 证据

- `tests/Bing.Offices.ResourceProbe/StagingResourceMatrix.cs` 定义 `MaxActualParallelism = 1` 并创建 `SemaphoreSlim(1,1)`；`src/` 中没有对应生产入口 gate。
- `artifacts/resource-matrix-round9.jsonl`：36/36、submitted/completed=`1530/1530`、maxActive=`1`、64 档 maxQueued=`63`、errors/leftovers=`0/0`。
- `docs/excel/async-io.md` 将生产入口共享 gate 明确列为部署要求，不是已完成实现证据。
- `.github/workflows/ci.yml` 已包含双 TFM、package consumer、资源契约、API identity 和 tracked-file 门禁，但当前没有绑定最终候选 head SHA 的绿色 run 记录。
- `git status --short` 显示四份 API/审批证据为 `??`；`git check-ignore` 确认它们已经可跟踪。

#### 影响

无法证明用户批准的生产资源约束在真实部署入口生效，也无法证明 Linux clean checkout 上的最终提交已完整包含证据并通过所有 CI 门禁。因此当前候选不能获得最终发布批准。

#### 修复目标

形成一个完整、可审计的最终候选：生产 `2C4G` 入口共享单槽位 gate 已由部署配置或运行采样证明；四份 baseline 必需证据包含在候选提交中；该提交对应的 GitHub Actions 全流程绿色。

#### 修复要求

1. 在目标部署仓库或生产配置中提供入口级共享 gate 容量为 `1` 的证据；限流必须发生在创建 Workbook/provider 工作之前。
2. 在目标 `2C4G` 环境运行 1/4/16/64 请求档位，全部完成且 `maxActiveOperations <= 1`，并保留错误、排队和临时文件清理指标。
3. 最终提交必须包含 `api-candidate.json`、`api-candidate-identity.json`、`fix-review-semantic-diff.json` 和 `api-breaking-approval.md` 四份 baseline 必需证据。
4. 推送最终候选并取得 GitHub Actions 全流程 green run；记录 run URL、head SHA、job/step 和上传 artifacts，head SHA 必须与发布候选一致。

#### 验证方式

核对生产入口配置或采样和 `2C4G` 资源矩阵；在最终提交 clean checkout 中执行 tracked-file 门禁、locked restore、Release build、双 TFM Unit/Integration、Docs、pack、package consumer、identity self-test 与 API compare；确认 GitHub run 的 head SHA 与候选一致且所有强制步骤成功。

## 5. 历史问题复核

| Fix | 状态 | 独立复核结论 |
| --- | --- | --- |
| `FIX-014` | `CLOSED` | verifier 继续校验候选源码、程序集、nupkg 和四份 artifact。 |
| `FIX-015` | `CLOSED` | CSV async BOM 与完整字节合同测试在双 TFM 通过。 |
| `FIX-016` | `CLOSED` | 当前候选正式性能证据和报告保持可追溯。 |
| `FIX-017` | `CLOSED` | Git blob/EOL 规范化、精确证据追踪、CI tracked-file 门禁及真实 LF/CRLF fresh clone 自测均通过。 |

## 6. 最终审查状态

`NEEDS_FIX`。仓库内实现、测试和 API provenance 缺陷已闭合，且当前没有未解决 `MUST_FIX`；但 `FIX-003` 仍是计划规定的 `SHOULD_FIX` 外部门禁。取得生产 2C4G 单槽位证据、提交四份必需 artifact 并获得最终候选 GitHub 全绿之前，不得批准发布。
