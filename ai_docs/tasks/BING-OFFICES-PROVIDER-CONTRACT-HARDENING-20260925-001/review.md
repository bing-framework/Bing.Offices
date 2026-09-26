<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-PROVIDER-CONTRACT-HARDENING-20260925-001
AI_REVIEWED_AT: 2026-09-25T21:26:00.1250230+08:00

# Independent Review

## Review Scope

本轮对照原 `plan.md`、补充 `gap-followup-plan.md`、`execution.md`、当前 Git Diff、真实生产/测试代码和最终证据完成独立复核。除保留既往 `FIX-001` 至 `FIX-003` 的关闭结论外，重点验收本轮 G1 Failure Workbook 失败边界、G2 Workbook 精确行数边界、G3 Stream best-effort 中途写失败和 G4 可审计最终候选证据。本轮只增加测试与报告，未修改生产实现。

## Statuses

```text
IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
API_STATUS: PASS
PACKAGE_CONSUMER_STATUS: PASS
PERFORMANCE_EVIDENCE_STATUS: REUSED_UNCHANGED_SCOPE
RESOURCE_EVIDENCE_STATUS: PASS
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL for CI/OS/font/production-capacity evidence
RELEASE_STATUS: BLOCKED_EXTERNAL
GOAL_STATUS: STOPPED_BLOCKED_EXTERNAL
OPEN_ACTIONABLE: 0
```

## Previous Findings

| ID | Requirement | State | Resolution evidence |
| --- | --- | --- | --- |
| FIX-001 | `SHOULD_FIX` | `CLOSED` | `CancellationContractTest.MidReadCancellation_ShouldStopAfterRealIo` 对三个 Provider 在首次真实读取后取消；文件目标合同另行验证 NPOI/ClosedXML 保留 sentinel 且无临时文件。 |
| FIX-002 | `SHOULD_FIX` | `CLOSED` | Converter、Validation、Exception、Cancellation 等共享断言使用 `FormatAssertionFailure`/`ProviderContractAssertions`；故意失败自测覆盖 scenario、provider、mode、format、expected profile、declared capabilities、actual outcome 七项诊断。 |
| FIX-003 | `SHOULD_FIX` | `CLOSED` | 执行、资源、集成、最终报告和生产符号映射已更新为当前方法与候选证据；历史数字均明确标识为历史，当前 ProviderContract 为每 TFM `192/192`。 |

## Follow-up Gap Findings

| Gap | State | Independent review conclusion |
| --- | --- | --- |
| G1 Failure Workbook 文件目标失败分支 | `CLOSED` | `FailureWorkbookFailureBoundaryContractTest` 覆盖 NPOI XLS/XLSX、ClosedXML XLSX，同步/异步和两种输出模式。48 个 case/TFM 分别检查新建/替换后的工作簿语义、提交失败及序列化预算失败时旧目标保全和临时文件清理；测试通过真实公共入口执行，不是只检查输出长度。 |
| G2 Workbook 精确 `MaxRows` 边界 | `CLOSED` | `ResourceLimitContractTest.MaxRowsAcrossSheets_ExactlyAtLimit_ShouldReturnCompleteWorkbook` 覆盖三个 Provider 的同步/异步共 6 个 case，两个 Sheet 合计恰好等于限制时返回完整实体且无错误；单 Sheet 超限用例同时收紧为空根实体。 |
| G3 Stream best-effort 中途写失败 | `CLOSED` | 公共合同流在实际写入 4 字节后抛错，覆盖 NPOI/ClosedXML、同步/异步和两种输出模式，验证异常传播、真实异步写入以及输入/输出 caller-owned。Core `DefaultFileExportCommitterTest` 另有实际临时文件写入后取消的目标保全和清理证据。 |
| G4 最终候选证据 | `CLOSED` | `artifacts/provider-followup-final/solution-test.exitcode` 为 0；独立解析 19 个 TRX 得到 total/executed/passed=`2378/2378/2378`，failed/error/timeout/aborted/inconclusive/notExecuted 均为 0。当前 API compare、Release build 和隔离 package consumer gate 均为 PASS。 |

## Verification

- 全解原始日志与 19 个 TRX 的项目计数一致：Core `204`/TFM、shared integration `30`/TFM、Docs `10`（net8）、NPOI `571`/TFM、MiniExcel `43`/TFM、ClosedXML `101`/TFM、各 Provider integration `30/9/4`/TFM、ProviderContract `192`/TFM；总计 `2378`，无失败、无跳过。
- Release solution build 记录为 exit code 0、0 warnings、0 errors；最终测试日志中的 Release 候选编译和运行均成功。
- 双 TFM API snapshot fresh compare 通过，批准的 `DestinationPath` 保持在公共快照中，无成员删除或签名漂移。
- 当前包源为 `artifacts/api-candidate`，隔离缓存为 `artifacts/consumer/cache-provider-followup`。独立重算五个 `Bing.Offices*.nupkg` SHA-256，feed/cache 逐包完全一致；两个普通 consumer 与两个 third-party consumer 构建和运行均 exit 0。
- `git diff --check` 通过；测试、最终报告、验证清单和生产符号映射均指向当前候选，而非历史结果。

## Reviewed Behaviors

- 三个 Provider 的 Workbook 级 `MaxRows` 共享预算一致，超限返回结构化 `ResourceLimit` 且不暴露此前 Sheet 的部分实体。
- NPOI/ClosedXML 的 `MaxCandidateErrorRows` 仅约束 `ErrorRowsOnly`，按 `(SheetName, RowIndex)` 去重；不同 Sheet 的相同 RowIndex 分别计数，`AnnotatedOriginal` 不受该预算限制。
- NPOI/ClosedXML 的 `DestinationPath` 复用 Core 原子提交；成功、替换/新建、取消、序列化/复制/提交失败、临时文件清理和 caller-owned Stream 边界均有直接证据。MiniExcel 在创建文件或写流前返回结构化 Unsupported。
- ClosedXML validation 覆盖五类比较规则的八个运算符、空值、1900/1904 日期、显式/同 Sheet/跨 Sheet 列表、安全 Custom、Unsupported 的 Report/Fail 及稳定执行顺序；未发现静默回退。
- 公共诊断、异常元数据、双 TFM API snapshot、Core IVT 和生产符号到测试方法的追溯均满足计划。

## Accepted Boundaries

- 通用调用方 Stream 保持 caller-owned，并明确为完整序列化后的 best-effort 复制；文件路径才提供原子替换保证。
- 命名范围、外部/失效引用及复杂 Custom 公式返回结构化 Unsupported，不引入公式引擎近似。
- 外部 CI、其他 OS、字体与生产容量证据保持 `BLOCKED_EXTERNAL`，不计入 `OPEN_ACTIONABLE`，也不触发自动修复。

结论：既往 `FIX-001` 至 `FIX-003` 及本轮 G1 至 G4 全部 `CLOSED`。当前生产实现、公共合同测试、API/package consumer 门禁和本地交付证据满足批准计划，`OPEN_ACTIONABLE=0`。
