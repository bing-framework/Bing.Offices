<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BO-API-USABILITY-20261025
AI_REVIEWED_AT: 2026-10-26T23:09:29.9742903+08:00

# 易用性实现与 Round 1 修复复查

## 结论与范围

计划中的六项易用性改进及支撑文档/契约工作已实现；FIX-001、FIX-002 均已闭环。本轮未发现新的必须修复或建议修复的功能缺陷。保留一个可选的独立 NuGet 消费者验证增强，以及当前候选包身份、远端 CI/部署证据的边界，因此结论为 PASS_WITH_ISSUES，不代表发布门禁已经全部通过。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
OPEN_ACTIONABLE: 0
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE
RESOURCE_EVIDENCE_STATUS: VERIFIED_BOUNDARY（实际默认输入阈值；不包含 DOM 峰值或生产容量）
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL
RELEASE_STATUS: NOT_VERIFIED
GOAL_STATUS: STOPPED_BLOCKED

审查基于 HEAD `84958c1cc5eec5afbae51d76b0d07b4ab133b748` 的本地工作区，范围为本任务计划、实现及两个 P2 修复；不将其他未提交的文档、工作流或 BO-IMPORT-CI-OPTIMIZATION 任务视为本轮已经验收。读取工作区与暂存区差异，暂存区为空；相关生产/测试源码没有未提交修改。

本轮只更新审查材料，未修改生产代码、测试、plan.md 或 execution.md。原 NEEDS_FIX 报告完整保存在 [review-before-round1-fix.md](review-before-round1-fix.md)，其中问题和状态为修复前历史证据。未提交、推送、发布或发送外部通知。

## 计划验收矩阵

| 计划项 | 验收 | 实现与证据 |
| --- | --- | --- |
| U-01 正则示例 | PASS | 示例使用正确单反斜杠；Docs 测试从 Markdown 读取模型并执行有效/无效数据导入 |
| U-02 模板流所有权 | PASS | 文档明确默认关闭模板、leaveOpen=true 保留，以及预取消语义；既有所有权实现未改变 |
| U-03 默认语义 | PASS | 明确按行首错、100 列默认上限及坐标约定；多行/宽表有直接测试；缺省输入预算现已生效 |
| U-04 Core 注册 | PASS | AddBingOfficesCore 接入默认 CSV 服务；TryAdd 保留调用方实现；NPOI 注册复用入口 |
| U-05 单表六入口 | PASS | 通过私有根模型委托既有 Workbook 管线；完整保留数据、源行、全局错误、取消及流所有权 |
| U-06 统一状态及限额 | PASS | IsComplete/StopReason 与结构、资源、错误上限及普通行错误一致；跨 Sheet 限额诊断已修复 |
| U-07 文档、API、追溯 | PASS | 完整快速开始及成员测试映射存在；双 TFM 规范化公开成员快照通过；当前包身份另列待验证 |

计划内功能 7/7 完成；本轮没有新增公开 API 或删除旧成员。API 成员契约通过不能替代当前 nupkg 身份及独立安装验证。

## 已关闭的问题

### FIX-001：跨 Sheet 行限额诊断表名

- 原严重程度：MEDIUM / P2；原处理要求：SHOULD_FIX。
- 状态：CLOSED。
- 当前实现：`src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs:408` 使用 `resolvedSheet.Name ?? GetSelectorDescription(...)`。已解析 Sheet 返回真实名称，缺失 Sheet 保留 selector 回退；保持 ResourceLimit、0 行列、部分成功数据及单一错误。
- 直接测试：`tests/Bing.Offices.Tests/SheetBoundaryDiagnosticTest.cs:29` 的 `Import_RowLimitAtSheetBoundary_ShouldReportResolvedNameAndCompleteFailureSummary`。
- 覆盖：XLS/XLSX × Sync/Async × 索引/不同大小写名称/缺失名称，共每 TFM 12 例；断言全部错误字段、成功项和源行、完整八列失败摘要及流所有权。
- 当前证据：net6.0、net8.0 各 12 例通过；原错误 `#1` 和请求大小写不再替代真实表名。

### FIX-002：缺省输入预算接线

- 原严重程度：MEDIUM / P2；原处理要求：SHOULD_FIX；来源为既有 Workbook 管线遗留。
- 状态：CLOSED。
- 当前实现：`NpoiExcelImporter.cs:209`、`:333` 在复制前取得 `request.ResourceLimits ?? new ExcelResourceLimits()`，再读取 MaxInputBytes。省略配置启用 128 MiB；显式对象的 MaxInputBytes=null 继续关闭该限制。
- 直接测试：`tests/Bing.Offices.Tests/DefaultInputLimitTest.cs:38` 的 `Import_DefaultBudget_ShouldEnforceBoundaryAndHonorExplicitNull`，以及 `:77` 的 `Import_RealFile_ShouldSupportDefaultExactDisabledAndRejectOverflow`。
- 覆盖：Workbook/单表 × Sync/Async 的实际 128 MiB 边界、超出一字节和显式关闭，共 12 例；真实 XLS/XLSX 文件的默认、精确额度、显式关闭和超限，共 8 例。每 TFM 合计 20 例。
- 当前证据：两 TFM 均全部通过。默认超限在解析前抛出完整 ResourceLimitExceeded / Import / NPOI / Open 诊断；精确额度及显式关闭能够到达终点哨兵。异步合成流拒绝同步 Read，真实文件流启用异步选项，调用方流保持打开。
- 边界说明：实际供给并复制了默认额度附近的字节；终点哨兵用于避免解析大型无效 DOM，不声称 128 MiB Office DOM 可成功加载。依赖旧缺省无限制行为的调用方需要显式调整或关闭限额，迁移文档已说明。

## 剩余欠缺及发布边界

### OPT-001：新增 API 的独立 NuGet 消费覆盖

- 严重程度：LOW；处理要求：OPTIONAL；状态：DEFERRED。
- 证据：`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj:21` 使用项目引用；QuickStartConsumerTest 的 CSV 编译过滤 NPOI，但仍借用测试宿主引用。现有 net6/net8 PackageReference 消费者使用旧 Workbook 入口，未调用 AddBingOfficesCore、ImportSheet 或新增结果摘要。
- 当前影响：源码层运行和 API 契约已验证；尚未证明从干净消费者环境恢复当前包后，新 API 及 Core-only CSV 的依赖闭包完整。未发现实际缺失依赖，故不作为生产缺陷。
- 建议：后续任务“新增 API 包消费者验证”，从当前源码打包到隔离本地源，在干净 net6/net8 消费项目中调用上述 API；CSV 项目只安装 Core，并检查完整输出和结构化失败结果。
- 延后原因：原 Review 已归为可选增强，Round 1 明确未纳入；本次只审查，不扩大实施。它不阻止本轮实现验收，但正式发布前应补充包消费证据。

### 当前候选身份与外部门禁

| 项目 | 状态 | 证据边界与后续 |
| --- | --- | --- |
| 修复后完整候选包身份 | DEFERRED / NOT_VERIFIED | baseline 的 candidateIdentity 仍指向旧提交 7496bd40… 和历史源码/产物；execution 明确 P2 修复后未重打包。双 TFM Unit 只比较公开成员，不验证当前源码清单、nupkg 内容或安装结果。正式发布候选验证时应重新生成当前候选包并运行既有身份检查，不为通过检查而放宽或覆盖批准的 API 成员基线 |
| 对应候选的远端 CI | BLOCKED_EXTERNAL | YAML 配置有测试与包消费者任务，但未取得与当前候选对应的绿色 run 证据；本次未查询或触发远端 CI，不能由配置文件推断通过或失败 |
| 生产资源/并发部署证据 | BLOCKED_EXTERNAL | 默认输入阈值和本地正确性测试不等于部署环境的 DOM 峰值、吞吐量或并发容量验证；保持原计划外部门禁边界 |

候选身份检查 CLI 已有实现，当前 CI 未调用完整身份比较入口。这属于后续发布候选验证的证据/流程补齐，不是本轮新确认的功能故障。建议在“当前候选发布验证”任务中处理；当前 RELEASE_STATUS 保持 NOT_VERIFIED。

## 验证与证据复用

影响范围：本轮 ChangedFiles 仅审查报告；被复查生产范围为 NPOI 输入复制预算与跨 Sheet 限额诊断。ChangedPublicContracts=无，ChangedTFMs=无，ChangedBuildPackaging=无，ChangedBenchmarkHarness=无；本轮风险 LOW，被审查修复风险 MEDIUM。

本轮实际执行：

```powershell
foreach ($tfm in @('net6.0', 'net8.0')) {
    dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f $tfm --no-restore --filter 'FullyQualifiedName~DefaultInputLimitTest|FullyQualifiedName~SheetBoundaryDiagnosticTest|FullyQualifiedName~PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot' --logger "trx;LogFileName=review-$tfm.trx" --results-directory TestResults/usability-review-r2
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
```

| 层级 | 证据 | 结果 |
| --- | --- | --- |
| L0 | 目标测试及其生产依赖 Release 构建；git diff --check | 通过 |
| L1 net6 | TestResults/usability-review-r2/review-net6.0.trx | 33/33，0 失败，0 跳过 |
| L1 net8 | TestResults/usability-review-r2/review-net8.0.trx | 33/33，0 失败，0 跳过 |
| 本次合计 | 32 个修复用例 + 1 个 API 快照用例，分别执行两个 TFM | 66/66 |
| L2 复用 | usability-fix-r1/unit-net6.trx、unit-net8.trx | 各 956/956 |
| L3 复用 | usability-fix-r1/integration-net6.0.trx、integration-net8.0.trx | 各 38/38 |
| Docs 复用 | usability-fix-r1/docs-net8.trx | 11/11 |

已实际解析上述 5 份前轮 TRX，合计 1,999/1,999，0 失败、0 跳过；此数量是历史完整套件记录，不冒充本轮新执行结果，也不与本轮 66 次重复相加。相关业务、测试和 TFM 未出现影响复用的变化；本轮未重新跑 L2/L3 全套、全 solution 或重型资源矩阵。

前轮 docs-check 的 6 处示例一致性证据沿用 execution 记录；本次未重新执行。未重新打包，未重新捕获/放宽 API 基线。luna_worker 独立只读核对包消费者、快照身份及 CI 证据边界；主代理负责真实修复行为、直接测试和最终结论。

## 本轮状态归档

- Task / Round：BO-API-USABILITY-20261025 / Round 1 修复后的复查。
- Implementation：计划项 7/7；原 P2 2/2 CLOSED。
- Release Gates：本轮未对全部发布门禁重新计数；当前候选身份、远端 CI、部署证明未完成，不能声明发布 PASS。
- Completed：两个 P2 验收、双 TFM 66 次复验、历史 1,999 次记录复核、原报告归档。
- Open Actionable：无。
- Blocked Approval：无新增。
- Blocked External：对应候选的远端 CI、部署资源/并发证明。
- Not Applicable：性能 before/after；无性能提升声明。
- Accepted Limitations：NPOI DOM 阶段仍同步；输入字节限额不等于 DOM 峰值上限。旧 Workbook.IsSuccess 继续兼容原有仅检查 Errors 的行为，手工构造不一致结果时需结合 IsComplete；高级配置继续使用 Workbook Request。
- Verified Boundaries：默认输入预算精确阈值及超限拒绝、显式关闭、真实文件成功/失败、实际表名及缺失回退。
- Deferred：OPTIONAL 新 API 独立 NuGet 消费者；正式发布候选身份重新验证。
- No-Progress Check：CHANGED，本轮取得双 TFM 新证据并正式关闭两个旧 Finding。
- Next Action：STOP。没有可执行缺陷，不启动下一轮自动 Fix。
