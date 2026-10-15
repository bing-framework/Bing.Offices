# 任务文档索引

复核日期：2026-10-14

技能任务归并更新：2026-10-15；其余任务记录状态仍沿用上述复核日期。

本索引按任务最后保留的执行记录和独立审查结论整理，并对照当前源码、测试、CI 与技能消费文档核验关键状态。历史测试数字不代表本次重新运行；`artifacts/` 下的机器报告和原始证据未纳入清理，也未修改。

## 阅读规则

- 未完成任务的 `plan.md`、执行记录和审查结论保持原样；本索引不替代或改写其目标、背景、验收标准和独立 Review。
- 已完成任务的可复用内容优先归入对应功能或使用说明；本索引保留任务 ID、承载位置、验证限制及后续事项，不重复维护总结。
- `artifacts/` 是忽略的本地验证证据目录，保留原样；不视为可删除的重复任务文档。
- 旧任务的 `PARTIAL`、`BLOCKED` 或 `NEEDS_FIX` 是其最后记录状态。除非后续任务明确复验并关闭相同验收项，不因代码中出现相应类型或测试而推定旧计划已完成。

## 当前任务

| Task | 记录状态 | 当前核对与处理 |
| --- | --- | --- |
| [BO-RC-20260908-002](BO-RC-20260908-002/plan.md) | `NEEDS_FIX`，`FIX-003` 外部门禁未关闭 | 本地 CI 已移除独立 API Snapshot compare；Package consumers 保留，两个 NuGet.Config 使用跨平台 feed 路径。仍缺生产入口 `2C4G` 实际并行度 `1` 证明及当前候选 GitHub 全绿。保留计划、执行和 Review；最终报告已追加当前配置复核。 |
| [BO-RC-20260907-001](BO-RC-20260907-001/plan.md) | 本地实现 `PASS`；发布条件未闭环 | 扩展方法直接覆盖和双 TFM 证据已通过；生产入口与远端 CI 条件由后续 09-08 任务继续跟踪。保留批准、基线、契约与符号追溯资料。 |

## 未完成及历史待办

以下任务最后记录仍有部分计划项、独立 Review 问题或发布门禁未关闭。后续任务承接了其中部分实现，但没有逐项重审并关闭这些原始计划，因此保留所有原任务文档。

| Task | 最后记录状态 | 未闭合范围摘要 |
| --- | --- | --- |
| [bing-offices-excel-import-export-enhancement-v2](bing-offices-excel-import-export-enhancement-v2/plan.md) | `PARTIAL` / `PASS_WITH_ISSUES` | 仍有 ErrorRowsOnly 保真、API 成员级审批、Office 互操作及格式基线等原计划缺口。 |
| [BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001](BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001/plan.md) | `PASS_WITH_ISSUES` | `FIX-006` 的集成、文档与 Benchmark 完整性仍为部分完成。 |
| [BING-OFFICES-PRE-RC-CLEANUP-20260901-001](BING-OFFICES-PRE-RC-CLEANUP-20260901-001/plan.md) | `BLOCKED` / `No-Go` | 最后审查仍要求补齐完整性能/资源、目标框架与 clean-clone 交付证据。 |
| [BING-OFFICES-RC-CLOSURE-20260905-001](BING-OFFICES-RC-CLOSURE-20260905-001/plan.md) | `PARTIAL` / `NEEDS_FIX` | 原报告保留 API、预算和 RC 发布门禁；后续承接情况未对本任务逐项复审。 |
| [BING-OFFICES-RC-HARDENING-20260904-001](BING-OFFICES-RC-HARDENING-20260904-001/plan.md) | `BLOCKED` / `NEEDS_FIX` | `FIX-004/005` 与完整资源、API、发布证据未在该任务 Review 中关闭。 |
| [BING-OFFICES-RC-HARDENING-20260906-001](BING-OFFICES-RC-HARDENING-20260906-001/plan.md) | `IN_PROGRESS` / `BLOCKED` | 任务状态未收口；其发布证明工作由 09-07、09-08 后续任务承接，原件保留作历史追溯。 |
| [BING-OFFICES-RELEASE-HARDENING-20260825-01](BING-OFFICES-RELEASE-HARDENING-20260825-01/plan.md) | `PARTIAL` / `NEEDS_FIX` | 原计划的 API、结构拆分、性能和发布证据未全部完成。 |
| [BING-OFFICES-RELEASE-HARDENING-20260827-001](BING-OFFICES-RELEASE-HARDENING-20260827-001/plan.md) | 本轮执行完成 / `NEEDS_FIX` | 执行完成仅代表指定 Review Fix 范围结束；原任务尚有未关闭 Fix。 |
| [BING-OFFICES-RELEASE-HARDENING-20260831-001](BING-OFFICES-RELEASE-HARDENING-20260831-001/plan.md) | `PARTIAL` / `NEEDS_FIX` | 最后 Review 记录 API、预算和交付门禁未闭环；其中自动快照比较属于当时方案，当前 CI 决策见 09-08 任务补充。 |
| [TASK-BING-OFFICES-20260821-MAPVAL-V2](TASK-BING-OFFICES-20260821-MAPVAL-V2/plan.md) | `PARTIAL` / `PASS_WITH_ISSUES` | Plan compiler、诊断/别名与错误路径、关系规则及性能/资源矩阵仍有原计划缺口。 |
| [TASK-BING-OFFICES-STABILIZE-20260825-001](TASK-BING-OFFICES-STABILIZE-20260825-001/plan.md) | `PARTIAL` / `PASS_WITH_ISSUES` | 发布仍为 `NO-RELEASE`；Office 互操作、旧运行时和部分性能/消费者验证未闭环。 |

## 已整合任务

| Task | 状态与后续事项 |
| --- | --- |
| [BO-SKILL-CHINESE-COMMENTS-20260914-001](../../docs/ai-workflow/CHINESE-COMMENTS.md#交付与维护状态) | 原计划交付已完成，交付范围、原验证记录与限制、来源迁移及待同步事项已合入技能使用说明。项目 `.agents`/`.claude` 副本彼此一致，但与 awesome-copilot 规范源不同；后续同步仍待确认。四种 CLI 的认证交互未实测。 |

## 当前实现索引

- NPOI 导入/导出共享列计划：`src/Bing.Offices.Npoi/Bing/Offices/ExcelColumnPlan.cs`，导入和导出主链均引用该计划。
- Workbook 验证区域索引：`src/Bing.Offices.Npoi/Bing/Offices/Imports/ValidationRangeIndex.cs`，对应大矩形、离散范围和候选缩减测试。
- 公开扩展声明覆盖门禁：`tests/Bing.Offices.Tests/PublicExtensionCoverageTest.cs`；后续成员批准删除后，当前计划追踪的是 81 个公开扩展声明，而非旧报告的 85 个。
- CI 当前保留 net6/net8 Package consumers 和公开 API Unit 合同测试；独立 API Snapshot compare 已不在 workflow 中。
- 仓库内 `SemaphoreSlim(1,1)` 证据来自 ResourceProbe，不是生产部署入口证明；生产入口和 GitHub 最新候选状态仍以 09-08 Review 的外部验收为准。

## 整理范围

2026-10-14 已将原计划完成的中文注释技能任务文档合并为 `summary.md`，删除其旧 `plan.md` 与 `execution.md`。2026-10-15 经用户确认，将该总结的独有内容合入上述技能使用说明后删除 `summary.md`；任务正文目录由 14 个减为 13 个。

其余 13 个任务的 131 份正文保留原有目标、背景、验收、执行历史、Review 和专项报告；本次没有删除任何任务 artifacts 或其他任务文件，没有修改技能副本或执行安装。
