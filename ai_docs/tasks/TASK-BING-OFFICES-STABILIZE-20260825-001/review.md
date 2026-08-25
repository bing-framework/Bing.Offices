<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: TASK-BING-OFFICES-STABILIZE-20260825-001
AI_REVIEWED_AT: 2026-08-25T18:27:09.9959654+08:00

# Round 9 独立复审报告

## 验收摘要

结论：`PASS_WITH_ISSUES`。

Round 8 已完整处理上一轮唯一 `MUST_FIX` 的 `FIX-005`。`NpoiExcelImporter.ImportSheet<T>()` 现在在 `NpoiWorkbookValidationPipeline.Validate(...)` 返回 `false` 后，无论 `ValidateMode` 是 `Continue` 还是 `StopOnFirstFailure`，都会回滚当前 Unique 事务并跳过 raw validation、转换和实体物化。新增的真实 XLSX 回归覆盖 `WorkbookRules` 和 `ConfiguredAndWorkbook`，输入同时违反原生显式列表规则且不能转换为 `int`，只产生 `WorkbookValidation`。

本轮未发现新的 BLOCKER、HIGH、MEDIUM 或未处理的 MUST_FIX/SHOULD_FIX。任务整体仍不得宣称 release-ready：旧 target runtime testhost 与 Office/LibreOffice/WPS 互操作尚无当前环境证据，NPOI DOM 的有界输入缓冲及大规模分配也仍是已知产品约束。这些是执行报告已诚实标记的 `PARTIAL / NO-RELEASE` 残余，不构成本轮代码修复阻断。

## Review 边界

- 计划：[plan.md](plan.md)。
- 执行记录：[execution.md](execution.md)，Round 8 为 `COMPLETED`，只表示当前 Review Fix 已执行完毕，不替代独立 Review。
- 上一轮 Review：[review.md](review.md) 的 `FIX-005` 是本轮首要验证目标。
- 审阅当前 Git Diff、Importer/Workbook validation pipeline、P0 测试、运行时状态和已有 performance evidence；未清理、暂存、提交、回滚或修改业务/测试代码。
- 当前工作区包含本任务历史的大范围未提交改动与 Round 7 internal collaborator 未跟踪文件。它们在本地编译中被包含，但最终变更交付时仍须确保纳入版本控制。

## 上轮 FIX 核验

| 问题 | 状态 | 独立证据 |
| --- | --- | --- |
| `FIX-005`：Workbook validation `Continue` 路径仍物化失败行 | RESOLVED | `ImportSheet<T>()` 在 `!workbookValid` 后无条件 rollback/continue；`Import_WorkbookValidationFailure_Continue_ShouldSkipMaterialization` 以真实 XLSX 覆盖两种 Workbook-enabled 模式；net8 P0 `27/27`、net8 Integration `12/12` 通过。 |
| 旧 `FIX-003`：协作者拆分和 Stream 对照 | PARTIAL，非本轮修复项 | Row materializer、Sheet writer、Document validator 已实际接入，旧 Exporter 第二实现已移除；Stream 仅有 benchmark-only 缓冲适配器对照。保留 no-release 风险，不重复创建无边界重构修复项。 |

## 计划验收矩阵

| Phase / 范围 | 结论 | 实际证据 |
| --- | --- | --- |
| Phase 1 P0 导入校验 | PASS | 四模式矩阵明确 Workbook 失败优先；失败行不进入 configured validation/转换/Unique；真实 XLSX 回归通过。 |
| Phase 2 Mapping / Profile | PASS | 当前 Diff 未改变已通过的方向、合并、JSON/XML validator 或 Profile 公共契约；生产搜索未见新增风险。 |
| Phase 3 API 收敛 | PASS | 新协作者均为 internal；生产源码未见新增 `Task.FromResult`、同步 `.Result`/`.Wait()` 或非测试 production IVT。 |
| Phase 4 内部职责拆分 | PASS_WITH_ISSUES | Importer row materializer、Exporter sheet writer、Loader document validator 均从主流程调用；Importer 拆分所致的控制流回归已修复。 |
| Phase 5 测试与集成 | PASS | 本轮独立 net8 P0 `27/27`、net8 Integration `12/12`；执行记录还包含 net8/net6 相关回归各 `183/183`。 |
| Phase 6 性能与资源 | PARTIAL | 同负载 Stream benchmark-only 缓冲对照已存在，但不等同于历史产品实现前后基线；NPOI DOM 高分配仍为已知约束。 |
| Phase 7 发布准备 | PARTIAL | `PARTIAL / NO-RELEASE` 结论仍正确；缺旧 runtime 执行与 Office/LibreOffice/WPS 互操作证据。 |

## 功能与真实接入

- `NpoiWorkbookValidationPipeline.Validate(...)` 在 `Continue` 时会继续遍历本行 native validation 以收集全部 Workbook 错误，并返回总体 `false`。
- `NpoiExcelImporter.ImportSheet<T>()` 随后立即执行 `RollbackRow()`（只在 configured validation enabled 时）并 `continue`，因此不会调用 `NpoiImportRowMaterializer.ValidateRawValues(...)` 或 `TryCreateItem(...)`。
- 回归测试断言无成功实体且唯一错误为 `WorkbookValidation`，直接防止 `ValueConversion` 回归。
- 验证模式矩阵将 `ConfiguredAndWorkbook` 的同单元格双失败期望更新为 Workbook 失败优先，符合计划 TASK-1.2 的 Workbook-first 与失败行边界约束。

## API、维护性与安全

- 生产源码搜索仅保留测试程序集的 `InternalsVisibleTo`，未发现新增 production IVT。
- 未发现 `Task.FromResult`、同步 `.Result` 或 `.Wait()` 残留。
- FIX-005 是单一控制流修复，没有新增公开 API、接口、配置开关或第二实现。
- NPOI/测试编译仍报告既有 obsolete attribute 和 netcoreapp3.1 依赖支持警告；本轮没有新增编译错误。

## 性能、测试与文档

- 历史 Stream 报告中的 1,000/10,000/100,000 行缓冲对照已明确标注 benchmark-only，未被误述为 streaming 或 zero-GC。
- 本轮独立执行：`ExcelP0RegressionTest` net8 `27/27`；`Bing.Offices.Tests.Integration` net8 `12/12`。
- Round 8 Executor 记录的 net8/net6 相关回归各 `183/183`、Release build 和 `git diff --check` 通过，作为补充证据；本轮未重复完整解决方案矩阵。
- `git diff --check` 未报告空白错误，仅有工作区既有 CRLF/LF 提示。
- `execution.md` 已记录 Round 8 根因、修复、测试和终态；其总体 `PARTIAL / NO-RELEASE` 文字描述与当前证据边界一致。

## 问题分级

### BLOCKER

无。

### HIGH

无。

### MEDIUM

无。

### LOW

- 当前环境未安装 net5/net7/netcoreapp3.1 testhost runtime，且没有 Excel/LibreOffice/WPS 互操作环境。它们限制整体发布验收，不影响 `FIX-005` 结论。
- Round 7 新增 internal collaborator 和任务文档仍处于未跟踪工作区状态；最终交付前需由提交者确认全部必要文件被纳入变更集。本 Reviewer 未执行 `git add`。

## 最终验收 Checklist

- [x] 已读取 `plan.md`、当前 `execution.md`、旧 `review.md`、实际源码、测试和 Git Diff。
- [x] 已验证 `FIX-005` 的无条件 rollback/continue 控制流。
- [x] 已验证真实 XLSX 回归同时覆盖 `WorkbookRules` 与 `ConfiguredAndWorkbook`。
- [x] 已独立运行 net8 P0：`27/27` 通过。
- [x] 已独立运行 net8 Integration：`12/12` 通过。
- [x] 已检查生产 IVT、伪异步/同步等待和 `git diff --check`。
- [x] 无未处理 MUST_FIX/SHOULD_FIX。
- [ ] 整体任务 release-ready：仍为 `NO-RELEASE`，需补齐运行时与 Office 互操作等环境证据。