<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BING-OFFICES-CLOSEDXML-PROVIDER-20260922-001
AI_REVIEWED_AT: 2026-09-24T13:17:36.5352824+08:00

# 第三轮独立代码审查

## 结论

本轮对照计划、执行记录、最新源码、Git Diff 和直接测试重新审查。上一轮的 `FIX-001`、`FIX-002`、`FIX-003` 均已关闭，未发现新的本地可执行缺陷，`OPEN_ACTIONABLE=0`。

ClosedXML Unit `74×2`、ClosedXML Integration `4×2` 全部通过；公共 Integration 为 `38/39×2`，每个 TFM 唯一失败均为未批准的 Public API snapshot。该失败属于 `BLOCKED_APPROVAL`，不构成实现回归。当前实现审查通过，但发布状态仍受审批和外部环境门禁约束，因此总体状态为 `PASS_WITH_ISSUES`，不得将其表述为发布通过。

## Findings

### FIX-001：统一物理列布局与稀疏列宽

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 状态：`CLOSED`
- 实现证据：`ClosedXmlExportColumnPlanner.Create` 先锚定具有显式 `ColumnIndex` 的固定列，再按声明顺序将未配置固定列填入最靠前空槽；动态列的显式物理索引和相对定位继续基于同一物理布局规划。
- 宽度证据：`ClosedXmlExcelExporter.ApplyColumnWidth` 只遍历规划结果中的实际 `PhysicalColumnIndex`，不再按 `1..max` 修改稀疏布局中的空隙列。
- 测试证据：`MixedFixedColumnIndex_ShouldFillRemainingSlotsAndSkipSparseWidthGaps` 真实读回“一列显式 B + 一列默认 A + 动态 D”，并验证 C 空隙宽度未被覆盖。
- 结论：上一轮指出的混合固定索引错位和空隙列宽污染均已解决。

### FIX-002：普通 Workbook 异步导出的 Admission 生命周期

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 状态：`CLOSED`
- 实现证据：`ClosedXmlExcelExporter.ExportAsync` 的 admission lease 仅包围 `ExportWorkbookCore`；生成 Workbook 完成并释放 gate 后，才进行目标流的 `CopyToAsync` 和 `FlushAsync`。
- 测试证据：`ExportAsync_ShouldReleaseAdmissionBeforeBlockedDestinationCopy` 阻塞第一个目标复制期间，第二个真实导出仍可取得 admission 并完成。
- 结论：外围目标 IO 不再占用 ClosedXML DOM 并发名额。

### FIX-003：职责级与 Cross-provider 证据矩阵

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 状态：`CLOSED`
- Mapping Plan：`MappingPlanFactory_ShouldIsolateTypeConfigurationAndDynamicKeysAndRecoverAfterFailure` 覆盖类型、配置、动态列隔离以及工厂失败后的状态恢复；既有测试覆盖命中、方向隔离和重复使用。
- 导入边界：`ImportErrorCell_ShouldReportConversionForIncompatibleTargetType` 验证 Error Cell 绑定不兼容类型时产生结构化转换错误；`ImportAsync_NonSeekableInput_ShouldRoundTripAndKeepSourceOpen` 验证 non-seekable 异步输入及所有权。
- Cross-provider：`MappingContract_ShouldMatchAcrossAllXlsxProviders` 覆盖 NPOI、MiniExcel、ClosedXML 的共同 Mapping 语义；`FormulaCachedValueContract_ShouldMatchNpoiAndClosedXml` 覆盖数值和布尔缓存值，并固定字符串公式的 provider policy 差异。
- 结果证据：ClosedXML Unit 从 `70×2` 增至 `74×2`；公共 Integration 中 cross-provider contract 为 `10×2`，并包含在本轮 `38×2` 通过项中。
- 结论：上一轮列出的最小职责证据缺口已经补齐，未发现因测试替身掩盖真实 Workbook 行为的情况。

## 非开放门禁与限制

| Area | Status | Evidence / Boundary |
| --- | --- | --- |
| API baseline | `BLOCKED_APPROVAL` | 公共 Integration `38/39×2`；唯一失败为 Abstractions 新增 11 个 additive 成员及 ClosedXML assembly 尚未进入批准 baseline。执行 Agent 不得自行更新。 |
| Entity Relations 的 `MaxErrors` | `BLOCKED_APPROVAL` | Entity API 当前没有资源限制输入；ClosedXML 与 NPOI 的 Entity relation collector 均没有可由现有契约传入的 `MaxErrors`。需要公共 API 决策。 |
| 500K/1M 与正式容量阈值 | `BLOCKED_APPROVAL` | 继续需要维护者批准资源预算和观察窗口。 |
| Linux、无字体、AutoFit 峰值 | `BLOCKED_EXTERNAL` | 当前 Windows 本地证据不能替代目标外部环境验证。 |
| Entity List Region 动态列 | `ACCEPTED_LIMITATION` | v1 明确 Unsupported，并在执行前拒绝。 |
| Chart、Pivot、Image/Table 创建、Failure Workbook、Workbook 原生 Validation、XLS/XLSM/Macro、完整公式引擎 | `ACCEPTED_LIMITATION` | 属于已批准的 v1 边界，不新增 Rich capability bit。 |

## Phase 验收矩阵

| Phase | 状态 | 审查结论 |
| --- | --- | --- |
| 现有契约修复 | `PASS` | 混合固定列、动态稀疏物理列、Header/Body/Merge/Width 的统一布局已有实现与真实读回证据。 |
| 布局与 Entity | `PASS_WITH_BLOCKED_APPROVAL` | 行高、多 List Region 和 Relations 已接入；Entity `MaxErrors` 仍需公共 API 决策。 |
| 资源预检与 Admission | `PASS_BASIC` | Sheet/Column/Cell DOM 前预检、队列与并发边界、取消和 gate 释放均有直接测试。 |
| 职责测试与契约 | `PASS` | ClosedXML Unit `74×2`、Integration `4×2`、cross-provider contract `10×2` 通过。 |
| 验证与收口 | `PASS_WITH_BLOCKED_GATES` | 本地可执行正确性工作已闭环；批准 baseline、资源预算和外部环境仍阻塞发布。 |

## Fresh 验证

- `dotnet test tests/Bing.Offices.ClosedXml.Tests/Bing.Offices.ClosedXml.Tests.csproj -c Release --no-restore -v:minimal`：PASS，`74×2`。
- `dotnet test tests/Bing.Offices.ClosedXml.Tests.Integration/Bing.Offices.ClosedXml.Tests.Integration.csproj -c Release --no-restore -v:minimal`：PASS，`4×2`。
- `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release --no-restore -v:minimal`：`38/39×2`；每个 TFM 唯一失败均为预期的 API snapshot `BLOCKED_APPROVAL`。
- `git diff --check`：无 whitespace error；仅有 dirty worktree 中既有行尾转换警告。

## 审查状态

- `IMPLEMENTATION_STATUS: PASS`
- `TEST_STATUS: PASS`
- `PERFORMANCE_EVIDENCE_STATUS: PASS_WITH_BLOCKED_APPROVAL`
- `RESOURCE_EVIDENCE_STATUS: PASS_BASIC`
- `EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL`
- `RELEASE_STATUS: BLOCKED`
- `GOAL_STATUS: STOPPED_BLOCKED`
- `OPEN_ACTIONABLE: 0`

依据自动循环停止规则，本轮不得因 `BLOCKED_APPROVAL`、`BLOCKED_EXTERNAL` 或 `ACCEPTED_LIMITATION` 自动进入下一轮修复。
