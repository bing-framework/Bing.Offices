<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-ENTITY-INDEPENDENT-PAGE-ANCHORS-20261005-001
AI_REVIEWED_AT: 2026-10-05T10:22:28+08:00

# Round 1 最终独立复核

## 结论

本轮独立核对了 F-001 计划、历史 review/FIX、当前生产调用链、直接职责测试、符号追溯、文档、最终候选 API/包消费者和最终门禁。没有发现当前未解决的 MUST_FIX 或 SHOULD_FIX，结论为 PASS。该结论仅覆盖 F-001 的独立逐页页眉、每页 Footer/签字区和命名锚点范围，不把 F-002～F-007 或 WPS、LibreOffice 等后续环境验收计入本阶段缺陷。

历史 `review.md`、进度报告和失败命令仍作为审计快照保留；本报告不改写它们。执行记录在本报告落盘前保持最终复核中的收口状态，后续 Finalizer 可据此完成本地任务终态。

## 计划与调用链

| 范围 | 状态 | 独立核对结果 |
| --- | --- | --- |
| P1 公共契约与快照 | PASS | `ExcelEntityListRegionBuilder` 的 `PageHeaderNamed`、`PageFooter`、`PageFooterNamed` 先校验后提交状态；策略枚举和五个 Provider SPI 属性经 `ExcelEntityLayoutBuilder`、构造函数及 `WithResolvedStart` 传递。旧入口和默认值保持兼容。 |
| P2 共享行规划与导入识别 | PASS | `ExcelEntityRowPlan` 统一规划明细、组小计、分页小计、每页 Footer、分页符和后续页眉，覆盖末页与空页；`ExcelEntitySubtotalImportScanner` 按已渲染 marker、GapRows、高度和页容量校验并跳过辅助行，保留原始坐标。 |
| P3 Provider 与模板 | PASS | NPOI 的 XLS/HSSF、XLSX/XSSF 以及 ClosedXML XLSX 均消费同一行规划和独立名称策略；先预检当前系列的作用域、单格引用、跨 Sheet/失效/公式/外部引用和边界，再清理合法过期名称。无关模板名称不进入清理集合。 |
| P4 业务、组合与资源 | PASS | 模型、边界、资源、marker、三格式真实往返和组合职责均有直接测试；动态/计算列、命名起点、原始错误坐标、原生失败整行跳过及 Unique 状态隔离已由当前组合测试覆盖。 |
| P5 API、文档与交付 | PASS | Docs 最终 19/19；Release、9 个 pack、4 个实际消费者运行及其 11 条 restore/build/run 命令均成功；最终 API 身份与批准候选一致。 |

## FIX-001～014

| Finding | 要求 | 状态与证据 |
| --- | --- | --- |
| FIX-001 | MUST_FIX | CLOSED。原报告所指的公开配置传递问题已解决：公共 Builder 保存的三个入口、策略及五个 SPI 属性经 `CreateListRegion`、区域构造和 `WithResolvedStart` 进入真实布局/解析链；模型双 TFM 职责产物通过。 |
| FIX-002 | MUST_FIX | CLOSED。原报告所指的 Provider 与导入消费者缺失已解决：共享 `ExcelEntityRowPlan`、`ExcelEntitySubtotalImportScanner` 和 `ExcelEntityPageAnchorNames` 被 NPOI/ClosedXML 真实导出与导入路径调用，覆盖空页、末页、分页边界、GapRows 和 marker 错位。 |
| FIX-003 | SHOULD_FIX | CLOSED。原报告所指的新公开能力缺少当前候选验证已解决：三格式职责、边界、资源、双 TFM、包消费者、API、Docs 和常规门禁均有当前候选证据。 |
| FIX-004 | SHOULD_FIX | CLOSED。原报告所指的新枚举文件 C# BOM 偏差已修正；策略枚举和受影响公开成员的 XML 注释、BOM/LF 已在当前 C# 编码清单中通过。 |
| FIX-005 | MUST_FIX | CLOSED。导出后读取实际 marker 单元格执行共享扫描，固定文本或工厂产生的扫描列重复 marker 在目标序列化前拒绝，非扫描列同文本仍可用。 |
| FIX-006 | SHOULD_FIX | CLOSED。NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML 的边界、资源、名称策略和损坏 marker 直接用例均纳入最终常规/职责结果。 |
| FIX-007 | SHOULD_FIX | CLOSED。双 TFM API 审批、成员级追溯、Docs 示例、普通消费者和第三方消费者均已绑定当前候选；没有以测试源码或旧包代替当前包证据。 |
| FIX-008 | SHOULD_FIX | CLOSED。NuGet 迁移示例的页眉/列表坐标已修正，Docs 围栏真实执行通过；历史 Docs 失败产物保留。 |
| FIX-009 | SHOULD_FIX | CLOSED。1001 页、255 字符名称边界、外部旧名称、动态/计算列、命名起点和原始坐标组合已由四次 Composition 命令覆盖。 |
| FIX-010 | SHOULD_FIX | CLOSED。最终 Release、pack、消费者和受影响常规门禁均由当前候选重新执行，未复用修复前候选替代验证。 |
| FIX-011 | SHOULD_FIX | CLOSED。外部名称夹具经公开关系 API 序列化/重载后验证；NPOI 原生校验已跳过计算列，并用动态列 `Key` 替代不存在的 `Property.Name`。 |
| FIX-012 | SHOULD_FIX | CLOSED。原报告所指的最终候选验收与执行记录已闭环：当前组合、常规门禁、Docs、Release、9 个包、11 条消费者命令、API 候选比较和编码清单均有终态记录，最终独立复核现已落盘。原生校验整行跳过、配置 Unique 回滚和原始坐标属于 FIX-009/FIX-011 的组合验收证据，不将其错归为本 Finding。 |
| FIX-013 | SHOULD_FIX | CLOSED。最终候选 API 比较退出 0；双 TFM 八个程序集相对已批准候选均为 `added=[]`、`removed=[]`，仓库基线身份已在比较通过后同步。 |
| FIX-014 | SHOULD_FIX | CLOSED。本报告补齐最终独立复核；阶段矩阵、验证摘要和历史证据均可追溯到当前候选。 |

## 直接源码核查

- `ExcelEntityListRegionBuilder` 的三个新入口在配置委托、重复声明、PageBreak、marker 和策略值验证完成后才写入字段；`ExcelEntityListRegion` 的五个 SPI 属性在布局创建和命名起点解析后仍保留。
- `ExcelEntityRowPlan` 将 PageFooter 放在边界小计之后、分页符之前，末页和零明细均有逻辑页 Footer；最终 Footer 的明细行段不包含签字区数值。
- NPOI `NpoiEntityImportExecutor.CreateRegionSheet` 对原生校验跳过 `Calculated` 列，按属性名或动态 Key 生成字段键；只有成功临时行才进入 `sourceRows` 的成功映射。`NpoiImportSheetExecutor` 对原生失败回滚 pending Unique 状态。
- ClosedXML `ReadRegion` 在原生校验失败时直接跳过整行，配置校验失败走 `RollbackRow`，成功行才提交并记录原始行；其行为与 NPOI 的成功实体坐标投影一致。
- NPOI 与 ClosedXML 的独立名称清理只枚举当前前缀的数字系列，拒绝歧义、跨 Sheet 和非法引用；`RemoveObsolete` 只删除通过预检且属于当前系列的旧单格名称，未发现影响无关模板名称的路径。

## 最终证据

- `artifacts/verification-summary.json`：27 组常规命令、4,805/4,805，通过/失败/跳过为 4,805/0/0；组合测试 4 命令、30/30；Release 退出 0；9 个 pack 全部退出 0；消费者 11 条命令、4 次实际运行全部退出 0。
- `artifacts/final-candidate/api-compare.command.json` 退出 0；`api-member-diff.json` 在 net6.0/net8.0 的八个程序集均无新增、无删除。原始 F-001 批准范围的 12 条 Abstractions 新增成员仍由 `api-approval.md` 和初始差异保留，未被零差异修复比较掩盖。
- `artifacts/regular-results-final.json` 使用最终 27 条命令和最终 TRX；旧 Docs 18/19 失败记录仍保留，最终 Docs 19/19 已进入最终集合。
- `artifacts/final-candidate/encoding-changed-files.json` 的 143 个改动 C# 文件全部通过严格 BOM/LF/EOF 检查；3 个非 C# 存量/生成例外与任务记录一致，未扩大为本阶段源码问题。`git diff --check` 退出 0。

当前 `OPEN_ACTIONABLE=0`，无须进入下一轮代码修复。外部 Office 客户端差异和 F-002～F-007 继续按各自任务处理。
