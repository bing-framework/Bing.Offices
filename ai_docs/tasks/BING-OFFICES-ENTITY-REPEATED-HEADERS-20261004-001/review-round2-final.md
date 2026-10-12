<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001
AI_REVIEWED_AT: 2026-10-04T21:49:04Z

# Round 2 最终复核

## 结论

PASS。按主计划、FIX-008～FIX-012、四个后续页眉阶段、当前源码调用链和可读取的实际测试结果复核，没有发现当前未解决的 MUST_FIX 或 SHOULD_FIX。

旧 [review.md](review.md) 保持原样，仍是修复前的 Round 2 `NEEDS_FIX` 审查快照；它列出的 FIX-008～FIX-012 不代表当前仍未完成。本报告是对后续修复的独立复核，不改写该历史记录。主 [execution.md](execution.md:9) 也将“独立 Reviewer 尚需重新确认”列为最后收口事项。

## 计划验收矩阵

| 部分 | 状态 | 核对结果 |
| --- | --- | --- |
| P1 公共契约 | PASS | `PrintTitleRows` 与实际 `PageHeader` 分开；页眉配置、页码上下文、标记及列表表头约束进入布局契约。 |
| P2 Provider 与导入 | PASS | 共享行规划器生成页眉段；NPOI 与 ClosedXML 在实际导出前校验覆盖冲突，导入识别、跳过并校验页眉，错误保留工作表原坐标。 |
| P3 验证与交付 | PASS | 当前 Provider、合同 TRX 含重复页眉用例且全部通过；当前 SOURCE 候选记录 API 差异为空，net6.0/net8.0 NuGet 消费者成功。 |

## FIX-008～FIX-012

| 修复 | 状态 | 当前实现与直接测试 |
| --- | --- | --- |
| FIX-008 固定字段占用 | PASS | `NpoiEntityExportExecutor.ValidatePageHeader` 与 `ClosedXmlEntityLayoutExecutor.ValidatePageHeader` 检查后续页目标上的固定字段声明、既有内容和合并；声明值为空也纳入检查。`PageHeader_WithoutTemplateFixedCellOverlap_ShouldRejectBeforeOutput` 在两个 Provider 验证非空/空值冲突和目标流不变，模板及 marker 冲突也有独立用例。 |
| FIX-009 命名起点漂移 | PASS | 名称解析后 NPOI 与 ClosedXML 都调用共享 `ExcelEntityPageHeaderImportValidation.ValidateStart`。`PageHeader_NamedListWithoutHeader_ShouldValidateResolvedAdjacency` 验证 `Header(false)` 合法往返及命名起点移位后的结构化拒绝；NPOI 覆盖 XLS/HSSF 与 XLSX/XSSF。 |
| FIX-010 End 预留行 | PASS | 两个 Provider 以区域实际内容确定页眉列表终点；`PageHeader_EndReservedRows_ShouldRoundTripOrReportMissingMarker` 验证保留空行可往返，损坏 marker 仍按原始坐标拒绝。NPOI 覆盖两种格式。 |
| FIX-011 缺失 marker 坐标 | PASS | 共享 `ExcelEntitySubtotalImportScanner` 按分页容量、小计占用和页眉 marker 偏移推导应有位置。`PageHeader_WithoutPageSubtotal_MissingMarkerShouldReportExactCoordinate` 和分组边界测试断言准确位置；NPOI 覆盖 XLS/HSSF、XLSX/XSSF。 |
| FIX-012 组合与资源验收 | PASS | 动态/计算列物理范围、非法源高度、异步成功、输出写入/提交失败及取消均有直接用例。目标保全、暂存清理和观察次数由文件失败用例断言；大 End、资源限制和长扫描/表头核验取消各有确定性测试。 |

上述直接测试在当前 `SOURCE/artifacts/regular-tests` TRX 中可见，包括 NPOI `PageHeader_NamedListWithoutHeader_ShouldValidateResolvedAdjacency`、`PageHeader_EndReservedRows_ShouldRoundTripOrReportMissingMarker`、`PageHeader_WithoutPageSubtotal_MissingMarkerShouldReportExactCoordinate`、`PageHeader_DynamicAndCalculatedColumns_ShouldUseCompletePhysicalRange`、`PageHeader_InvalidHeightForHeaderMode_ShouldRejectBeforeOutput`、`PageHeader_AsyncExport_ShouldRoundTrip`、`PageHeader_FileCommitFailure_ShouldPreserveTargetAndCleanupTemporaryFile`，以及 ClosedXML 对应的命名起点、End、精确坐标、动态列、异步与提交/写入失败用例；这些结果均为 `Passed`。相关测试源码在 [NPOI 测试](../../../tests/Bing.Offices.Npoi.Tests/EntityRepeatedPageHeaderTest.cs:31) 与 [ClosedXML 测试](../../../tests/Bing.Offices.ClosedXml.Tests/EntityRepeatedPageHeaderTest.cs:28)。

## 后续阶段

- BOUNDARIES 完成 DESIGN-001/002：marker 保留值冲突、无表头 marker 列位置、Fluent 表头修改与合并边界，以及复制页表头篡改均接入真实 Provider 检查和导入坐标测试。
- METADATA 完成 DESIGN-003：多页复制隐藏行与锁定样式；源富文本、超链接、批注、数据校验、条件格式以及配置式校验重叠在写目标前拒绝；单页不复制时保留模板元数据。
- RESOURCES 完成 DESIGN-004：导入末行回溯、marker 扫描和复制表头核验均传递取消令牌；真实稀疏大 End、远端无关数据、资源上限和确定性取消用例覆盖此路径。
- CONTRACTS 完成 DESIGN-005：可复用 `VerifyPageHeaders` 通过公开 Abstractions 被 NPOI、ClosedXML 和第三方消费者调用；引擎特有 XLS/HSSF、XLSX/XSSF 结构断言仍在 Provider 测试中。当前合同 TRX 中两个 TFM 的 `PageHeaders_ShouldSatisfyReusableContract` 与 `PageHeaderInputLimit_ShouldRejectBeforeRead` 均通过。

## 验证证据

未启动构建、测试或 Excel；本复核只读取并解析已有 TRX/XML 结果及当前源码。

- `ai_docs/tasks/BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001/artifacts/regular-tests/` 的当前 Release TRX：NPOI Provider net6.0/net8.0 各 947/947，ClosedXML 各 339/339，Provider 合同各 441/441；重复页眉相关用例在这些结果中均为通过。
- 同一目录的 Provider 合同 TRX 可直接看到 `EntityLayoutReusableContractTest.PageHeaders_ShouldSatisfyReusableContract` 和 `PageHeaderInputLimit_ShouldRejectBeforeRead`，NPOI 与 ClosedXML 两个 Provider、两个 TFM 均通过。
- `ai_docs/tasks/BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001/artifacts/verification-summary.json` 记录当前验证候选的 net6.0/net8.0 API diff 均为空、候选与已验证身份一致；消费者使用版本 `2.0.0` 并分别输出 `package-consumer-ok`。
- 同一摘要记录的常规总计为 4473/4473。各历史页眉阶段的用例总数随阶段新增测试而递增；这些历史计数差异不是失败证据，本次以较后的当前 TRX 与当前候选身份记录为准。

## 已知边界

`PageBreak` 的容量只计算明细行；共享行规划器会把小计和插入的页眉计入工作表实际行位置并检查行索引边界，但不按纸张高度、页边距、缩放或行高计算可打印页高。`PrintTitleRows` 只设置打印元数据；没有真实 Excel/PDF/打印机分页验收。因此逻辑页不保证等同物理打印页，此项已作为 DESIGN-006 范围单列。

页眉复制保留公式文本，不计算公式，也不按目标页重新定位公式引用；绘图对象不复制。这些限制与当前计划及 [能力说明](../../../docs/excel/dynamic-columns.md:71) 一致，并非本轮修复遗漏。

未新增 FIX。旧 review 保留，未修改生产代码、测试、计划、执行记录、运行时或 API 基线。
