<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-PIVOT-CLIENT-LIFECYCLE-20261005-001
AI_REVIEWED_AT: 2026-10-04T22:07:23Z

# 当前实现与计划差距检查

> 状态补充（2026-10-07）：本文下方的初始检查记录保留为历史证据。后续
> `execution.md`、`review.md` 和 `artifacts/verification-summary.json` 已完成
> CLIENT-001～004 的交付收口；继续执行时以本补充和各后续任务的最新执行记录为准。

> 最新核对（2026-10-07）：NPOI 缓存定义关系与文本类型标记修复后，Provider Contract
> `FullyQualifiedName~TemplatePivot` 在 net8.0、net6.0 各 70/70 通过。该结果只覆盖模板
> 保留/更新的 OPC 结构与 Provider 边界，不替代客户端刷新或跨客户端显示验收。

> P0 增量（2026-10-07）：离线对照 NPOI 2.7.4 原生 Count XML 后，数据字段的
> `baseField/baseItem` 已从 `0/0` 对齐为 `-1/1048832`。NPOI Pivot 直接职责双 TFM 各
> 73/73，Provider Contract 透视筛选双 TFM 各 113/113；真实 Excel Count 仍待有效会话复验。

## 当前状态补充

| 范围 | 当前状态 | 依据与边界 |
| --- | --- | --- |
| CLIENT-001～004 | 已完成 | Excel 双 TFM 同步/异步、资源矩阵、Docs、工具构建和独立复核均已记录；历史失败保留 |
| SOURCE-001～003 | 已完成 | 来源预检、类型/坐标错误、取消/资源、包消费者和 API 差异已由 SOURCE 执行记录收口 |
| F-001 逐页锚点 | 已完成 | Entity 独立页锚点任务已完成并通过 NPOI/ClosedXML 双格式职责测试 |
| F-002 高级透视 | 部分完成 | 公共模型、ClosedXML/NPOI 导出和结构/资源合同已实现；文本 Count 的真实 Excel 组合仍未通过 |
| F-003 NPOI 完整缓存 | 部分完成 | XLSX 缓存字段、记录和关系已接入；高级客户端 Count 结果仍待解决，XLS 保持拒绝 |
| F-004 模板透视保留/更新 | 部分完成 | NPOI/ClosedXML 已支持受限保留、部分/全部更新、缓存拆分和 `refreshOnLoad`；客户端刷新与任意扩展更新未承诺 |
| F-005 ExcelDataReader 失败工作簿 | 已完成范围内能力 | XLSX `ErrorRowsOnly`/`AnnotatedOriginal` 及批次入口已实现；XLS/XLSB 源保真仍结构化拒绝 |
| F-006 Entity 低保留导入 | 部分完成 | NPOI/ClosedXML 受限 Forward 与 `SummaryOnly` 已实现；DOM/实体/关系索引的端到端解析期预算仍未完成 |
| F-007 Entity 失败工作簿 | 已完成范围内能力 | 固定字段、多列表、模板/合并/分页/Footer、坐标和资源合同已覆盖；客户端显示仍待环境验证 |

当前可继续实施的代码待办只有三类：

- [ ] 为 ExcelDataReader 和 Entity 失败工作簿完成 Excel/WPS/LibreOffice 客户端显示验证；当前环境没有可复现的三客户端验收条件。
- [ ] 为 Entity Forward/批次路径设计并实现 DOM、实体集合、唯一索引、关系索引和错误暂存的端到端预算；现有 `SummaryOnly` 只降低返回结果保留量，不等同解析期低内存。
- [ ] 继续收口高级透视的真实 Excel 文本 Count 组合及其 NPOI/模板更新交付；不能以静态汇总或替换为数值 Count 规避合同。

本次核对后的排序为：

1. 先解决高级透视文本 `Count` 的真实 Excel 合同，或基于完整差异证据明确兼容边界；当前结构测试不能关闭该 TODO。
2. 在 `SummaryOnly` 之外设计 Forward 解析期低保留模型，分别预算输入暂存、实体物化、错误集合、Unique 和关系索引，再决定是否增加公共契约。
3. 收口模板更新的客户端刷新、任意扩展更新边界和外部 Provider 验收；当前 NPOI/ClosedXML 70/70 只证明结构链路。

## 结论与计数

本次只检查当前源码、计划、执行记录、Git 差异和既有验证产物，新增本报告；没有实施功能、修改历史报告或重新运行构建、测试及 Office。

当前已批准的重复页眉、透视 MVP 和透视来源契约，未定位到新的核心功能漏项。长期路线图仍有 **7 组功能扩展**；当前 CLIENT 阶段有 **2 组交付收口事项**；另有 **2 组环境验收**。后续设计考虑单独列出，不与功能数量相加，不据此换算完成百分比。

本报告的 NEEDS_FIX 针对当前交付记录与文档问题，不表示七组后续能力属于已批准 MVP 的回归缺陷。本轮检查也不替代 CLIENT 计划要求的最终独立复核。

## 实现对照

| 阶段 | 当前结论 | 依据与边界 |
| --- | --- | --- |
| 属性固定单元格、多动态字典、Footer、业务布局 | 已实现 | 当前布局与 Provider 入口；不重复列为缺失 |
| Workbook/Entity 字典模板、名称列表、Entity 原生校验 | 已实现 | LOOKUP 执行记录完成；配置与原生校验路径已接入 |
| 填写区保护、两级级联字典 | 已实现，范围有界 | 同行两级级联；递归多级与跨行引用属于可选扩展 |
| Workbook 分批 Unique、原生校验、动态列、跨 Sheet 关系 | 已实现，范围有界 | 后续分批阶段已完成；不能沿用首阶段历史拒绝列表判定当前实现 |
| DOM Workbook 失败工作簿与 Entity 批次交付 | 已实现，范围有界 | NPOI/ClosedXML；Entity 批次仍先完整导入再投影 |
| 分组分页组合、分页小计名称系列 | 已实现 | GROUP-PAGINATION、REPEATED-PAGE-ANCHORS；名称依附非末页的分页小计 |
| Entity 打印标题、实际重复页眉及导入识别 | 已实现并通过独立复核 | 主任务 review-round2-final.md 为 PASS；四个后续边界、元数据、资源、共享合同阶段已纳入复核 |
| ClosedXML XLSX 透视 MVP | 已实现，范围有界 | 单行字段、单 Sum、完整缓存；NPOI 仍明确拒绝 |
| 透视来源契约 | 已完成 | SOURCE 最终 review PASS；无公式/错误的来源合同及原始坐标 |
| CLIENT-001 真实客户端生命周期 | 本机 Excel 验证通过 | net6/net8 × 同步/异步四组，首次显示、编辑刷新、保存重开均正确 |
| CLIENT-002 兼容差异 | 已有事实与文档 | 保留两份历史 Failed；没有据此确认需要生产修复；不推定其它客户端通过 |
| CLIENT-003 资源矩阵 | 已有完整有界证据 | 六个 1K/10K 来源缓存场景、四个来源枚举取消场景，全部退出 0 |
| CLIENT-004 最终交付 | 已完成 | execution、独立复核、工具、资源和 Docs 证据已收口；WPS/LibreOffice 仍未验证 |

## 当前验证事实

- `artifacts/client-results.json` 包含四组最终成功结果与两组历史失败。本轮解析四份原始报告，确认状态 Passed、退出码 0、清理错误为空，输入与保存副本均存在。
- 四组 Excel 的完整区域均为 C2:D5，字段角色为 Warehouse/Amount，首次金额为 40/20/60，显式刷新与保存重开均为 45/20/65。本机版本为 Excel 16.0、Build 20430。保存重开属于客户端副本验证，不代表 Provider 已支持保留透视模板。
- 初始空矩阵的历史失败仍保留。被动读取取得完整结果不证明所有初始化问题已解决；Ready 和 CalculationState 也不能单独代替透视结果断言。原始 JSON 中中文标题实际正确，不把终端显示编码问题记为文件损坏。
- `artifacts/resources/results.json` 的十个场景均退出 0。同源四透视实际共享一个缓存，不同源分别建缓存；业务枚举计数不等于内部 DOM 扫描次数。取消发生在来源枚举阶段，不能推导为缓存构建和序列化阶段的时延保证。
- 新工具双 TFM Release 构建退出 0；最新 Docs TRX 为 19/19。本轮读取 TRX 和日志摘要，没有重跑命令。
- SOURCE 的最近常规证据为 27 份 TRX、4,473/4,473；双 TFM API 差异为空、隔离包消费者成功。这些是既有生产候选验证，不冒充本轮新测试结果。
- 本轮 git diff --check 没有空白错误；ProfileFixtures.xml 的既有 CRLF 提示保留，未转换。工具两份 C# 为 BOM/LF，ps1、csproj 为无 BOM/CRLF，严格 UTF-8 解码及末尾换行通过。

## 当前阶段 TODO：优先收口

- [x] C-001：更新 CLIENT execution，与四组 Excel、十组资源结果、19 项 Docs、双 TFM 工具构建和保留的失败记录一致；完成独立最终复核和本地任务状态收口。
- [x] C-002：删除 pivot-tables.md 中重复的“客户端显示与缓存生命周期”“多透视资源边界”章节，保留一份最新说明；复核运行示例与文档职责。

### FIX-001：执行记录与已有结果不一致

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：execution.md 仍为 IN_PROGRESS，正文称客户端正在验证、CLIENT-002～004 尚未完成；实际完整结果、资源报告和追溯表已存在，尚无 CLIENT 最终独立审查报告。
- 证据：execution.md:1、:15、:17 与 artifacts/client-results.json、resources/results.json、verification-summary.json。
- 影响：检查容易把已通过的客户端和资源工作重复列为未实现，无法形成可追溯的最终交付。
- 修复目标：据真实产物登记最终进度，保留历史失败、未知根因和未验证环境；完成独立复核。不得仅改状态来代替验收。
- 验证方式：计划矩阵、执行记录、原始产物、最终审查和运行时状态一致；没有生产修改时沿用 CLIENT 的工具/文档验证边界。

### FIX-002：透视文档重复追加

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：同一客户端生命周期与资源章节连续出现两份。
- 证据：docs/excel/pivot-tables.md:43、:62 与 :68、:87。
- 影响：维护时容易更新一份、遗漏另一份，能力说明与执行链接也可能继续分歧。
- 修复目标：合并为一份完整说明，不改变能力承诺，不删除历史失败事实。
- 验证方式：章节仅出现一次，链接与示例有效，Docs 职责和文本编码检查通过。

## 后续功能 TODO：7 组

- [x] F-001：独立逐页页眉与签字区命名锚点已实现；首页、末页、空明细、名称稳定性、局部名称和旧名称清理均有职责证据。
- [~] F-002：高级透视模型及 Provider 已实现；真实 Excel 的文本 Count 与部分多轴聚合仍未通过，保留为客户端限制与后续诊断。
- [~] F-003：NPOI XLSX 完整透视缓存已接入并通过结构/资源合同；高级客户端结果仍待收口，XLS/HSSF 继续明确拒绝。
- [~] F-004：已有透视模板受限保留与部分/全部更新已实现；客户端刷新、任意扩展更新和完整跨客户端验收仍未完成。
- [x] F-005：ExcelDataReader XLSX 失败工作簿和批次入口已实现；XLS/XLSB 源保真按计划结构化拒绝。
- [~] F-006：受限 Forward 与 `SummaryOnly` 已实现；真正解析期低保留和实体/唯一/关系索引预算仍需独立设计。
- [x] F-007：Entity 失败工作簿的结构、资源、取消和提交合同已实现；三客户端显示验证仍待环境条件。

这七组属于后续扩展，规模并不相同。F-005 与现有 DOM Workbook 失败产物、F-006 与现有批次回调是不同能力，不能合并计为已经完成。

### 源码定位

> 下表是初始检查时的限制定位，保留用于追溯；后续任务已改变其中部分源码。当前状态以本文“当前状态补充”和各任务最新 `execution.md` 为准。

| TODO | 当前限制的直接证据 |
| --- | --- |
| F-001 | `src/Bing.Offices.Abstractions/Bing/Offices/Entities/Lists/ExcelEntityListRegionBuilder.cs:357` 的 PageSubtotalNamed 与 :376 的 PageHeader；后者没有独立命名入口 |
| F-002 | `src/Bing.Offices.Abstractions/Bing/Offices/Exports/Reports/ExcelPivotTableDefinition.cs:54` 及 `src/Bing.Offices.ClosedXml/Bing/Offices/Exports/ClosedXmlReportWriter.cs:44` 的单字段 Sum |
| F-003 | `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:23` 的透视预检拒绝 |
| F-004 | `src/Bing.Offices.ClosedXml/Bing/Offices/Internals/ClosedXmlTemplatePreflight.cs:41` 拒绝 Pivot 部件；ReportWriter :50 拒绝已存在的目标 Sheet |
| F-005 | `src/Bing.Offices.ExcelDataReader/Bing/Offices/ExcelDataReader/ExcelDataReaderExcelImporter.cs:2279` 的 Failure Workbook 拒绝 |
| F-006 | `src/Bing.Offices.Npoi/Bing/Offices/Entities/NpoiEntityBatchImporter.cs:47`、`src/Bing.Offices.ClosedXml/Bing/Offices/Entities/ClosedXmlEntityBatchImporter.cs:48` 先完整导入再 Deliver |
| F-007 | `src/Bing.Offices.Abstractions/Bing/Offices/Entities/Import/ExcelEntityImportOptions.cs` 无失败工作簿选项；`docs/excel/09-providers.md:69`、:148 明确记录边界 |

七项源码边界另经 luna_worker 只读核对；该核对不替代 CLIENT 最终交付的独立审查。

## 环境验收 TODO：2 组

- [ ] E-001：WPS/LibreOffice 分别验证首次显示、来源修改后的刷新、保存重开及完整缓存/结果；本机没有相应成功证据，Excel 成功不能代替。
- [ ] E-002：有效 Aspose.Cells 许可证及固定字体/部署环境下验证目标函数重算、PDF、逐页图片与格式保真；无许可证拒绝测试不等于商业成功验收。

## 尚需补充考虑或证据的 TODO

以下是已有提示尚未闭环及新增设计提醒，并非全部此前完全未考虑，也不表示已经确认生产缺陷。

- [x] D-001：能力发现、来源/目标定位和创建/保留/更新边界已在能力位、结构化错误和模板依赖合同中区分；客户端计算/渲染仍单独验收。
- [~] D-002：已补来源面积、透视数量、同源缓存和页眉复制的定向资源证据；完整 DOM 峰值预算仍需随低保留设计建立。
- [~] D-003：已验证同源缓存复用、不同源隔离、来源枚举取消和失败后状态；缓存构建/序列化阶段的细粒度取消测量仍未完成。
- [x] D-004：表格、保护、隐藏 Sheet、模板、多透视及字典/级联辅助名称冲突已有组合合同。
- [ ] D-005：真实打印分页验收。纸张、缩放、行高、分组/分页小计和签字区共同影响打印；逻辑页容量不等于物理打印页高。
- [x] D-006：复制公式与绘图边界已明确为保留公式文本、不重定位/重算且不复制绘图对象。
- [~] D-007：失败产物和批次回调的取消、提交、关系及资源边界已覆盖；重试幂等和完整错误暂存预算仍需独立设计。
- [x] D-008：客户端探针已补外部超时、任务实例所有权和自建进程清理合同；同步 COM 调用的墙钟上限仍不作产品承诺。
- [ ] D-009（可选扩展）：递归多级级联、跨行父级引用及无客户端静态汇总。当前两级字典与缓存输出有明确边界；只有业务确有需求时单独规划，不计入七组已确认功能。

## 建议顺序

CLIENT/SOURCE 的历史收口已完成。当前顺序为：先明确 F-002/F-003 的真实 Excel Count 边界并收口 F-004 模板更新交付，再为 F-006 建立端到端低保留预算；F-005/F-007 的三客户端显示验证在环境具备时推进。公共模型先行、Provider 按明确支持范围实现，资源与组合合同随对应阶段补齐；E-001/E-002 独立推进，不用环境缺口冒充源码未实现。

历史 review-todo-20261005.md 保持原样，其 C-001/C-002 客户端与资源缺证据、A-001 页眉独立复核条目已有后续结果；不应直接照抄为当前 TODO。未提交、推送、发布、改版本号或发送外部通知。

## 2026-10-07 增量：公开 NPOI Count 形状回归

NPOI 公开导出测试已把数据字段 `baseField=-1`、`baseItem=1048832` 哨兵接入基础/高级同步与异步出口的最终 ZIP 断言。NPOI Pivot 双 TFM 各 `73/73`、Provider Contract 透视筛选双 TFM 各 `113/113` 通过；NPOI net6.0 首次并行构建的共享输出锁竞争已串行重跑通过。该证据加强了 F-002/F-003 的本地结构覆盖，但没有改变真实 Excel Count 客户端仍待有效会话复验的状态。

F-006 仍保持 `PARTIAL`：Forward/SummaryOnly 已减少实体列表和源行索引留存；共享字符串正文及固定索引改为磁盘分块读写。输入、错误、唯一值和关系键可按既有资源限制收敛；DOM 投影及解析期总预算仍未形成端到端契约，因此本轮没有增加未经审批的公共参数。
