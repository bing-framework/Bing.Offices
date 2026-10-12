<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ENTITY-INDEPENDENT-PAGE-ANCHORS-20261005-001
AI_REVIEWED_AT: 2026-10-05T06:57:19+08:00

# 当前实现、计划差距与后续 TODO

## 结论与范围

本轮按用户最新要求，只检查计划、执行记录、当前源码和既有验证产物，新增本报告；没有修改业务代码、测试、API 基线或执行状态，没有重新构建、测试或运行 Office。实施代理已停止编辑，保留本轮开始之前及审计前已写入的改动。

按功能组统计，后续路线图仍有 **7 组能力尚未完整交付：F-001 部分实现，F-002～007 尚未实现**。另有 **2 组环境验收**。下面的当前阶段收口事项和设计考虑是这些功能的组成或验收条件，不重复加到功能数量中；各组规模不同，不换算完成百分比。

用户已批准的 Entity 重复页眉主计划，已有最终独立 PASS；透视来源契约与本机 Excel 生命周期阶段也已完成并独立 PASS。新 F-001 是后续扩展，不能把其未完成反推为原重复页眉计划漏项。CLIENT 的执行记录及重复文档两项历史收口 TODO 已处理，不再列为当前缺口。历史 NEEDS_FIX 报告保留，判断以各阶段较后的最终复核为准。

当前 F-001 的 NEEDS_FIX 针对正在实施的公开入口尚未贯通主流程及缺少验收，不代表既有已完成阶段被重新判为失败。本报告由主代理检查其前序参与过的共享规划代码，不冒充本阶段最终独立审查。

## 已完成能力与边界

| 能力 | 当前判断 | 边界及依据 |
| --- | --- | --- |
| 属性固定单元格、多动态字典、相对 Footer、业务实体布局 | 已实现 | 现有公共布局及两个 Entity Provider 路径；不重复登记为缺失 |
| Workbook/Entity 字典模板、名称列表、Entity 原生校验 | 已实现 | LOOKUP 后续阶段已接入；配置规则与原生规则都有真实调用链 |
| 填写区保护、级联字典 | 已实现，范围有界 | 同行两级；递归多级和跨行仍属可选扩展 |
| Workbook 分批 Unique、原生规则、动态列、跨 Sheet 关系 | 已实现，范围有界 | 不沿用初始阶段的拒绝清单误判后续实现 |
| DOM Workbook 失败工作簿、Entity 批次交付 | 已实现，范围有界 | NPOI/ClosedXML；Entity 批次仍先完整导入再分批投影 |
| 分组分页、分页小计名称、Entity 打印标题及实际重复页眉 | 已实现，最终复核 PASS | REPEATED-HEADERS/review-round2-final.md；名称依附中间分页小计的旧约定不变 |
| ClosedXML XLSX 透视 MVP | 已实现，范围有界 | 一个行字段、一个 Sum 值字段；NPOI 未提供完整缓存 |
| 透视来源合同 | 已完成，最终复核 PASS | SOURCE/review.md；来源公式和错误拒绝，携带原始坐标 |
| 本机 Excel 首次显示、编辑刷新、保存重开与有界资源探针 | 已完成，最终复核 PASS | CLIENT/review.md；结果不外推到 WPS、LibreOffice 或商业渲染 |

## F-001 五阶段对照

| 计划阶段 | 当前状态 | 实际情况 |
| --- | --- | --- |
| P1 公共模型、构建与快照 | PARTIAL | 三个 Fluent 入口、策略枚举、五个区域属性及 WithResolvedStart 传递已写入；LayoutBuilder 创建区域时仍未传入新描述，布局级冲突校验与模型测试未完成 |
| P2 共享行规划、导入识别 | PARTIAL | RowPlan 增加 PageFooter 行段，包括中间、末页及空明细；未验证。扫描器只有 Group/Page/Header，尚不识别每页 Footer |
| P3 Provider、名称及模板 | FAIL（尚未实现） | 两个 Provider 未消费新名称及清理策略；分段处理仍把非 GroupSubtotal/PageHeader 段视为 PageSubtotal，未增加 PageFooter 分支与占用预检 |
| P4 业务、资源及包消费者 | FAIL（尚未实现） | 未找到本阶段职责测试或真实三格式往返产物；既有消费者未演示新能力 |
| P5 API、文档及交付 | PARTIAL（仅计划与审批） | plan/api-approval 存在，execution 仍为启动模板；双 TFM API 差异、快照、追溯、迁移及最终门禁未完成 |

### 当前阶段 TODO：先贯通再验收

- [ ] T-001：LayoutBuilder 传入页眉/页尾前缀、策略和 PageFooter；补 PageBreak、marker、系列名称、静态锚点及公式边界校验；验证失败后构建器可复用、快照隔离、命名起点解析和旧调用兼容。
- [ ] T-002：完成 PageFooter 物理行规划与导入跳过，覆盖 GapRows、高度、组页同边界、无分页小计、首页/末页/空明细；缺失、重复和错位 marker 给出原始坐标。
- [ ] T-003：NPOI XLS/XLSX 与 ClosedXML XLSX 接入写入、逐页名称、Reject/RemoveObsolete 两策略、模板作用域与运行时占用预检；保留无关名称及输入模板。
- [ ] T-004：新增模型和各 Provider 独立业务/资源测试、同步异步完整往返及双 TFM NuGet 消费者；覆盖旧目标保全、流所有权、取消、提交失败、清理及观察次数。
- [ ] T-005：完成成员级 API 差异审批与双 TFM 快照、追溯、中文 XML、迁移示例、execution；修正新增枚举文件缺少 C# BOM 的偏差；运行受影响回归、Release 和 Category!=Large 门禁，再做最终独立复核。

### FIX-001：新公开配置在布局构建链中丢失

- 严重程度：HIGH；处理要求：MUST_FIX。
- 问题：新增 Fluent 方法把配置保存到 Builder，但 CreateListRegion 调用仍停在旧参数列表。
- 证据：ExcelEntityListRegionBuilder.cs:436、:461、:490；ExcelEntityLayoutBuilder.cs:357～369，构造调用末项仍为 builder.PageHeaderDefinition；ExcelEntityListRegion.cs 的新增参数均有默认值。
- 影响：普通公开布局构建会静默丢弃 PageFooter、独立名称与策略；PageHeaderNamed 只留下旧 PageHeader 内容，不能产生承诺的名称。
- 修复目标：完整传递新描述，并完成布局级冲突/公式校验。
- 修复要求：同时覆盖 ListRegion 与 ListRegionNamed、WithResolvedStart、快照及旧 API；不通过公开内部工具或生产友元绕过边界。
- 验证方式：职责测试从公开构建入口取得区域，完整断言新属性、失败后状态和命名起点解析结果；双 TFM 通过。

### FIX-002：共享新行段尚无 Provider 和导入消费者

- 严重程度：HIGH；处理要求：MUST_FIX。
- 问题：RowPlan 已生成 PageFooter，但 Provider 未处理此类别、未输出新名称或实施清理策略；导入扫描器未识别新 marker。
- 证据：src/provider-shared/ExcelEntityRowPlan.cs:153、:169；NpoiEntityExportExecutor.cs:475、:572 与 ClosedXmlEntityLayoutExecutor.cs:556、:575 的选择仍是 GroupSubtotal/Footer 或 PageSubtotal；ExcelEntitySubtotalImportScanner.cs:13 的 MarkerKind 仅为 Group/Page/Header。当前 Provider 中没有 PageFooterAnchorPrefix/PageHeaderAnchorPrefix 的消费。
- 影响：只修参数传递还不能交付；新段进入旧路径可能被错误处理，导入会把签字区当明细或报错。当前不能宣称新入口可用。
- 修复目标：接通所有段类型、名称预检及写入、模板清理和原坐标导入。
- 修复要求：三种格式行为一致；旧 PageSubtotalNamed 语义保留；全部系列预检之后才清理/重定位，失败不损坏已有输出及模板。
- 验证方式：真实工作簿断言每个名称的实际引用、全部数据/页尾/合并/样式，以及 marker 损坏时的完整结构化错误；覆盖三格式成功与失败。

### FIX-003：新增公开能力缺少当前候选验证

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：新增模型和共享规划没有本阶段职责测试或构建结果，API/消费者及交付未完成。
- 证据：tests 中未检索到 PageFooter、PageHeaderNamed；本任务当前仅 plan.md、api-approval.md 和启动模板 execution.md。前序 SOURCE/CLIENT 验证先于新改动，不能覆盖新候选。
- 影响：无法确认兼容性、默认行为、格式边界和资源合同；前序 PASS 不能当作当前工作区通过。
- 修复目标：按 P4/P5 完成职责、三格式、双 TFM、包消费者、API 及门禁证据，更新真实执行记录。
- 修复要求：不覆盖旧证据或 API 基线来掩盖差异，不省略资源失败场景，不引入无关 Large 矩阵。
- 验证方式：命令退出码、TRX、真实包产物、成员差异、符号追溯和最终独立复核相互一致。

### FIX-004：新增枚举文件缺少仓库规定的 BOM

- 严重程度：LOW；处理要求：SHOULD_FIX（仓库强制编码门禁）。
- 问题：ExcelEntityPageAnchorStaleNamePolicy.cs 是 UTF-8 无 BOM，而 .editorconfig 对 C# 明确要求 UTF-8 BOM。
- 证据：严格字节检查，该文件 424 字节，首字节为 namespace 文本，未以 EF BB BF 开始；LF 和末尾换行正确。
- 影响：不符合本计划及仓库的 C# 编码验收；不表示中文损坏或业务失败。
- 修复目标：仅为该新增文件补 BOM，不改内容或行尾。
- 修复要求：执行阶段精确修正，不做全仓库转换；本次只读检查不修改源码。
- 验证方式：严格 UTF-8 解码、BOM、LF、末尾换行及最小差异检查均通过。

## 后续功能 TODO：7 组

- [ ] F-001：独立逐页页眉与签字区锚点——进行中，先完成上面的 T-001～005。
- [ ] F-002：高级透视——多行/列字段、多值、筛选、Count/Average/Min/Max、排序和小计。当前 ExcelPivotTableDefinition 只有 RowField/ValueField，ClosedXmlReportWriter 固定 XLPivotSummary.Sum。
- [ ] F-003：NPOI XLSX 完整透视缓存——字段、记录、OOXML 关系及客户端生命周期验收；HSSF 单独定界。NpoiReportWriter.cs:23～24 当前明确拒绝。
- [ ] F-004：已有透视模板保留与更新——定义来源/缓存刷新、名称冲突及未知配置保真。ClosedXmlTemplatePreflight.cs:41～43 仍拒绝 Pivot 部件，ReportWriter.cs:50～52 拒绝已存在目标 Sheet；Excel 保存重开不是 Provider 模板保留能力。
- [ ] F-005：ExcelDataReader 前向失败工作簿——错误行暂存、可选输出、取消与提交合同。ValidateSupportedRequest 中 FailureOptions 非 None 仍抛 Unsupported("Failure Workbook")。
- [ ] F-006：端到端低内存 Entity 导入——固定字段、多区域、动态列、关系及分页/Footer 的前向读取。NpoiEntityBatchImporter.cs:47～48、ClosedXmlEntityBatchImporter.cs:48～49 仍先 ImportEntity 再 Deliver；分批回调不等于低内存读取。
- [ ] F-007：Entity 失败工作簿——固定字段/多区域标注、原始坐标、模板保真及失败产物生命周期。ExcelEntityImportOptions 当前只有资源、失败继续策略、校验模式和 Unsupported 策略，没有失败工作簿配置。

## 欠缺考虑或尚未闭环的设计 TODO

下面包含已考虑但缺少实现/证据的事项，不全是首次发现的问题；不据此扩大既有计划承诺。

- [ ] D-001：细分能力发现，区分创建透视、保留模板、刷新缓存、计算公式和渲染；错误明确来源、目标、名称及实际区域。
- [ ] D-002：组合资源预算覆盖来源面积、字段宽度、不同来源数、透视展开和页眉复制单元格数；文件字节上限不等于 DOM 内存上限。
- [ ] D-003：测量来源重复预检、缓存构建及序列化阶段取消；已有同源缓存共享，不再重复实现。当前来源枚举取消证据不能外推为全部阶段响应保证。
- [ ] D-004：补表格、保护、隐藏 Sheet、模板、多透视及单据布局组合合同；已有字典/级联名称冲突直测不重复登记为缺失。
- [ ] D-005：真实打印分页验收纳入纸张、缩放、行高、小计、页眉和签字区；PageBreak 的逻辑明细容量不保证物理打印页数。
- [ ] D-006：明确复制公式与绘图扩展策略；当前只保留公式文本，不重定位、不重算、不复制绘图。如扩大合同须独立设计。
- [ ] D-007：分批事务、后续失败与取消、重试幂等、跨 Sheet 关系状态及错误行暂存的内存/磁盘预算；随 F-005～007 落实。
- [ ] D-008：自动化客户端探针需要专用环境、外部墙钟超时与自身实例清理；同步 COM 可能阻塞，循环的 30 秒预算不能保证硬超时。README 文案已修正，隔离运行器尚未实现。
- [ ] D-009（可选）：递归多级级联、跨行父级引用和无客户端静态汇总；有业务需求时独立规划，不算已批准 MVP 漏项。
- [ ] D-010：新增每页 Footer 后重新核对最终 Footer 的公式覆盖区间。签字区内数字不能混入明细合计；连续 Sum 的布局校验必须纳入新行段，DetailSum 继续使用实际明细行范围。当前 CreateListRegion 的连续求和检查仍只提及 GroupSubtotal/PageSubtotal。

## 环境验收 TODO：2 组

- [ ] E-001：WPS、LibreOffice 分别验收首次显示、来源编辑刷新和保存重开；当前为 NOT_VERIFIED，不用本机 Excel 通过替代。
- [ ] E-002：有效 Aspose.Cells 许可证及固定字体/部署环境下验证目标函数重算、PDF、逐页图片及格式保真；拒绝无许可证测试不等于商业成功验收。

## 验证证据与限制

- SOURCE 最新 verification-summary：常规门禁 4,473/4,473，来源 follow-up 双 TFM 合计 132/132，Provider 合同合计 882/882、ClosedXML 合计 678/678；API contract 双 TFM 合计 22/22，API 成员差异为空，net6/net8 包消费者退出 0。这些是前序已验证候选，不是 F-001 当前工作区验证。
- CLIENT 当前摘要及原始结果：本机 Excel 四组成功；六组资源与四组来源取消均退出 0；两组历史 Failed 保留。工具双 TFM 构建通过，最终 Docs 19/19；最终独立 review 为 PASS。
- 本轮读取完整 F-001、SOURCE、CLIENT 计划及相关最终复核，检查当前源码调用链；Luna 仅协助核对既有验证产物，没有运行新测试。
- 本轮 git diff --check 退出 0；已有 ProfileFixtures.xml CRLF 提示保留。git diff --stat 为 64 个已跟踪文件、6,835 行新增、595 行删除，包含前序任务，不含大量未跟踪文件，不冒充本轮修改量。
- 本轮报告严格 UTF-8、无 BOM、LF、单个末尾换行检查通过。读取共享 RowPlan、列表 Builder、区域描述三个改动 C# 文件，均为 BOM/LF/末尾换行；新增策略枚举缺 BOM，记录为 FIX-004，未修改源码。
- 本轮只新增本报告；生产代码不修复，计划和 execution 不改写，旧审查证据不覆盖。

## 建议顺序

T-001 → T-002 → T-003 → T-004 → T-005，先将 F-001 完整闭环；再 F-002 → F-003/F-004 → F-005 → F-006/F-007。D-010 随 T-001/T-002 处理，D-001～004 随透视阶段，D-007 随导入阶段；环境验收独立推进。没有新增提交、推送、发布或外部通知。
