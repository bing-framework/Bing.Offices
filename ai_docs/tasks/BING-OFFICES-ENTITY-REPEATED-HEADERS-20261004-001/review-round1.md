<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001
AI_REVIEWED_AT: 2026-10-04T07:35:17Z

# 对照计划审查与 TODO

## 结论

两种模式及主要正常调用链已经实现，但尚未完整满足计划。剩余 6 项行为/契约修正和 1 项验证交付收口。此次只审查，不修改业务源码、计划或 execution；execution 中“无计划内未完成项”的判断不能作为最终验收结论。

- REVIEW_STATUS：NEEDS_FIX
- IMPLEMENTATION_STATUS：PARTIAL
- TEST_STATUS：PARTIAL
- RELEASE_STATUS：NOT_VERIFIED（本轮未重新执行发布门禁）
- OPEN_ACTIONABLE：7
- 外部阻塞：无；性能探针：不适用，本轮未修改运行路径。

## 计划验收矩阵

| 阶段 | 状态 | 已有实现 | 剩余问题 |
| --- | --- | --- | --- |
| P1 公共契约 | PARTIAL | PrintTitleRows、PageHeader、上下文、快照及隐藏 SPI 已存在 | marker 占用冲突、源区域宽度、同 Sheet 多列表限制 |
| P2 Provider/导入 | PARTIAL | 两个 Provider 已接入复制、行规划、导入跳过和打印标题 | 模板打印互斥、首页 marker、结构化错误坐标 |
| P3 验证与交付 | PARTIAL | 已有双 TFM 测试、消费者、API 审批、文档和执行证据 | 补齐负向/资源用例，修正文档和完成结论 |

## 证据及复用边界

复用 execution 记录的 Release 构建、双 TFM Provider/合同测试、API 和包消费者证据，不重复运行全量测试。它们说明既有用例通过，不证明本报告中新发现的边界已满足。

本轮审查了当前源码、Git 差异及测试，并运行隔离 net8.0 包消费者（artifacts/review-probe），使用上一实施阶段生成的本地 NuGet 包。以下为 XLSX 的实际输出；HSSF 和 net6.0 对新增场景尚未运行，修复时应纳入回归。

```text
NPOI: missing first marker accepted=True, rows=3
NPOI: narrow source accepted; original C2=Qty, repeated C6=
NPOI: physical header plus template print titles accepted; print titles=CellRangeAddress [1:2]
NPOI: fixed-cell marker collision accepted; A1=HEAD
ClosedXML: missing first marker accepted=True, rows=3
ClosedXML: narrow source accepted; original C2=Qty, repeated C6=
ClosedXML: physical header plus template print titles accepted; print titles=CellRangeAddress [1:2]
ClosedXML: fixed-cell marker collision accepted; A1=HEAD
Non-overlapping ordinary region rejected: 实际重复页眉所在工作表只能包含一个列表区域。
```

## FIX-001：marker 覆盖固定业务字段

- 严重程度：HIGH；处理要求：MUST_FIX；状态：OPEN_ACTIONABLE。
- 问题：固定绑定与 PageHeader marker 可以占用同一单元格，导出静默覆盖业务值。
- 证据：布局绑定 A1 为 Title=PURCHASE，页眉 marker 也指定 A1；两个 Provider 最终 A1 均为 HEAD。ExcelEntityPageHeaderBuilder 仅校验其自身 Cell 配置，未与外层固定绑定核对；Provider 写入 marker 时覆盖先前固定值。
- 影响：生成单据丢失业务字段，且不报冲突。
- 修复目标/要求：预检保留 marker 专用位置，检查普通及命名固定绑定、启用的列表表头；保留非 marker 源单元格的合法 Fluent 覆盖。
- 验证方式：两个 Provider、NPOI 两种格式均添加冲突失败用例，确认写入前失败、原目标完整。

## FIX-002：源页眉可遗漏列表表头列

- 严重程度：HIGH；处理要求：MUST_FIX；状态：OPEN_ACTIONABLE。
- 问题：预检只约束源区域末行与明细相邻，未保证覆盖完整列表表头。
- 证据：列表 B2:C2，源 A1:B2 被接受；两个 Provider 首页 C2=Qty，第二页 C6 为空。NpoiEntityExportExecutor.cs:565、ClosedXmlEntityLayoutExecutor.cs:689 的 ValidatePageHeader 未核对映射后的完整表头列范围。
- 影响：后续页缺失标题，违反源区域包含列表表头的约定。
- 修复目标/要求：在物理列规划完成后校验启用表头的完整覆盖，包含动态列、计算列、显式物理列；无表头模式遵循其独立约定。
- 验证方式：窄源区域明确失败，完整源区域复制全部列；覆盖 XLS/XLSX 和动态列。

## FIX-003：模板打印标题与实际页眉未互斥

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：互斥只检查 Fluent PrintTitleRows，遗漏模板既有打印标题。
- 证据：模板 S.RepeatingRows=1:2，使用实际 PageHeader 后两个 Provider 输出仍同时保留打印标题和实际页眉。ExcelEntityLayoutBuilder.cs:453 只查询布局字典。
- 影响：同 Sheet 打印时可能出现双份页眉。
- 修复目标/要求：加载模板后预检同 Sheet 的重复打印标题与 PageHeader 冲突；不改动其他工作表无关打印设置。
- 验证方式：两个 Provider 拒绝冲突模板；其他 Sheet 的打印设置正常保留。

## FIX-004：首页 marker 未校验

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：导入只扫描首个明细行及后续位置，未检查首页源区域 marker。
- 证据：将导出后的 A1 从 HEAD 改为 BROKEN，两个 Provider 仍成功导入全部 3 行。ExcelEntitySubtotalImportScanner.cs:34 从 firstRow 开始扫描，两个调用方传入首个明细行。
- 影响：损坏的首页标识未被识别，不符合 marker 缺失应失败的约定。
- 修复目标/要求：在正确解析源区域及命名起点后校验首页 marker；继续校验后续页缺失、重复、错位和非法高度。
- 验证方式：空/单/多页分别验证首页缺失和合法标记；补重复 marker 场景，并保持原始数据坐标。

## FIX-005：marker 异常缺少结构化定位

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：错误只提供包含列表起点的消息，没有实际错误的 SheetName、RowIndex、ColumnIndex。
- 证据：ExcelEntitySubtotalImportScanner.cs:119 的 InvalidMarker 只传 message 和 stage；BingOfficesConfigurationException 现有构造未传位置字段。此项为源码调用链确认，非本轮独立运行断言。
- 影响：应用无法按结构化字段定位损坏页眉，计划要求的原始坐标验收未满足。
- 修复目标/要求：使用兼容的错误表达补充实际位置；若需新增公共构造，补成员审批，不删除旧签名。错误消息也应明确涉及页眉而非仅小计。
- 验证方式：断言错误代码、阶段、工作表和原始行列；分页、小计跳过后的坐标不得变成临时投影坐标。

## FIX-006：禁止同 Sheet 普通列表属于额外限制

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：计划只限制每个 Sheet 至多一个启用实际页眉的分页列表，实现却禁止该 Sheet 存在任何第二个列表。
- 证据：ExcelEntityLayoutBuilder.cs:455-457 对同 Sheet 全部列表计数；B2:C20 的 PageHeader 列表与 J2:K20 的普通列表也构建失败。
- 影响：不能表达计划允许的无重叠并列明细。
- 修复目标/要求：仅限制实际 PageHeader 区域数量，同时按真实占用校验所有普通区域及复制位置冲突；同步纠正文档中的额外限制。
- 验证方式：无重叠普通列表成功往返；两个实际页眉列表失败；复制区域与普通列表实际重叠仍在输出前失败。

## FIX-007：补齐负向/资源证据并修正交付记录

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：现有测试和完成记录未覆盖上述边界，部分计划要求的资源断言仍缺失。这是验收证据缺口，不据此断言所有相关运行行为错误。
- 证据：重复页眉测试缺重复 marker；现有 marker 失败断言未核对结构化坐标。ClosedXML 缺与 NPOI 对等的 marker/合并冲突、Fluent/模板合并冲突和值工厂异常测试；PageHeader 失败测试缺直接临时文件清理断言。execution 仍写“无计划内未完成项”和“无功能偏差”。
- 修复目标/要求：补职责级回归及资源失败矩阵；更新生产符号到测试方法映射、能力文档和 execution，明确新增用例结果。
- 验证方式：先运行新增职责测试，再受影响 Provider 双 TFM/合同/消费者；根据最终影响执行 API、Release 常规门禁，沿用 Category!=Large，不重跑无关性能矩阵。

## 建议推进顺序

1. FIX-001、FIX-002、FIX-003：先完善写入前预检，防止输出内容错误。
2. FIX-004、FIX-005：完善导入标识检查和结构化定位。
3. FIX-006：放开计划允许的多列表组合并验证真实冲突。
4. FIX-007：贯穿各修复补测试，最后统一更新文档、追溯和验收记录。

透视表、公式计算引擎、绘图对象复制与物理打印分页控制不属于本计划欠账。
