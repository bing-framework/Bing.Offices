<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ENTITY-INDEPENDENT-PAGE-ANCHORS-20261005-001
AI_REVIEWED_AT: 2026-10-05T02:16:02.453771+00:00

# 当前实现与计划差距及 TODO / 第三次进度核查

## 结论与统计口径

当前后续路线图有 **6 组功能尚未实现（F-002～007）**、**1 组功能已实现但交付未收口（F-001）**，另有 **2 组环境验收**。设计与验收 TODO 不重复计入功能数；各项规模不同，不据此换算整体完成百分比。

原 Entity 重复页眉计划已有 review-round2-final.md 的独立 PASS。打印标题、实际重复页眉、导入识别，以及透视来源合同和本机 Excel 生命周期，不再列为未实现。属性固定单元格、动态分组、Footer、字典模板、名称列表与 Entity 原生校验、保护、两级级联、Workbook 分批扩展、分组分页及分页小计锚点也已有后续实现。

本报告是当前进度核查，不替代 F-001 最终独立审查。仅新增此报告，保留源码、测试、计划、execution、API 基线、工作流状态及历史失败记录；未重新运行 dotnet 或 Office。Luna 负责只读汇总最新验证产物，主代理核对计划与真实调用链。

## F-001 对照计划

| 阶段 | 当前情况 | 剩余事项 |
| --- | --- | --- |
| P1 模型、构建与快照 | 已实现；模型双 TFM 各 9/9 | 纳入最终独立复核 |
| P2 共享物理行与导入扫描 | 每页 Footer、空页/末页、marker、高度、顺序、原始坐标已接通；新增动态/计算列组合通过 | 纳入最终独立复核 |
| P3 双 Provider、名称与模板 | NPOI XLS/XLSX、ClosedXML XLSX 已消费名称策略；外部旧名称、千页及长度边界组合通过 | 纳入最终独立复核 |
| P4 业务、资源与消费者 | 主体职责和新增组合均有真实成功产物；最终隔离包消费者四次运行成功 | 无需沿用上次“夹具待修”结论 |
| P5 API、文档与交付 | 最新 Release、常规门禁、打包、消费者及临时候选 API 比较均已通过 | 仓库候选身份、执行记录和最终独立审查仍待收口 |

### 相比第二次检查已完成

- [x] FIX-011：两个 Provider 的 B3 相对 Footer 坐标、实际聚合配置和读取流夹具已修正。
- [x] NPOI 外部名称夹具具有真实 HSSF 关联；XSSF 使用公开关系类型生成、序列化并读回外部关联后设置名称。最终夹具没有使用内部 GetCTName 或反射绕过可见性。
- [x] 实际组合测试定位并修复 NPOI 原生规则中的空属性引用：计算列不回写，动态列以 Key 作为字段标识。
- [x] 新增组合双 TFM 共 30/30：NPOI 每个 TFM 10/10，ClosedXML 每个 TFM 5/5；四次命令退出 0。包含完整工作簿、名称、原始坐标和失败行 Unique 状态隔离断言。
- [x] 当前候选常规门禁、Release 构建、最终隔离打包及包消费者已有成功产物。
- [x] 最新 API 临时候选比较退出 0；双 TFM 八个程序集公共成员与已批准 F-001 基线一致，无额外新增或删除。

原生校验失败按既有合同跳过整行；对应 Unique 测试证明失败行不占用动态值，不单独证明配置校验路径的 RollbackRow 调用。配置 Unique 失败及原始坐标由另一工作簿副本直接验证，追溯表已说明该区别。

### 最新验证证据

- artifacts/regular-results-final.json：最终选取 27 条命令全部退出 0，对应 27 份 TRX 共 4,805/4,805，通过、失败 0、跳过 0。目录内另有旧 Docs 18/19 的失败 TRX，保留但不并入最终计数。
- artifacts/composition-results.json：四次双 TFM 组合命令退出 0，共 30/30，不将此前失败候选并入成功计数。
- artifacts/final-candidate：Release 构建成功；9 个包及四个消费者运行成功，消费者 restore/build/run 共 11 条命令退出 0；api-compare.command.json 退出 0。
- final-candidate/api-member-diff.json：双 TFM、八个程序集的 added/removed 全为空。这是在 F-001 已批准新增成员之后的比较；F-001 原始审批差异仍为每 TFM 12 条新增、0 删除。
- final-candidate/encoding-changed-files.json：275 个文本文件中 143 个改动 C# 严格 UTF-8/BOM/LF/EOF 均通过；三项非 C# 例外为旧 BATCH-RELATIONS 的 plan.md、symbol-test-map.md 末尾换行及生成的 EntityProviderContracts.xml 行尾，不在本轮检查中转换。

## 当前阶段收口 TODO

- [ ] T-001：将已经验证的最终候选身份纳入仓库 API 基线，保持公共成员不变，保留修复前候选和实际差异。
- [ ] T-002：同步 execution 的阶段矩阵、FIX 终态及最新验证证据；已有符号追溯需与最终候选共同核对。
- [ ] T-003：完成并落盘最终独立审查，处理其真实问题后结束本地任务状态。

### FIX-013：仓库基线身份尚未对齐最终修复候选

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：NPOI 内部修复后的最终候选已通过 API 比较，但 build/api-snapshot-baseline.json 仍保留较早的候选身份。
- 证据：artifacts/final-candidate/api-compare.command.json 退出 0，比较对象是 final-candidate/api-approved-candidate.json。仓库基线与该候选的 assemblies 完全一致，差异仅为 approvedAt、candidateIdentity。
- 影响：公共签名兼容证据成立，但仓库基线未绑定当前已验证源码、程序集与包候选。
- 修复目标：使用已通过比较的审批内候选更新身份；不得新增未批准成员或覆盖差异掩盖签名变化。
- 验证方式：仓库基线与已验证候选完全一致，双 TFM 成员差异仍为空；候选或审批内容改变时重新比较。

### FIX-014：交付记录与最终独立审查尚未闭环

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：execution 仍为 IN_PROGRESS，部分段落写“组合夹具修复中”“常规门禁尚未执行”；最终成功产物已存在，未找到 review-round1-final.md。
- 证据：execution 的阶段表及 FIX-012 段落与最新结果不一致；当前 runtime 仍 active / review-fix / implementing。
- 影响：会重复安排已完成工作，且无法据任务记录宣布 F-001 完整交付。
- 修复目标：保留历史失败时点，追加并明确最终结果；落盘独立最终结论，通过后完成本地收口。
- 验证方式：执行记录、候选、原始命令与 TRX、追溯、最终独立报告及本地状态一致。

## 后续功能 TODO：六组尚未实现

- [ ] F-002 / 高级透视：多行字段、列字段、多值字段、筛选、Count/Average/Min/Max、排序与小计。当前 ExcelPivotTableDefinition 只有 RowField/ValueField；ClosedXmlReportWriter 仍固定单行字段 Sum。
- [ ] F-003 / NPOI XLSX 完整透视缓存：缓存字段、记录、OOXML 关系及真实客户端生命周期。NpoiReportWriter:23 仍拒绝透视请求；XLS/HSSF 支持边界须另行确定。
- [ ] F-004 / 已有透视模板保留与更新：来源变更、缓存刷新、名称冲突和未知设置保真。ClosedXmlTemplatePreflight:43 仍拒绝透视部件；ReportWriter 仍要求新建目标 Sheet。
- [ ] F-005 / ExcelDataReader 前向失败工作簿：错误行暂存、标注、产物提交、取消与清理。ExcelDataReaderExcelImporter:2278～2280 仍拒绝非 None 的 FailureOptions。
- [ ] F-006 / 端到端低内存 Entity 导入：固定字段、多区域、动态列、关系及页眉/Footer 的前向读取。NpoiEntityBatchImporter:47～48、ClosedXmlEntityBatchImporter:48～49 仍先完整 ImportEntity，再 Deliver；当前分批回调没有消除 DOM 和实体完整物化。
- [ ] F-007 / Entity 失败工作簿：固定字段、多区域标注、原始坐标、模板保真与失败产物生命周期。ExcelEntityImportOptions 只有资源、失败继续策略、校验模式与 Unsupported 策略，没有失败产物配置。

## 欠缺考虑或尚缺证据的 TODO

以下是需随功能设计落实的合同、待补验收或可选扩展，不全部属于原批准计划漏项。

- [ ] D-001 / 能力发现与诊断：区分透视创建、保留、更新、刷新、计算和渲染；错误携带透视名称、来源/目标与实际区域。当前 PivotTables 只表示新建。
- [ ] D-002 / 聚合语义：先明确金额精度与舍入、空值、混合数值类型、Count 计数规则、排序稳定性、多值字段标题冲突；再实现高级透视。
- [ ] D-003 / 输出资源预算：统一考虑来源面积、字段宽度、不同缓存数量、透视展开、复制页数及页眉/签字区单元格量。现有有界资源探针尚不能替代通用输出预算。
- [ ] D-004 / 分阶段取消：补来源预检、缓存构建、页眉复制和序列化的定向证据；来源枚举阶段成功不能代表所有阶段的响应时延。
- [ ] D-005 / 组合隔离：保护、隐藏 Sheet、字典/级联、表格、多透视、模板和复杂单据，以及同源缓存复用后的隔离；当前已测组合保留，按新增能力扩展合同。
- [ ] D-006 / 真实打印：固定纸张、缩放和字体，核对行高、组/页小计、页眉、签字区与最终 Footer。逻辑 PageBreak 不保证实际纸页边界相同。
- [ ] D-007 / 公式及绘图复制：相对和跨 Sheet 引用重定位、图片/图表复制另行设计。当前只保留公式文本，未承诺重定位、计算或绘图复制。
- [ ] D-008 / 分批事务：明确已交付或持久化批次遇后续失败/取消的结果，重试与幂等，跨 Sheet 关系状态，错误暂存内存/磁盘预算，以及失败产物提交顺序。
- [ ] D-009 / 客户端探针隔离：外部墙钟硬超时、专用环境和自身 Excel 实例清理。现有 30 秒轮询不能打断阻塞的同步 COM 调用。
- [ ] D-010 / 可选业务扩展：递归多级级联、跨行父级引用、无客户端静态汇总，按业务证据独立立项，不计入两级级联或透视 MVP 缺陷。

## 环境验收 TODO

- [ ] E-001：WPS、LibreOffice 分别验证首次显示、来源修改后刷新、保存重开。现有本机 Excel 成功证据不替代这两种客户端。
- [ ] E-002：有效 Aspose.Cells 许可证、固定字体和目标部署环境下的公式重算、PDF、逐页图片与格式保真。无许可证拒绝测试不等于商业成功验收。

## 建议推进顺序

T-001 → T-002 → T-003，先完整交付 F-001；之后 F-002 → F-003/F-004 → F-005 → F-006/F-007。聚合语义、能力细分和输出预算随透视阶段，事务与暂存预算随导入阶段实施；环境验收独立推进。无需重复安排上次已解决的夹具和全量门禁。

本轮 git diff --check 退出 0，存量 ProfileFixtures.xml 的 CRLF 提示保留。累计已跟踪差异为 66 文件、7,440 行新增、595 行删除，不含未跟踪文件，也不是本轮修改量。未提交、推送、发布、修改版本或发送外部通知。
