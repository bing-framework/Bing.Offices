<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001
AI_REVIEWED_AT: 2026-10-05T12:14:47Z

# 当前实现与计划差距核查

## 结论与统计口径

原 Entity 重复页眉与打印标题计划已完成，最终独立 Review 为 PASS；后续 F-001 独立逐页页眉、Footer/签字区和命名锚点也已独立 PASS。不能把后续透视、批次和环境扩展反记为原页眉计划未实现。

当前后续路线仍有 **6 组未完整交付：F-002/F-003 两组主链已经实现但完整验收未通过；F-004～007 四组仍有目标能力缺口**。另有 **2 组环境验收**。测试、设计和文档 TODO 不重复计为功能，不按不同规模的任务数估算完成百分比。

本轮只读核查生产、测试、计划和实际结果，仅新增本报告；没有修改业务代码、测试、plan、execution、历史 Review、API 基线或运行状态，没有运行 dotnet、pack 或 Office。主代理参与过之前实现，本报告是进度核查，不替代最终独立验收。Luna 独立只读确认 F-005～007；其原 F4 测试任务已停止，未创建测试文件。

## 本次相对上次新增的完成证据

| 项目 | 实际结果 | 范围限制 |
| --- | --- | --- |
| F3 Release 与常规矩阵 | regular-results-01.json 共 28 条命令全部退出 0：1 个 solution Release、27 个项目/TFM 常规单元；TRX 合计 5113/5113，失败及跳过均为 0 | 对应 F3 验证候选；之后新增 F4 生产代码不能借用这批结果验收 |
| NPOI Provider 双 TFM | net6.0/net8.0 各 1112/1112；新增组合 8 项、业务组合 2 项在两 TFM 均 Passed | 真实包结构与来源合同，不等于高级客户端数值正确 |
| ClosedXML、共享合同及 Docs | ClosedXML 各 424/424；ProviderContract 各 511/511；Docs 19/19 | 对应上述 F3 候选；之后文档/生产修改须按影响重验 |
| NPOI 基础透视真实 Excel | net6.0/net8.0、同步/异步四次均 Passed：首次完整矩阵 40/20/60，编辑来源刷新后 45/20/65，保存重开保持一致 | Excel 16.0 Build 20430；不外推高级、WPS 或 LibreOffice |
| 基础客户端探针所有权 | 12 条工具测试通过；四次实际 Excel 的父进程句柄所有权、正常打开、自身进程退出及清理均有 JSON 终态 | 工具测试和客户端数值验收分别记录 |
| 新包与 API | F3 九个 pack、十条消费者 restore/build/run 全部退出 0；普通/第三方、双 TFM 四次实际运行通过；双 TFM 零成员差异及候选身份比较通过 | 只覆盖 F3 候选；F4 新源码须重新绑定身份 |
| F4 起步 | 完整 plan 已落盘；共享 OPC 部件发现器接入默认预检，ClosedXML 双 TFM Release build 退出 0 | 未完成完整依赖图，无保留/更新公共入口，无新能力位，无 P1 直接测试 |

上次报告里的 net6 新组合、F3 常规矩阵、基础客户端、探针 README 参数说明不再作为待从零完成事项。F3 追溯表第 43 行仍称 net6 与全量正在运行，与第 47 行新终态矛盾；四次基础客户端的当前摘要和追溯仍需同步。

## 功能 TODO

- [ ] **F-002 / P0：高级透视完整数值合同。** ClosedXML 的多维轴、五聚合、筛选、排序、小计及完整缓存已实现。NPOI 也复现同一文本 Count 问题：真实 Excel V9 预期 3、实际为空，总 Count 预期 5、实际为空；显式刷新及原生重建仍为空。其它四聚合本次总值为 150、30、10、50。定位根因并通过首开、来源编辑刷新和保存重开，不用数字替代 Count、不裁剪缓存、不写静态汇总，也不据现有证据直接认定为 Provider 独有 Bug。
- [ ] **F-003 / P1：NPOI XLSX 透视最终收口。** 生产主链、双 TFM 职责/组合/业务、本地门禁、包/API及基础客户端均已有证据。剩余高级完整客户端矩阵、列轴/多筛选等实际显示验收，以及最终独立 Review。XLS/HSSF 透视提前拒绝是批准边界，不列为遗漏。
- [ ] **F-004 / P1：已有透视模板保留与更新。** 当前只开始 P1 发现部件。尚缺 workbook cacheId、Sheet 所属、缓存消费者和关系 ID 的不可变清单；新公共模式及快照；保留原生部件；精确更新来源、轴、度量和筛选；共享缓存局部更新时拆分；完整资源/客户端/交付验收。普通模板中新建透视不能代替该能力。
- [ ] **F-005 / P1：ExcelDataReader 分批失败工作簿。** Workbook batch 仍在读取前拒绝非 None 的 FailureOptions，单表批次也没有失败产物选项。补错误行暂存、原始坐标、可写输出格式、公开提交服务、取消/清理和最终结果。现有 DataReader 暂存后读取不等于零磁盘直接流式。
- [ ] **F-006 / P1：解析中交付的 Entity 分批模式。** 两个 Provider 仍先完整 ImportEntity，再 ToArray/投影回调，摘要保留完整 Result。增加兼容的低保留模式，在解析中交付并限制实体集合、错误及批次工作集；分别说明 DOM、跨批次校验状态和数据集合的资源成本。
- [ ] **F-007 / P1：Entity 失败工作簿。** 完整及分批 Entity options/result 都没有失败产物配置/结果，未接入生成链。补固定字段、多 Sheet/区域、跳过页眉/小计/Footer 后原始物理坐标的标注；覆盖同步/异步、模板、流/文件、公开提交器与资源合同。

## F4 当前五阶段差距

| 阶段 | 状态 | 尚需完成 |
| --- | --- | --- |
| P1 部件清单与依赖图 | PARTIAL | 完整 Manifest；非标准路径、共享缓存、异常/取消/所有权的两个 Provider 直接职责；NPOI 当前变更编译证据 |
| P2 公共快照与预检 | 未实现 | PreserveTemplatePivotTables、UpdatePivotTable、Request SPI、成员审批、互斥及快照、其他 Provider 拒绝 |
| P3 保留/更新输出 | 未实现 | 原生部件与未知配置保全、缓存拆分、索引重建、完整预检、取消/提交/清理 |
| P4 业务与客户端 | 未验收 | 实际旧透视模板、来源扩缩、多缓存、五聚合、首开/刷新/重开 |
| P5 交付 | 未收口 | 新候选 API/包、双 TFM、Release/常规/Docs、追溯及独立 Review |

## 欠缺考虑或需要落实的 TODO

区分尚需设计、计划已有要求待落实和证据不足，不将假设风险写成已复现 Bug。

- [ ] **D-001 / 尚需设计：总资源预算。** 除单透视 MaxExpandedCells 外，明确输入解压、DOM、完整缓存、多透视、错误集合、源暂存和失败文件的总内存/磁盘限额；已有 100K 探针不证明端到端恒定内存。
- [ ] **D-002 / 尚需设计：低保留返回值。** 摘要继续持有完整 Result 时，分批回调无法降低结果保留成本。定义是否保留实体、固定字段何时可用、区域交付顺序和摘要终态，保持旧调用默认兼容。
- [ ] **D-003 / 尚需设计：跨批次状态。** Unique、跨 Sheet 关系根集合、错误原文和失败行随输入增长；定义各自预算、超限终态及是否采用磁盘索引。
- [ ] **D-004 / 尚需设计：批次事务和重试。** 前面批次已入库后发生取消、回调异常、末尾关系失败或失败文件提交失败，如何报告部分成功；明确幂等重试、Unique 回滚、关系最终绑定和文件提交顺序。
- [x] **D-005 / F4 已完成：模板依赖安全。** 透视图、透视切片器和时间线已识别，并在更新前按功能拒绝；保留模式原样保留关联部件。详见 PTPL-003 执行记录和依赖边界合同测试。
- [x] **D-006 / F4 P1 已完成：扫描资源与取消。** `ExcelPivotTemplateParts.Find` 接收并逐元素检查取消令牌，ClosedXML 预检传入调用令牌；元数据读取按 `ExcelResourceLimits` 限制 XML 字符数和深度，ZIP 预检负责包级预算。NPOI、ClosedXML 的 `ExcelPivotTemplatePartsTest` 覆盖预取消、DTD、字符数及深度限制；见 P1 双 TFM 执行记录。
- [ ] **D-007 / 尚需设计：失败文件语义。** 分开声明输入读取格式和失败产物写入格式；原件标注与失败行副本的选择、固定字段和多区域坐标、无失败行、错误截断、取消及提交失败的产物策略。
- [ ] **D-008 / 补证据：类型及地区。** 文本 Count、空字符串/真空值、大整数和金额精度、1900/1904 日期、标签大小写/空白、地区排序需要实际客户端矩阵；已有来源类型职责不重记为未实现。
- [ ] **D-009 / 尚需明确：能力与数据权限。** 分开声明新建、模板保留、更新、刷新、计算和渲染及格式组合。完整缓存保留全部来源，筛选、隐藏 Sheet 或保护不能当作数据权限隔离。
- [ ] **D-010 / 补证据：物理打印。** 逻辑 PageBreak 不等于纸页；固定字体、纸张、缩放、行高和签字区位置后做 PDF/逐页图片验收。公式文本复制不重定位引用、绘图不复制仍是已声明边界。

## 环境 TODO

- [ ] **E-001：** WPS、LibreOffice 分别完成首开、来源编辑刷新、保存重开的完整数值与布局验收，不从 Excel 结果推定。
- [ ] **E-002：** 有效 Aspose.Cells 许可证及固定字体/部署环境下验证目标公式重算、PDF、逐页图片和格式保真；无许可证拒绝测试不代替成功验证。

多级/跨行级联、公式引用重定位、图片/图表复制、按聚合值排序、自定义聚合和完整公式引擎属于另行设计的可选扩展，不混入六组当前必做。

## 建议顺序

1. 保留 F3 已通过证据，补当前摘要/追溯矛盾；优先定位 F-002/003 Count，成功后补高级矩阵及独立收口。
2. F4 按 P1 直接测试与完整依赖图 → P2 公共快照 → P3 保留/更新 → P4 业务客户端 → P5 交付推进。Count 未解决时继续可独立完成的模板/本地工作，但不将高级客户端项关闭。
3. 共同设计 F-005～007 的坐标、资源、失败产物与批次终态，再依 F-005 → F-006 → F-007 实施。
4. 环境具备时补 E-001/E-002，保持各环境独立结果；物理渲染结论不由逻辑布局推定。

## 可消费 Findings

### FIX-001

- 严重程度：HIGH；处理要求：MUST_FIX。
- 问题：F-002/003 高级 Count 客户端合同仍失败。
- 证据：F3 artifacts/client/npoi-net8-advanced-sync-client-01.json Status=Failed，V9 expected=3 actual=空；LastGrandValues 和 DiagnosticAfterExplicitRefresh 的 Count 为 null；DiagnosticNativeRebuiltCount=null。前序 ClosedXML 相同失败保留。
- 影响：五聚合完整数值未验收，不能将结构/基础客户端通过外推为高级完成。
- 修复目标与要求：定位根因，保持原始类型、完整来源和五聚合合同，不弱化断言或静态替代。
- 验证方式：最终候选、两个 Provider 的首开/编辑刷新/保存重开完整矩阵，随后独立 Review。

### FIX-002

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：F4 新预检实现尚无直接职责和完整验证。
- 证据：新 helper 已被两 Provider 消费；artifacts 只有 parts-build-01 的 ClosedXML 编译终态；测试目录未找到新 helper 的直接用例。Find 返回名称集合，尚无 cacheId/所属 Sheet/消费者清单。
- 影响：非标准路径、安全扫描、取消、默认拒绝及资源行为不能凭编译认定成立；也尚不能据此更新旧透视。
- 修复目标与要求：按 P1 完成直接测试与完整依赖图；补 NPOI 编译，传递资源/取消约束；不得将默认拒绝改进命名为保留/更新完成。
- 验证方式：两 Provider、双 TFM 实际模板及非标准部件路径职责，完整清单和异常/资源断言。

### FIX-003

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 问题：F3 当前摘要和追溯存在过时表述，F4 新源码的候选身份还未完整验证。
- 证据：production-test-map.md 第 43 行仍称 net6/门禁正在运行，第 47 行记录 5113/5113；四份基础客户端 Passed 已存在，但当前摘要未完整同步。F4 helper/项目链接/预检是 F3 门禁之后的生产变更。
- 影响：容易重复已完成工作，或把旧候选验证误当新代码验证。
- 修复目标与要求：同步当前摘要、矩阵和符号追溯，保留历史记录；按变更影响重新绑定 F4 API/包/门禁，不覆盖旧审批掩盖差异。
- 验证方式：原始命令、TRX、客户端 JSON 与当前文案一致，最终身份检查及独立复核。

## 源码与原始证据

- 已完成页眉：ENTITY-REPEATED-HEADERS/review-round2-final.md；独立锚点：ENTITY-INDEPENDENT-PAGE-ANCHORS/review-round1-final.md，均为最终 PASS。
- F4：src/provider-shared/ExcelPivotTemplateParts.cs:57；ClosedXmlTemplatePreflight.cs:37；NpoiReportWriter.cs:65。公共 PreserveTemplatePivotTables、UpdatePivotTable、TemplatePivotTables 当前在 src/tests/docs 中未发现。
- F5：src/Bing.Offices.ExcelDataReader/Bing/Offices/ExcelDataReader/ExcelDataReaderExcelImporter.cs:2278。
- F6：src/Bing.Offices.Npoi/Bing/Offices/Entities/NpoiEntityBatchImporter.cs:47；ClosedXML 对应文件第 48 行；src/provider-shared/ExcelEntityBatchProjection.cs:48、142；ExcelEntityBatchImportSummary.cs:31。
- F7：src/Bing.Offices.Abstractions/Bing/Offices/Entities/Import 下完整及分批 options/result 未见失败文件契约。
- F3 原始文件：artifacts/regular-results-01.json、regular-tests/*.trx、package-pack-results-01.json、package-consumers-results-01.json、pivot-api-verify-01.command.json；client/npoi-net6-basic-sync-client-01.json、npoi-net6-basic-async-client-01.json、npoi-net8-basic-sync-client-01.json、npoi-net8-basic-async-client-01.json 及各自 harness；高级失败记录与基础工具 12 项记录均保留。
- F4 原始文件：artifacts/parts-build-01.command.json，实际退出 0，仅证明当前 ClosedXML 编译。

本轮 git diff --check 退出 0，只有存量 ProfileFixtures.xml CRLF 提示；210 个已改/未跟踪 C# 文件严格 UTF-8 BOM/LF/单 EOF 检查无偏差。生产友元仍仅指向测试程序集。已跟踪累计差异为 71 文件、7778 新增、606 删除，不含未跟踪内容，不代表本轮改动量。

本报告 UTF-8 无 BOM、LF、单 EOF；未提交、推送、发布、改版本或外部通知。
