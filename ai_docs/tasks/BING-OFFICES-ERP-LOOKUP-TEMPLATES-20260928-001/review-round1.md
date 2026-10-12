<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001
AI_REVIEWED_AT: 2026-09-28T23:01:59Z

# 计划差距检查与 TODO

## 结论与范围

针对用户批准的 ERP 字典模板与导入校验计划，当前基础能力已实现，但不能按完整验收标准视为全部完成。发现 6 组实现边界问题、3 组测试/验收缺口，共 9 组 OPEN_ACTIONABLE；另有 4 个原定后续阶段，不计为本轮漏实现。分组数量不代表工作量百分比。

本轮只审查，不修复业务代码，也不修改原 plan.md/execution.md。上轮 execution.md 的 COMPLETED 是执行方结论，本报告提供复核后的差异。

## 验收矩阵

| 计划项 | 状态 | 证据与差距 |
|---|---|---|
| Workbook/Entity LookupList、ListSourceName、DataValidation、导入选项 | PASS | 公共契约及 Provider 调用链存在，原构造保留 |
| 文本快照、去重、隐藏 Sheet、名称下拉 | PASS | Model、两个 Provider 真实文件测试及 Consumer 已覆盖基本路径 |
| 名称预检与不依赖 Sheet 声明顺序 | PARTIAL | 自动字典可用；手工名称前向引用失败，二维来源未拒绝 |
| 内联/名称列表混合 | FAIL | NPOI 同范围混合时改写名称公式 |
| 名称列表导入 | PARTIAL | 局部优先、公式字典等已有实现；合法带逗号 Sheet 名被拒绝 |
| Entity 原生规则优先执行 | PARTIAL | 已接入；ClosedXML 按单元格交错执行，失败行仍产生后续转换错误 |
| 单次导入的名称与允许值集合缓存 | PARTIAL | NPOI 缓存数组但逐值匹配，ClosedXML 逐行读取名称区域 |
| 资源、格式、Provider 拒绝及业务组合测试 | PARTIAL | 共用基础测试已通过，新增 LookupList 组合路径证据不完整 |
| API、NuGet、文档 | PASS_WITH_GAPS | 前轮双 TFM +14/-0，Consumer 已接入；完整验收结论需在修复后更新 |

## 验证方法及实际输出

阅读当前 Git Diff、新增文件、调用链、测试及执行记录。复用前轮 Release/双 TFM/Consumer 证据，不重跑无关全量测试。

在 `artifacts/review-erp-lookup` 创建独立 net8.0 诊断程序，引用前轮 Consumer 输出中的库，生成 XLSX 再读回。诊断程序不修改生产源码或正式测试。构建在受限环境中因写入 DLL 权限失败，提升权限后成功运行，退出码 0。

| 探针 | NPOI | ClosedXML |
|---|---|---|
| Data Sheet 引用后声明 Dictionary Sheet 的 Allowed | ConfigurationException，名称不存在，输出 0 字节 | 同左 |
| Allowed 指向二维 `$A$1:$B$2` | 导出成功 | 导出成功 |
| A2 先内联 Yes/No，再名称 Allowed | 两条公式读回为 `&#34;Yes,No&#34;` 和 `"Yes,No"`，Allowed 丢失 | 名称 Allowed 保留 |
| Entity 第一列违反名称规则，第二列不可转 decimal | 仅 WorkbookValidation | WorkbookValidation 加 ValueConversion |
| Allowed 指向 `'Lookup,2026'!$A$1`，值匹配 | Unsupported，明细未接纳 | 同左 |

请求中的 LookupList 与手工 NamedRange 同名由 Workbook builder 正确拒绝，不列为缺陷。

## FIX-001：防止内联修正覆盖名称规则

- 严重程度：HIGH；处理要求：MUST_FIX；状态：OPEN_ACTIONABLE；优先级：P1。
- 问题：XLSX XML 修正按 type=list 和 sqref 选最后一条规则，未核对规则来源。
- 证据：`src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiSheetContentWriter.cs:100`；同范围混合探针已复现。
- 影响：名称规则被改成内联内容，允许值改变；原内联规则的引号也可能未修正。
- 修复目标：只改写对应的显式列表节点，名称公式始终保留。
- 修复要求：建立稳定的规则对应关系；如决定拒绝重叠规则，应在预检明确拒绝，不能悄悄改写。
- 验证方式：同范围的两种添加顺序、不同范围、模板已有规则，读取完整验证 XML 和工作簿规则。

## FIX-002：统一建立全部名称后再应用校验

- 严重程度：HIGH；处理要求：MUST_FIX；状态：OPEN_ACTIONABLE；优先级：P1。
- 问题：自动字典先建立，但手工名称仍在逐 Sheet 的 ReportWriter 中建立。
- 证据：NPOI `NpoiExcelExporter.cs:489,860`；ClosedXML `ClosedXmlExcelExporter.cs:562,715`。前向引用探针两个 Provider 均失败。
- 影响：同一合法请求只因 Sheet 顺序不同就导出失败，违反明确的计划约定。
- 修复目标：Workbook 级预建全部请求名称，随后应用各 Sheet 校验。
- 修复要求：保留局部作用域及冲突检查，避免 ReportWriter 二次注册；兼容模板名称。
- 验证方式：调换业务/字典 Sheet 声明顺序结果等价；覆盖 NPOI XLS/XLSX、ClosedXML XLSX。

## FIX-003：补齐名称来源形态预检

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P1。
- 问题：ApplyValidations 仅检查名称存在，没有检查是否解析到一维有界区域。
- 证据：NPOI `NpoiSheetContentWriter.cs:172`；ClosedXML `ClosedXmlSheetContentWriter.cs:44`。二维来源探针两个 Provider 均成功导出。
- 影响：可生成不符合计划契约的下拉，随后原生导入又判 Unsupported。
- 修复目标：导出前解析来源及局部优先关系，拒绝不支持形态。
- 修复要求：涵盖请求和模板名称、缺失名称、无效局部遮蔽、名称链、动态/二维/联合/外部引用；保持输出前失败。
- 验证方式：合法单格/单行/单列往返；非法来源断言结构化 Plan 错误和目标未写入。

## FIX-004：统一 Entity 行级校验顺序

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P1。
- 问题：ClosedXML 在每个单元格上交错执行原生、转换、配置；Continue 模式原生失败后继续转换后续单元格。
- 证据：`ClosedXmlEntityLayoutExecutor.cs:788,812`；探针得到额外 ValueConversion，NPOI 只报 WorkbookValidation。
- 影响：Provider 结果不一致，失败行会执行不必要的转换器，可能触发应用侧副作用。
- 修复目标：先完成当前行原生规则检查；失败时不进入该行转换和配置物化。
- 修复要求：保留 Report/Fail、Continue/StopOnFirstFailure 的原生错误收集及 Unique 回滚语义。
- 验证方式：转换器计数、混合类型错误、Unique 后续行重用，断言完整错误及两 Provider 一致性。

## FIX-005：正确解析带引号的合法 Sheet 名

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P1。
- 问题：引用解析在拆分引号前全局拒绝逗号；NPOI 还全局拒绝括号等字符。
- 证据：NPOI `NpoiWorkbookValidationPipeline.cs:393`；ClosedXML `ClosedXmlWorkbookValidationPipeline.cs:443`；合法 `'Lookup,2026'!$A$1` 在两个 Provider 均判 Unsupported。
- 影响：模板使用合法业务 Sheet 名时名称下拉无法导入。
- 修复目标：区分引号内 Sheet 名字符与引号外公式/联合区域语法。
- 修复要求：继续拒绝真正的外部、联合和动态引用；兼容转义单引号。
- 验证方式：逗号、括号、感叹号及转义引号 Sheet 名，配合真正联合/公式反例。

## FIX-006：完成允许值集合缓存

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P1。
- 问题：NPOI 缓存 `string[]`，每个值调用 `Contains(..., Ordinal)` 线性扫描；ClosedXML 每次调用重新解析并遍历名称区域。
- 证据：NPOI `NpoiWorkbookValidationPipeline.cs:120,353,493`；ClosedXML `ClosedXmlWorkbookValidationPipeline.cs:250,306`。
- 影响：行数 R、字典项 D 时匹配仍可能达到 O(R×D)，未完成计划“避免逐行扫描整份字典”的要求。
- 修复目标：按一次导入、名称与作用域缓存解析结果及允许值集合。
- 修复要求：保持 Ordinal 比较、公式单元格拒绝、失效局部不回退；扫描期间响应取消，跨工作簿隔离。
- 验证方式：用计数或稳定的定向探针证明字典只扫描一次、重复匹配不逐项扫描；不要求大型性能矩阵。

## FIX-007：补齐业务组合与格式边界测试

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P2。
- 问题：当前新增职责测试主要是 OrderStatus/Allowed 小型样例。
- 证据：两个 Provider 的 `EntityNativeValidationTest`、`LookupListExportStructureTest`，NPOI `LookupListModelTest`、`NamedListValidationTest`。
- 影响：尚无直接证据验证 LookupList 与 ValueMap/Converter、分页/小计跳过后的坐标、模板冲突组合。
- 修复目标：补采购类别/仓库/状态场景、空白填写模板、编码往返及完整错误坐标。
- 修复要求：模板已有名称冲突、辅助 Sheet 同名顺延与原内容保留、超 255 文本、格式边界；补 NPOI 联合/外部名称边界。
- 验证方式：双 TFM、NPOI XLS/XLSX、ClosedXML XLSX 的职责级真实文件断言；不以文档示例替代测试。

## FIX-008：补齐不支持 Provider 的拒绝合同

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P2。
- 问题：MiniExcel、SpreadCheetah 已增加拒绝分支，但缺少新增请求的直接测试。
- 证据：对应测试目录未发现 LookupList/NamedListValidation 专项请求覆盖。
- 影响：常规旧测试不能证明新配置不会静默丢失或提前写出部分目标。
- 修复目标：为 LookupList 及仅 ListSourceName 两类请求固定拒绝行为。
- 修复要求：验证结构化错误、同步/异步入口、目标流/已有文件保护，保持旧内联列表能力。
- 验证方式：各受影响实现独立测试，断言完整错误和输出状态。

## FIX-009：补齐新增路径资源证据并更新交付结论

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE；优先级：P2。
- 问题：LookupList 与预取消、写入中取消、提交失败的组合无专项证据；执行报告把已有公共资源测试视作本轮充分验收。
- 证据：当前新增测试及 `execution.md` 的 P3/测试结论；已有普通导出测试不带 LookupList。
- 影响：新字典循环、名称扫描及 Entity 接入的资源边界尚未直接验证。
- 修复目标：增加定向资源与缓存测试，修复后更新符号追溯和 execution 结论。
- 修复要求：旧目标完整、临时文件清理、调用方流保持开放、观察次数一致；不跑无关大规模性能矩阵。
- 验证方式：先职责测试，再受影响双 TFM 与 Provider 合同；仅在源/API变更需要时重建包及 API 证据。

## 后续独立阶段（不计入本轮差距）

1. 填写区域保护：公共保护/可编辑区域契约及 Provider 一致行为。
2. 级联字典：依赖关系、空父值、非法组合和服务端校验契约。
3. 分批导入扩展：多 Sheet、动态列、关系和失败输出的明确边界。
4. 分组分页组合与透视报表：先组合行布局，再独立设计公共透视模型。

建议先 FIX-001～005，随后 FIX-006；每个修复同步补其直接测试，最后补 FIX-007～009 的组合验收，再进入后续阶段。
