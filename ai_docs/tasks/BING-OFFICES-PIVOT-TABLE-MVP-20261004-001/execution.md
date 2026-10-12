<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-PIVOT-TABLE-MVP-20261004-001
AI_EXECUTION_FINISHED_AT: 2026-10-04T18:17:10Z

# 实施执行报告

## 执行结论

Workbook 导出新增有界来源区域、单行字段和单求和字段的透视表定义。ClosedXML XLSX 输出真实透视表、缓存字段及数据记录；NPOI XLS/XLSX、MiniExcel、SpreadCheetah 在写入前明确拒绝该请求。现有请求和构造签名保持兼容。

## 计划执行情况

- 公共定义、Fluent 入口、请求快照、Provider 能力位及预检完成。
- ClosedXML XLSX 输出完成；NPOI 两种格式的透视缓存记录经真实文件验证不完整，因此本阶段不声明支持。
- 文档示例、Provider 矩阵、双 TFM 包消费者、成员审批和符号测试映射完成。
- 公式计算、真实打印渲染和更复杂的透视布局仍为独立后续设计，不属于本阶段。

## 初次实施验证结果（历史）

- `Bing.Offices.sln` Release 构建通过；仅有既存的 `Microsoft.Bcl.Memory` net6.0 支持提示。
- `PivotTableContractTest` 在 net6.0、net8.0 各 10 项通过；包含完整 XLSX 结构、XLS/HSSF 与 XLSX/XSSF 拒绝、配置失败、模板冲突、取消和目标流保全。
- `Category!=Large` 常规全量测试最终全部通过。首次运行发现 API 治理的两项显式类型清单未登记新类型；补充后 `PublicApiContractTest` 双 TFM 各 11 项通过，随后重新运行常规全量测试并全部通过。
- net6.0、net8.0 本地 NuGet 包消费者均输出 `package-consumer-ok`。
- 双 TFM API 比较各新增 14 个成员、删除 0 个；更新批准基线后，包、源码清单与成员快照比较通过。
- `git diff --check` 没有空白错误；本阶段 22 个文本文件通过严格 UTF-8、BOM、LF 与末尾换行检查。Git 另提示既有 `Bing.Offices.ProfileFixtures.xml` 的工作区 CRLF 状态，本阶段未修改该文件。

## 兼容性与风险

未修改版本号、未提交、未推送或发布。模板中已有透视表的保留仍不属于支持契约；Excel 客户端打开后的显示与刷新行为由客户端决定。工作区中的其他未提交任务改动保持原样。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX；Fix Scope：recommended。
- 原始审查：`review.md`（FIX-001～004）；补充审查：`review-progress-20261005.md`（FIX-005～006）。两份审查文件保持原样。
- 延续已经注册的 Review Round 1，处理六项 SHOULD_FIX；没有新增公共成员或修改版本号。

#### FIX-001：固定业务来源名单

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 根因：业务来源与已新增透视目标共享一个可变集合。
- 修改：`ExcelWorkbookExportBuilder.Build` 固定业务来源集合，目标占用独立累计。
- 验证：`PivotSource_ShouldRejectRegardlessOfDefinitionOrder`、`MultiplePivots_ShouldPreserveBusinessSourceAndAllDefinitions` 覆盖顺序、共享来源及完整来源数据；上一轮定向测试双 TFM 通过。

#### FIX-002：工作表名称预检

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 根因：非法目标名称延迟到引擎创建 Sheet，产生不同的同步/异步异常路径。
- 修改：增加内部 `ExcelSheetNameValidation`，透视定义及既有实体布局校验复用；拒绝非法字符、长度、首尾单引号及引擎明确拒绝的控制字符。
- 验证：`InvalidTargetName_ShouldFailBeforeEnumeratingBusinessRows` 覆盖十一种非法名称且不枚举业务数据；合法中文及大小写兼容由多透视成功用例覆盖。

#### FIX-003：透视展开边界

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 根因：只验证目标起点，没有预算实际结果宽度和行数。
- 修改：`ClosedXmlReportWriter` 使用 long 算术检查两列、表头、实际分组行及总计；显式保持行列总计设置。
- 验证：`PivotFootprint_ShouldRejectOverflowWithoutChangingDestination`、`PivotFootprint_ShouldAllowExactBoundaryWithDuplicateAndBlankGroups` 覆盖末行/列、整数溢出、重复/空键、精确边界与目标保护。不声称已验证 Excel 客户端实际打印或刷新展示。

#### FIX-004：透视职责及资源矩阵

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 修改：增强 `PivotTableContractTest`，新增 `PivotTableResourcesTest`；包消费者从仅检查部件存在改为验证完整来源、字典分配及透视缓存值。
- 验证：同步/异步完整缓存一致性、OOXML 关系、真实文件成功和配置失败、写入期间取消、序列化完成后取消、提交失败、旧目标完整、临时文件清理、调用方流保持开放及异常观察次数。
- 测试过程：初次新增资源用例错误使用 `rowField` 元素名，修正为真实输出的 `field`；修正后双 TFM 各 43/43 通过。保留先前失败为测试断言修正，不归因于生产行为。

#### FIX-005：XML 注释格式

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 根因：上一轮三行 summary 转换脚本使用可跨行空白匹配，插入裸空行，切断 XML 文档块。
- 修改：只移除四个受影响文件文档行之间的裸空行；保持业务代码、签名与 BOM/LF 不变。
- 验证：模型、Workbook 构建器、ClosedXML 报表写入器与合同测试的 XML 文档块分别为 12、18、12、25 个，全部独立解析通过；后续生产编译不再出现 CS1570。

#### FIX-006：辅助 Sheet 名称预留

- 严重程度：MEDIUM；处理要求：SHOULD_FIX。
- 根因：自动字典只检查当前工作簿 Sheet，未考虑稍后创建的透视目标。
- 修改：`ClosedXmlExcelExporter` 把所有透视目标名称传入 `ClosedXmlLookupListWriter`、`ClosedXmlCascadeListWriter` 的内部预留名称参数；两个分配器按大小写不敏感规则跳过已有及预留名称。Entity 原有调用继续使用空预留名单。
- 验证：新增 `PivotTableLookupCompositionTest` 覆盖普通/级联字典、多个共源透视、大小写、模板已有辅助 Sheet、完整文本/隐藏状态、名称和校验公式、模板样式及全部缓存关系。双 TFM 包消费者额外覆盖顺延名称与完整数据。

#### FIX-007：Entity 名称校验入口直测

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 补充审查：`review-progress-20261005-round2.md`，保留原文不修改。
- 新增 `EntitySheetNameValidationTest`，通过公开布局入口分别验证固定单元格、列表区域和合并区域。十二种非法名称涵盖全部禁止字符、32 字符、首尾单引号、NUL 与 U+0003；合法中文与 31 字符名称核对完整关键布局描述。
- 六个测试方法共 42 个案例，net6.0、net8.0 各 42/42。随后最终 Provider 合同项目双 TFM 全套各 391/391，包含上述案例及现有透视测试。
- 补齐测试模型构造与数据属性中文 XML 注释，维护共享名称校验、实体构建入口到六个方法的符号映射；没有改变生产行为或公共 API。

#### FIX-008：执行记录和最终证据收口

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 原初次验证保留在历史段落，最终候选验证单独列出，不把后续单项结果拼成另一次全量运行。
- 旧 PAGE-HEADER-BOUNDARIES 记录中的 DESIGN-003/004/005 已分别由 METADATA/RESOURCES/CONTRACTS 完成；历史记录保持原样，最新总清单以实际后续阶段为准。
- 本轮源码范围仅补 Entity 名称测试；原 FIX-001～006 的生产修复及包身份已经完成验证。需要由后续独立 Reviewer 判断最终 PASS，本 Executor 不改写原 review。

### 最终门禁

- 透视职责测试：net6.0/net8.0 各 47/47；TRX 位于 `artifacts/review-fix-tests`。
- 新增 Entity 名称直测：双 TFM 各 42/42；TRX 位于 `artifacts/entity-name-direct-tests`。注释及显式默认构造完善后，合同项目 Release 编译与全套回归双 TFM 各 391/391，见 `artifacts/contracts-final-tests`。
- 解决方案 `Category!=Large` 的一次运行通过：27 份 TRX、14 个测试项目，4,289/4,289，0 失败、0 跳过，命令最终退出码 0；见 `artifacts/review-fix-regular-tests`。最后仅补测试，随后运行上述合同项目回归，没有重复整个解决方案测试或 Large 矩阵。
- Review 修复候选 Release 构建通过，0 错误、4 条既有 Microsoft.Bcl.Memory net6.0 提示。最终测试完善后的增量 Release 构建也通过，0 错误、4 条相同既有提示（用时 1 分 11 秒）。
- 八个生产项目已串行打包到 `artifacts/review-fix-packages`，版本仍为 2.0.0。net6.0（.NET 6.0.36）、net8.0（.NET 8.0.31）消费者从各自隔离缓存运行，均输出 `package-consumer-ok`，验证完整来源、辅助 Sheet 顺延、名称下拉和透视缓存。
- 双 TFM 成员逐行/哈希/数量比较与审批基线一致，新增/删除为 0；先保存 `artifacts/review-fix-api/api-member-diff.json` 空差异，再仅刷新源码、程序集、包身份，正式比较退出码 0。最终 `artifacts/review-fix-api-verified/api-diff.json` 为空差异。后续测试/任务文档不属于生产源码身份范围。
- 过程中曾在新组合测试尚未完成时构建，缺少 Task using 导致编译失败；补齐后构建与所有测试通过。消费者首次错误假设透视缓存固定 ZIP 目录而失败，改为按真实部件名称查找，完整断言保留；修正后双 TFM 均通过。此前 XML 文档及 rowField 断言失败也保留了对应原因和修正记录。
- 本轮 14 个相关 C# 文件按严格 UTF-8 BOM、LF 和单个末尾换行复核；任务文档/API JSON 使用无 BOM UTF-8 与 LF。`git diff --check` 没有空白错误；既有 ProfileFixtures XML CRLF 提示保留，不做全仓库转换。

### Round 1 汇总

- MUST_FIX：0；SHOULD_FIX：FIX-001～008 共八项完成；PARTIAL/BLOCKED/FAILED：0。
- 后续源码能力与商业部署验收仍未完成，本任务修复收口不代表整个长期目标完成。
- 实际公式探针确认透视源中的新公式会触发引擎计算；D-001 已落入下一阶段 `BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001/plan.md`，由该阶段修正合同与行为，不宣称本阶段已经解决。
- 未修改版本号、未提交、推送、发布或发送外部通知。下一步为独立复核及继续来源数据合同阶段。
