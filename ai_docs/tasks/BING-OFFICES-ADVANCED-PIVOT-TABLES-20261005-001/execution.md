<!-- AI_EXECUTION_STATUS: IN_PROGRESS -->
AI_TASK_ID: BING-OFFICES-ADVANCED-PIVOT-TABLES-20261005-001
AI_EXECUTION_STARTED_AT: 2026-10-05T02:27:35.414Z

# 实施执行报告

## 当前结论

IN_PROGRESS。F-001 已独立 PASS、no-notify 本地收口。本阶段模型及拒绝合同双 TFM 各 47/47；ClosedXML 高级导出/边界/资源/格式/组合职责双 TFM 各 38/38。当前包的真实 Excel 文本 Count 组合仍失败，高级客户端未通过。API 双 TFM 精确批准新增 37 项、删除 0 项，候选源码/包身份比较及根快照更新已完成；四次普通/第三方包消费者成功。Release 构建、27 个常规项目/TFM 单元共 4,975/4,975 通过（含 7 个未受影响单元的既有结果），文档 19/19、25 个围栏成功。默认 Excel 进程所有权入口已统一，新增默认路径的正常会话验证仍待完成，不将自动化门禁通过当作客户端数值成功。

## 阶段矩阵

| 阶段 | 状态 | 事实 |
| --- | --- | --- |
| P1 公共模型与快照 | COMPLETED | 六个新单类型文件、Definition 校验/Clone；模型及拒绝边界双 TFM 各 47/47；重复请求与失败后复用通过，API/包身份已绑定 |
| P2 ClosedXML 预检与写入 | PARTIAL | 字段、五聚合、筛选和预算已接入，自定义格式已修复；双 TFM 职责各 38/38，完整客户端 Count 数值未通过 |
| P3 业务、资源与客户端 | IN_PROGRESS | 包结构与普通模板组合成功；显式父进程所有权模式的连接前后超时和任务哨兵隔离通过；已改为默认入口，其正常会话复验及完整 Count 生命周期待完成 |
| P4 API、文档与交付 | PARTIAL | 37 项批准增量、包身份、四次消费者、Docs、Release 和常规矩阵通过；严格编码检查排除一份编译器生成 XML，最终独立复核仍待核心客户端通过后执行 |

## 实施记录

- 本地依赖 ClosedXML 0.105.1 的缓存 XML 文档及公共程序集元数据确认 RowLabels/ColumnLabels/ReportFilters/Values、SortType、AddSubtotal 和所需五种 SummaryFormula；没有在生产中使用反射或 Provider 内部 API。
- 原 RowField/ValueField 保留旧路径，只有显式 ValueFields 非空进入高级分支。来源区域的全部公式/Excel 错误拒绝合同不变。
- 高级预检按单次数据扫描检查数值度量和筛选候选，记录行列轴的类型化层级；筛选按来源实际字面值回写，避免忽略大小写匹配成功却序列化原始大小写导致客户端选项失效。
- 筛选不删除源记录或缓存；预算按完整来源的轴组合、小计、度量、总计、筛选与表头估算。高级模式采用明确的保守上界，旧模式继续原精确格式边界；显式 MaxExpandedCells 同时对旧模式有效。
- 多值、筛选、排序和自动/无小计已映射公共 Pivot API，但效果以随后真实保存/读回与客户端证据为准。
- 细粒度 AdvancedPivotTables 能力位已列入计划和审批，表示高级新建，不扩大为模板更新或客户端刷新。

## 修改范围与职责

Luna 独占 Exports/Reports 的六个新描述/枚举、ExcelPivotTableDefinition.cs 和 AdvancedPivotModelTest.cs；主代理负责 ClosedXML ReportWriter、SourceValidator、两个新的内部职责类型、AdvancedPivotExportTest.cs、能力位和任务文档。各文件单一非部分类，沿用现有目录与公开命名空间。

## 验证与遗留

初始实施时尚未运行测试；以下更新替代该历史状态，不删除原始失败产物。

| 产物 | 结果 | 说明 |
| --- | --- | --- |
| model-net8-01.command.json | 退出 1 | 测试辅助调用的第五参数位置错误，已改为命名参数 |
| model-net8-02.trx / model-net6-01.trx | 各 39/39，命令退出 0 | 旧模型兼容、深快照、36 个无效定义、失败后 Builder 状态 |
| closedxml-net8-01/02.trx | 各 2/4 | 原断言误将单选重开 SelectedValues 视为最终单选状态；诊断 XML 后使用 pageField.item 与缓存值索引验证 |
| closedxml-net8-03.trx | 6/6，命令退出 0 | 同步/异步、文本单选/部分多选、全部字段/缓存、非法筛选与预算目标保护 |
| closedxml-net8-boundary-01.trx | 13/17，命令退出 1 | 四个新模板夹具的 A1 表头被空 DTO 入口覆盖；修正夹具，不改变生产以适配错误输入 |
| probe-pack-*.command.json | 三包退出 0 | 本阶段 Abstractions、Core、ClosedXML，验证用 PackageVersion=2.0.0，不修改仓库版本 |
| probe-restore-01 / probe-build-01 | 退出 0 | 隔离缓存和仅本地源映射；真实 NuGet 引用，双 TFM 编译，保留既有 NETSDK1138 警告 |
| probe-net8-sync-01 | 退出 0 | 两行轴、两列轴、五聚合、部分多选和全部七条缓存记录；数值客户端验收仍待完成 |

客户端初次受限会话返回 COM 80070520，记录保留；随后在已授权的本机 Windows 会话运行。第二次工具调用暴露 PowerShell 大小写不敏感导致 Books 集合变量与 books COM 变量冲突，已改为 OwnedWorkbooks，并修正进程退出查询的终态退出码；第三次运行仍以实际终态为准。这些属于探针/环境问题，不写作生产聚合错误。

P3 完整类型/空值、资源失败、Provider 拒绝、高级客户端与 P4 API、消费者、Docs、常规门禁仍未完成。后续 F-003～007 和环境验收不缩小或宣告完成。

## Git 与编码

保护前序大量未提交改动，不重置工作区；新增/修改 C# 保持 UTF-8 BOM/LF，文档与 JSON 无 BOM/LF。最终严格字节检查待执行。未自动 commit、push、创建 PR、发布、修改版本或发送外部通知。

## 2026-10-05 后续实施：职责、组合、Count 最小复现与硬超时

上一进度检查为只读审查；恢复实施后没有改写历史报告。当前状态继续 IN_PROGRESS，F-003～007 与环境验收仍保留在完整 Goal 中。

- Provider 模型及 NPOI XLS/XLSX、MiniExcel、SpreadCheetah 高级请求拒绝：provider-net8-duties-01.trx、provider-net6-duties-01.trx 各 47/47。
- 导出/边界/资源职责：duties-02～05 的 5 项资源失败为测试辅助路径假设；按实际 OPC 部件关系、默认 Sum 属性及筛选主体位置修正。duties-06～08 保留组合文件编写或 API 误用的编译失败。duties-09 资源 5 项和组合 2 项失败；duties-10/11 的原 34 项全部通过，仅组合 2 项仍失败。已修正表格名称缓存来源、.rels 解析及合成 Values 轴，最终重跑待执行。未通过删除断言或改变生产规避夹具问题。
- 新增 AdvancedPivotCompositionTest 及两个独立 DTO，覆盖同源双透视不同轴/筛选/聚合/格式、共享完整缓存、重复导出和预取消失败后的请求复用；组合真实普通模板、隐藏 Sheet、字典编号冲突、名称下拉、保护和表格。
- 高级 Excel 失败记录 excel-net8-sync-10 保留：Count 的 V9 预期 3 却为空，整体 Count 预期 5 却为空；其他四聚合总计正确。30 秒被动观察、显式刷新及重新添加 Count 不改变该结果。
- count-diagnostic-02～07 创建任务副本及全新 Excel 工作簿，通过公共 COM API 原生创建对照透视，均声明 DiagnosticOnly。原生单 Count=7；Count+Sum/Average 可正确，加入 Minimum 或 Maximum 并保留列轴时文本 Count 总计为空；改为数值 Count 后为 7；仅 Count 或移除列轴后正确。更改标题、度量位置、格式与列自动小计不解决完整组合。此为 Excel 16.0 / Build 20430 的实际对照证据，不能宣告高级客户端 PASS，也不通过替换 Count 来源或静态汇总降低完整合同。
- run-advanced-excel.py 增加显式诊断脚本与仅测试用的 owner 发布后延迟入口。timeout-owned-01.harness.json 的 12 秒超时触发成立，forcedOwnedExcelCleanup=true、ownedExcelExited=true、cleanupError=null；预期子探针状态仍 Failed/Missing，外围故障合同成立不等于生命周期成功。实例创建到 owner 发布前的窗口和既有用户实例保全尚缺专门验证。
- 完成 continuation-evidence-01.json，包含真实 TRX、各诊断数值和超时终态，不用 diagnostic 的 Passed 表示聚合通过。更新透视及 Provider 文档、探针说明，明确高级验收未完成和 Count 失败。
- 初步双 TFM API 捕获正在执行；只记录候选差异，不更新仓库基线或宣告最终 API 身份通过。

资源测试修订及组合测试均保持单个非部分类单文件、中文 XML 与 C# BOM/LF；PowerShell CRLF、Python/文档 LF。没有修改生产聚合语义、依赖版本、目标版本或提交算法。没有 commit、push、PR、发布、发送外部通知或 Finalize 本阶段。

## 自定义数字格式修复与双 TFM 职责

- `probe-format-build-01` 的双 TFM CS0104 属于诊断工具 `SaveOptions` 歧义；以完整类型名修复，`probe-format-build-02` 双 TFM 构建退出 0，保留既有 NETSDK1138 警告。
- `number-format-01.xlsx.probe.json` 通过直接公开引擎 API 对照：SetFormat("0.0") 调用后 Format 立即为空，公开 Format 属性赋值在保存/重开后保留 "0.0"。诊断不是 Provider 验收，也没有反射或复制内部实现。
- Provider 保留 SetFormat 对内建格式编号的识别；仅在 NumberFormatId<0 时通过公开 Format 属性保存自定义文本。旧内建 0.00 继续序列化 numFmtId=2，自定义 0.0 和带单位格式分别保存为自定义编号。
- 新 `AdvancedPivotNumberFormatTest` 覆盖同步/异步真实工作簿、同一 Amount 的三个格式、公开内建/自定义表示、dataField 与 styles.xml 格式映射、完整来源和缓存。DTO 拆为独立 `AdvancedPivotNumberFormatRow.cs`，private 字段和参数补齐中文 XML。
- `closedxml-net8-format-composition-01/02` 的失败保留：第一版属性统一赋值改变内建编号；第二版测试误将内建编号的空 Format 属性视为缺失。最终保留原内建编号并按公开 NumberFormatId 读取内建格式，未删除自定义或 OOXML 断言。
- `closedxml-net8-duties-13.trx`、`closedxml-net6-duties-01.trx` **各 38/38，命令退出 0**。包含原导出/边界/资源 34 项、两个普通模板组合方法和两个数字格式同步/异步场景。
- `api-capture-initial` 命令退出 0；逐成员检查双 TFM 37 项新增均在批准范围，删除与非批准增量为 0。最终包身份与仓库基线仍待收口。

Count 完整客户端失败仍保留；没有替换文本 Count 来源、减少缓存或写静态汇总。当前阶段和完整后续 Goal 均保持进行中。

## 2026-10-05 后续实施：常规矩阵、API 候选身份与默认所有权入口

上一轮 `review-progress-20261005-05.md` 为只读检查，本节记录之后的真实实施；不覆盖历史失败及独立报告，不缩小 F-003～007 与环境验收范围。

### 自动化验证与 API

- 修正文档匿名 DTO 的构造约束后，`docs-continuation-01.trx` 为 19/19；25 个围栏实际编译执行，包含高级透视公开 DTO、同步/异步结构和完整缓存断言。当前文档明确结构验证不等于数值成功。
- `continuation-consumer-results-04.json` 中普通 net6/net8、第三方 net6/net8 四次真实运行退出 0。9 个当前源码 pack 命令均退出 0；API 身份需绑定的八个生产包与调用程序集一致，合同包消费者也真实执行。没有生产友元或项目引用替代包消费。
- `api-continuation-01/api-before-f002.json` 保存原 F-001 基线，`api-approved-member-diff.json` 两个 TFM 均 +37/-0，非批准新增和缺少批准成员均为 0。`api-continuation-verify-01` 身份比较成功后更新根快照；集成 `integration-net6-regular-01`、`integration-net8-regular-02` 各 50/50。
- 本节同步 API 审批说明后捕获 `api-continuation-02`。第二次验证脚本读取生成快照时未处理 UTF-8 BOM，导致候选未生成；`api-continuation-verify-02` 的 missing baseline/退出 2 作为工具失败保留。显式 UTF-8-sig 读取后，确认 API 形状、源码清单、程序集及包身份均与上次批准候选一致，只刷新审批文档身份。`api-continuation-verify-03` 退出 0、双 TFM 空差异，再更新根快照；更新前身份另存 `api-before-identity-refresh.json`。
- 身份刷新后的 API 治理直测：`api-governance-net6-02.trx`、`api-governance-net8-02.trx` 各 11/11。相同测试已包含在常规矩阵，不额外累加。
- `release-build-01.command.json`：Release 构建退出 0，0 错误；4 条存量 Microsoft.Bcl.Memory net6.0 提示保留。
- `run-regular-gates.py` 串行完成 27 个项目/TFM 单元，`regular-results.json` 全部退出 0：20 个本次执行单元，加上 7 个未改变候选的 ClosedXML/Provider 合同/交叉集成/Docs 既有单元，共 4,975/4,975，无失败/跳过。没有运行无关 Large 性能矩阵。
- 严格检查当前 Git 中 211 个相关文本，210 个源码/文档/配置满足 UTF-8、BOM、EOL、EOF；唯一例外为 `common.props` 的 DocumentationFile 配置生成的 `tests/Bing.Offices.EntityProviderContracts/Bing.Offices.EntityProviderContracts.xml`，编译器产物使用 CRLF。保存原检查结果，按生成文件排除规则单列，不格式化或删除他人的存量文件。

### Excel 客户端与验证工具

- `run-advanced-excel.py` 的正式高级生命周期默认使用父进程创建、保有句柄和发布身份的隐藏专用 Excel，再附着该 PID 的原生窗口；保留原 `--owned-host` 开关兼容。显式 `--probe-script` 仍是独立诊断入口，不把诊断 Passed 当作客户端验收。
- 维持已经真实验证的 `/x /automation /r <generated fixture>` 只读启动，不改为空应用启动；原生对象获取依赖 EXCEL7 工作表窗口。只关闭自身实例的任务启动夹具，再按宏禁用、链接禁用和只读设置重新打开。Python、PowerShell 和 C# 窗口辅助类型各承担现有职责，不引入生产依赖。
- 已有 `owned-host-isolation-01/lifecycle.json` 两个故障合同成立：连接前 12 秒及连接后 20 秒外部超时，自身实例退出、清理无错误，任务自建的既有模拟实例仍存活并最终清理。该证据只覆盖任务模拟实例，不访问或外推用户交互工作簿。
- 最新当前包夹具 `owned-candidate-client-01.json` 仍为 Failed：V9 预期 3 实际为空，整体 Count 预期 5 实际为空；其余四聚合 150/30/10/50 正确，显式刷新诊断 Count 仍为空。约 43.76 秒、未超时、自身实例退出且 cleanupError=null。刷新及重开完整成功矩阵仍未完成。
- 原生版本 3/4/5/6 的 `count-version-diagnostic-01.json` Count 均为空，不用修改版本或生成静态汇总冒充修复。新增 `count-cache-options-01.json` 对照：Excel 不接受 worksheet cache 的 OptimizeCache 设置，四种选项均返回实际 COM 0x800A03EC；此结果否定通过该属性形成修复的方案，不宣称统计成功。参考仅为公开 API 文档：<https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache.optimizecache>。
- 新的原生全部轴小计对照 `diagnose-count-subtotals.ps1` 已准备并通过 PowerShell 解析，但两次正常会话启动请求均被自动审批检查超时挡住，进程未启动，不存在测试终态。新的默认入口负向矩阵也遇相同审批超时；沙箱尝试返回 WinError 1312（登录会话不可用）。这是验证环境/审批工具的实际限制，不是 Count 通过或业务代码失败。
- 根据实际启动失败补齐工具异常记录：父进程启动与身份发布进入既有 harness 异常/清理链；没有取得句柄时 `ownedExcelHandleObtained=false`、`ownedExcelExited=null`，不虚报已退出。生命周期哨兵启动失败也写出 `lifecycle.json`，未取得句柄时清理状态为未知；已有句柄只在 finally 清理自身进程。
- 实际沙箱启动失败直测：`default-owned-host-startup-failure-03.command.json`、`default-advanced-startup-failure-01.command.json` 均退出 1，记录 WinError 1312，JSON 正确保留 Failed、startupError、句柄未取得和退出未知；子 COM 探针未启动，没有误报通过。此负向异常记录成立，不替代正常登录会话的超时隔离成功矩阵。没有删除此前缺少 JSON 的失败目录。

当前 P1、常规自动化和 API/包交付证据已有实质进展；P2/P3 的 Count 完整数值合同、默认入口正常会话复验与最终独立 Review 仍保留。任务维持 IN_PROGRESS，未 task-finish，未 commit、push、PR、发布或改版本。

## 2026-10-05 后续实施：终态损坏报告与进程故障直测

- 对照第六次只读核查继续 F-002，当前运行状态仍为 implementing/IN_PROGRESS；没有覆盖或改写历史 Review，没有重新注册已启动任务。
- 定位外围工具缺陷：子进程完成且自身进程清理后，runner 在 finally 外直接读取结果 JSON；损坏 JSON 或非对象 JSON 会抛异常，导致最终 harness 缺失。现在将结果读取纳入失败报告链，新增 terminalError，保留清理证据、返回 Failed 和非零退出码；没有改写 Count/聚合业务合同。
- 新增 `build/PivotClientProbe/tests/test_owned_runner_failure_contracts.py`，九项标准库 unittest：启动无句柄、身份查询失败、owner 发布失败、子进程启动失败、非零退出、外部超时、正确终态控制组、非法 JSON、非对象 JSON。使用真实任务 Python 进程及独立哨兵，仅注入故障阶段与替换 Office 接口；不启动 Excel、COM 或操作用户进程。
- `owned-runner-failure-contracts-01.command.json` 退出 0，9/9 通过；身份与 owner 失败时未启动子探针，有句柄时确认 own 进程退出，哨兵仍存活。终态损坏明确返回 1、保留 ownedExcelExited=true、cleanupError=null，随后写出 Failed harness。README 增加可运行命令和证据边界。
- 该测试覆盖外围工具故障控制流，不替代 T-001 正常 Office 默认入口矩阵。先前正常会话的自动审批检查超时和沙箱 WinError 1312 仍保留，当前没有新的 Office 成功结果，也没有把待确认当作授权。
- Count 的原生所有轴小计对照尚未执行；真实当前包的文本 Count 组合仍 Failed。Microsoft 的公开计算说明确认 Count 统计非空值且区别于 Count Numbers，并未提供当前组合空值的确定根因或修复。参考：<https://support.microsoft.com/en-us/excel/calculate-values-in-a-pivottable>。
- Office 验证依赖尚缺时，独立核实 F-003 的本地 NPOI 2.7.4 公开接口与完整缓存可行性，产物放入 `BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001`。F-002 核心验收保持未完成，不据此关闭当前阶段或全 Goal；F-003～007、E-001/E-002 范围保留。

本轮未修改生产 Provider、公共 API、审批快照、依赖版本或仓库版本；只改变维护验证工具与职责测试。未 commit、push、PR、发布、发送外部通知或执行 task-finish。

## 2026-10-09 后续实施：透视缓存文本记录规范化

- 对 F-002 的最小失败形状继续定位：ClosedXML 生成的 `pivotCacheDefinition` 已声明文本 `sharedItems`，但对应 `pivotCacheRecords` 仍使用直接 `<s v="...">`；数值度量结构正常，Excel 客户端的文本 Count 仍为空。
- 新增内部 `ClosedXmlPivotCacheNormalizer`。它在生成透视工作簿保存到调用方流或模板桥接前，将每个缓存关系对应的字符串记录按字段共享项转换为 `<x v="...">`，保留无匹配项、数字记录、关系和其它 ZIP 部件；读取、重写和最终复制均响应取消。没有公开新 API，也没有公开或引用内部文件系统工具。
- `ClosedXmlExcelExporter` 对新建或更新透视先使用内存暂存，再执行规范化；模板透视保留-only 路径不经过生成器。模板桥接、提交、调用方流所有权以及现有同步/异步入口未改变。
- `AdvancedPivotExportTest.MultiAxisAndValues_ShouldPreserveCompleteWorkbook` 增加五个文本缓存字段的 `<x>` 和索引范围断言，覆盖同步/异步及筛选组合理论数据集。
- 隔离验证证据 `artifacts/closedxml-pivot-cache-normalizer-01.command.json`：net8/net6 源码编译均退出 0、0 警告/0 错误；net8 包消费者真实生成 XLSX，输出 `rows=3; normalizedTextFields=9`，逐条确认文本索引落在 `sharedItems` 范围内；新增职责测试在 net8、net6 各 6/6 通过，模板与重复导出组合测试 net8 2/2 通过。
- 本轮只修复 OOXML 缓存记录表示，不能替代 Excel 桌面端数值验收。已有 `owned-candidate-client-01.json` 的 Count 失败、默认入口正常会话复验和外部环境限制继续保留；F-002 不宣告完成，F-003～F-007 仍待后续阶段。

未修改版本号，未更新 API snapshot，未 commit、push、发布或发送外部通知。

## 2026-10-09 后续：混合类型共享项索引

- 修复 `ClosedXmlPivotCacheNormalizer.ReadFields`：为文本共享项保留其在全部 `sharedItems` 子项中的原始索引；重复文本继续使用首次出现的位置。
- 新增 `AdvancedPivotBoundaryTest.PivotCacheNormalizer_ShouldPreserveMixedSharedItemIndexes`，以数字项在前、文本项在后的最小透视缓存包验证数值记录保持原类型，文本记录引用索引 1。
- `net8.0` ClosedXML 高级透视边界职责测试经 `dotnet vstest` 执行 11/11 通过，覆盖新回归及同一测试类已有边界场景；项目构建 0 警告、0 错误。
- `net6.0` 构建未通过验证：现有 `project.assets.json` 指向 `C:\Users\CodexSandboxOffline\.nuget\packages`，ResolvePackageAssets 报缺少 `Microsoft.Extensions.Logging.Abstractions 6.0.0`；本机另一 NuGet 缓存存在该包，但本轮离线 restore 未能完成。故双 TFM 验收保持未完成，不将本轮标记为完整通过。
- 本轮未新增 API、依赖、CI、构建矩阵或基线；未修改版本号、未提交、未推送或发布。

## 2026-10-09 复验：混合类型共享项索引

- 使用隔离的 net6.0 输出目录和已构建测试程序集，重新运行 `PivotCacheNormalizer_ShouldPreserveMixedSharedItemIndexes`；测试通过 1/1。
- 本次复验关闭上一节记录的 net6.0 直接测试缺口。net8.0 同一测试已通过 1/1；双 TFM 构建结果及 net8.0 边界类 11/11 结果沿用本轮之前的真实产物。
- 本次只完成缓存规范化器的本地双 TFM 验证，不构成 Excel 客户端 Count 验收；F-002/F-003 状态及后续依赖不变。
- 验证仅使用系统临时输出，未更改项目 `obj/bin`、NuGet 配置、API 快照或 CI。

## 2026-10-09 后续：ClosedXML 透视缓存字段类型标志

- 将 `ClosedXmlPivotCacheNormalizer` 扩展为结合缓存共享项与关联缓存记录显式重写 `containsString`、`containsNumber`、`containsDate`、`containsBlank`、`containsMixedTypes`、`containsSemiMixedTypes`、`containsNonDate`、`containsInteger` 与 `longText`；判定规则与 NPOI 的 `SetSharedItemFlags` 对齐。修复未共享的直接文本记录没有反映到字段类型标志的问题。
- 扩展 `AdvancedPivotBoundaryTest.Aggregations_ShouldPreserveNumericAndCountSourceTypes`，并新增 `PivotCacheNormalizer_ShouldInferSharedItemFlagsFromInlineRecords`，覆盖真实导出字段及共享项为空/直接记录含文本的最小缓存包。
- 使用隔离临时目录构建 ClosedXML Provider 与测试项目。net8.0 构建成功，`AdvancedPivotBoundaryTest` 13/13 通过。net6.0 Provider 与测试项目编译成功；隔离 `vstest` 入口未发现测试，故 net6.0 运行结果不计为通过。构建期间有 NuGet 漏洞源不可达及既有注释/可空性警告。
- 本轮未运行 Excel 客户端；缓存元数据结构通过不等于文本 Count 显示正确。现有真实 Excel 文本 Count 失败与正常客户端复验限制继续保留，F-002 未完成，依赖它的 F-003/F-004 不前移。
- 未修改 API、版本、CI、快照或依赖；没有写入工作树 `obj/bin`，没有提交、推送或发布。

## 2026-10-09 后续：透视日期缓存系统双 TFM 往返职责

- 扩展 `NpoiPivotSourceValidatorTest.DateCacheExport_ShouldPreserveWorkbookDateSystem`，覆盖 1900 与 1904 日期系统下的透视缓存导出、缓存日期类型/索引、重开后的工作簿日期系统及来源日期。NPOI net6.0、net8.0 各 2/2 通过。
- 扩展 `AdvancedPivotBoundaryTest.DateCacheExport_ShouldPreserveWorkbookDateSystem`，使用 ClosedXML 公共 `Use1904DateSystem` 配置真实创建模板，覆盖两种日期系统的纪元日期和现代日期；验证原工作簿日期、`workbookPr`、透视缓存共享项/记录及重开日期。ClosedXML net6.0、net8.0 各 4/4 通过。
- 扩展 `NpoiPivotSourceValidatorTest.CountAggregation_ShouldAcceptMixedTypedAndEmptyValues`，让 Count 来源矩阵继续经过真实 NPOI 透视创建和 XLSX 序列化；核对 Count 字段、文本/数值/布尔/日期/空字符串/真空白的共享项类型、每行索引及重开透视结构。net6.0、net8.0 各 1/1 通过。
- 日期模板辅助方法变更后回归 `AdvancedPivotBoundaryTest.Aggregations_ShouldPreserveNumericAndCountSourceTypes`；ClosedXML net6.0、net8.0 各 1/1 通过，原 Count 类型和空值缓存断言保持成立。
- 初版 ClosedXML 序列夹具保留了原单元格 `t="d"` 类型，导致用例读到日期文本而非数值序列；改为通过 Provider 公共 Workbook API 生成日期系统模板后通过。未改生产 Provider 或缓存规范化逻辑。
- net6.0 首次 NPOI 运行复用了 net8.0 临时输出目录并触发 `System.Runtime, Version=7.0.0.0` 加载失败；按 TFM 分离输出目录后两项测试通过。一次隔离 NPOI restore 长时间无进展后终止；最终测试使用既有资产，并为 ClosedXML 单独完成关闭 NuGet audit 的临时 restore。工作树 `obj/bin` 未作为构建输出使用。
- 本节完成 D-001 中日期 Count 缓存往返和 NPOI 混合类型 Count 缓存结构的自动化子项，并完成 D-002 的 1900/1904 日期系统自动化子项；金额舍入、大整数及真实客户端 Count 值仍未完成。OOXML 与 Provider 重开通过不代表 Excel 客户端 Count 展示通过。
- 未新增 API、依赖、CI、矩阵或基线；未运行 Excel/WPS/LibreOffice，未更新版本或提交、推送、发布。

## 2026-10-10：F-002 T-001 数值与空值职责验证

- 变更影响：本轮没有生产代码变化；核查范围为高级透视来源预检、缓存字段类型、Count 混合值、筛选类型、错误坐标和目标流保全。
- ClosedXML `AdvancedPivotBoundaryTest` 在 net6.0、net8.0 各 17/17 通过；覆盖五种聚合配置、Count 的文本/数值/日期/布尔/空字符串/真空白、数值字段全空或混入非数值、类型化筛选、来源公式/错误及预检失败时的原始坐标与目标保全。
- NPOI `NpoiPivotSourceValidatorTest` 在 net6.0、net8.0 各 17/17 通过；覆盖 Count 混合类型与空值的真实透视缓存往返、数值聚合非法值坐标以及来源公式/错误。
- 构建和测试输出隔离到系统临时目录，复用仓库已有 NuGet 资产；未写入仓库 `obj/bin/output`。编译输出有 ClosedXML 既有 CS8632 和测试中的 xUnit2013 警告，无错误。
- T-001 标记为 VERIFIED。该职责验证只确认文件和缓存结构，不代表 Excel 客户端实际计算的 Count/五种聚合通过；客户端矩阵继续由 T-005 处理。下一项按优先级为 T-002 布局与请求隔离。
- 未修改生产代码、API、依赖、CI、快照或版本；未提交、推送或发布。

## 2026-10-10：F-002 T-002 布局与请求隔离验证

- 变更影响：验证 ClosedXML 高级透视的轴布局、多个度量/筛选、展开预算、格式边界、重复导出和同源透视配置隔离；未涉及 NPOI 生产改动。
- `AdvancedPivotBoundaryTest` 在 net6.0、net8.0 各 17/17 通过，覆盖仅列轴/多层列轴、多个筛选、精确与超限预算、最大行列边界和旧路径 `MaxExpandedCells`。
- `AdvancedPivotExportTest` 与 `AdvancedPivotCompositionTest` 合并筛选后在 net6.0、net8.0 各 8/8 通过，覆盖多值聚合与小计/总计结构、重复导出、同源透视配置隔离及失败后请求复用。
- T-002 标记为 VERIFIED。下一项按优先级为 T-003 资源生命周期与 Provider 边界；其测试范围保持独立。
- 本节仅更新任务证据，没有新增生产代码、API、依赖、CI、快照或版本，也未提交、推送或发布。

## 2026-10-10：F-002 T-003 资源生命周期与 Provider 边界验证

- 变更影响：验证高级透视同步/异步文件提交、目标保护、取消、异常观察、调用方流所有权，以及不支持 Provider 在枚举数据和写入目标前拒绝请求；没有生产代码修改。
- ClosedXML `AdvancedPivotResourceTest` 在 net6.0、net8.0 各 17/17 通过，覆盖成功提交、非法配置、预取消、来源枚举取消、暂存写入/序列化后取消、提交失败、观察次数和流所有权。
- NPOI `NpoiPivotResourceTest` 在 net6.0、net8.0 各 12/12 通过；公开 `NpoiPivotPublicExportTest.ExportToFile_FailureAndCancellation_ShouldPreserveTarget` 两个 TFM 各 2/2 通过。已通过的 `NpoiPivotSourceValidatorTest` 另覆盖预检取消与失败后重试隔离。
- Provider Contract `AdvancedPivotProviderBoundaryTest` 在 net6.0、net8.0 各 6/6 通过，覆盖 NPOI XLS、MiniExcel、SpreadCheetah 的同步/异步前置拒绝，并确认来源未枚举且目标流内容、位置和所有权未改变。
- T-003 标记为 VERIFIED。NPOI XLSX 的成功路径由高级透视职责测试覆盖；当前范围不把不支持 Provider 的拒绝测试扩大到 NPOI XLSX。
- 构建与测试程序集输出位于系统临时目录。编译保留既有 net6.0 支持提示及 CS8632 警告，无错误；未增加 CI、测试矩阵、基线、依赖或版本变更。
- 下一项按优先级为 T-004 双 TFM ClosedXML 职责及受影响合同回归；本节不代表 API、消费者、Release 全量门禁或 Excel 客户端 Count 验收通过。

## 2026-10-10：F-002 T-004 双 TFM 与受影响合同回归

- 变更影响：本轮只更新任务状态与证据记录；无生产代码、公共 API、Provider、TFM、构建/打包、CI 或测试矩阵变更。
- `AdvancedPivotNumberFormatTest` 在 net6.0、net8.0 各 2/2 通过；`AdvancedPivotModelTest` 与 `PivotTableContractTest` 合并后双 TFM 各 69/69 通过。
- 沿用本阶段已记录的 T-001～T-003 精准证据：ClosedXML 边界、导出组合与资源职责，以及 NPOI XLS/XLSX、资源和拒绝边界均有对应双 TFM 结果。T-004 标记为 VERIFIED。
- 历史失败仍保留，不用结构测试替代客户端数值验收。T-005 是下一项：文本 Count 在 Excel 16.0 Build 20430 中仍为空，原生对照也复现；另有正常 Office 会话启动被自动审批检查超时阻断、沙箱启动返回 WinError 1312。两类事实分别记录，启动阻塞不解释 Count 失败。
- Change Impact Analysis：ChangedFiles 仅为两份任务 Markdown；ChangedProjects、ChangedPublicContracts、ChangedRuntimePaths、ChangedProviders、ChangedTFMs、ChangedBuildPackaging、ChangedBenchmarkHarness 均为 none；ChangedDocsOnly=yes；AffectedDependents=none；RiskLevel=LOW。
- 验证：本轮不重跑 dotnet 测试，复用未变更源码对应的 T-004 双 TFM 结果；执行文档 UTF-8/BOM/LF/末尾换行和差异空白检查。未新增 workflow、job、矩阵、快照、依赖或构建步骤。
- 下一步仍严格按计划处理 T-005。当前其 Count 差异尚无确定根因，且正常客户端会话受外部工具/登录会话阻塞；不跳转到 F-003～F-007，也不将阻塞标为通过。

## 2026-10-10：T-005 已有 Count 组合产物复核

- 变更影响：只读取既有客户端 JSON/XLSX 产物并补充任务状态；没有新增 Excel/COM 启动、生产代码、测试代码、API 或构建配置变更。
- `count-diagnostic-07.json` 的原生 Excel 对照显示：文本 Count 单独、与 Sum/Average 组合、Count 字段重命名、Count 放在首位/中间，以及 Values 放在行轴时结果为 7；叠加 Minimum/Maximum 且 Values 位于列轴时 Count 结果为空。筛选和多列轴用例也为空。
- 该列轴组合在 Excel 16.0 Build 20430 的原生透视中同样复现，因此当前证据指向该客户端组合的兼容边界，但不能证明其为所有 Excel 版本的固有限制，也不能宣告既定合同通过。
- Values 改放行轴会偏离计划规定的多值列轴布局；本轮不改变已批准布局或 Count 来源来规避失败。T-005 仍未完成，需保留根因与版本范围不确定性。
- 验证：只解析已有 JSON 和工作簿部件列表，没有重跑客户端、dotnet 或 Python 测试；这是既有证据复核，不构成新的客户端成功结果。
- 下一步仍是 T-005：在允许的正常 Excel 会话中完成完整首开/刷新/保存重开矩阵，并检验是否存在保持列轴合同的修复。成功之前不进入 T-006 最终收口或 F-003。

## 2026-10-10：T-005 正常会话归属检查

- 变更影响：只记录会话可用性检查；没有修改生产代码、测试、API、构建配置或测试矩阵。
- 当前有两个 `EXCEL.EXE` 进程（PID 48668、52936，SessionId=1），但均不匹配本任务 `artifacts/client/*.owner.json` 中的 owner PID；其主窗口句柄为空。`Get-CimInstance Win32_Process` 返回拒绝访问，无法核实完整命令行或任务归属。
- 不附着、不终止这两个来源不明的 Excel 进程。当前没有可确认由本任务持有的客户端进程可继续等待；启动专用实例仍受 WinError 1312/审批超时限制。
- T-005 仍有文本 Count 列轴组合失败；原生 Excel 对照相同，Values 行轴替代布局不符合批准计划。当前没有足够证据支持符合原合同的生产代码修改。
- 状态：`BLOCKED_EXTERNAL`，阻塞于可安全启动并持有专用 Excel 实例的正常交互会话；T-005/T-006/T-007 和依赖它们的后续功能保持未完成。未重复启动相同客户端实验。

## 2026-10-10：D-001 数值往返职责子项

- 变更影响：只扩展 ClosedXML 与 NPOI 的高级透视数值边界职责测试及生产符号追溯；没有生产逻辑、公共 API、依赖、构建配置、CI 或 API 快照变更。
- 新增用例以应用显式按两位小数舍入后的金额及相邻 15 位整数为来源，验证 XLSX 重开后的业务单元格和透视缓存记录均保持数值类型与预期值。
- 初始 16 位整数夹具显示 ClosedXML/XLSX 数值存储会按 Excel 有效数字边界舍入；该结果不作为 Provider 回归，改用计划与文档合同内的 15 位数值输入，没有要求 Excel 数值保留精度范围外的值。
- `AdvancedPivotBoundaryTest.RoundedAmountsAndLargeIntegers_ShouldRemainNumericInCache`：net6.0、net8.0 各 1/1 通过。
- `NpoiPivotSourceValidatorTest.RoundedAmountsAndLargeIntegers_ShouldRemainNumericInCache`：net6.0、net8.0 各 1/1 通过。
- 测试项目采用 `BuildProjectReferences=false` 复用已有 Provider 程序集；本轮测试源码编译和用例执行通过。完整依赖项目构建仍受工作区既有 `output`/`obj` 生成源码重复编译及目录写权限影响，未修改用户的构建输出或项目配置。
- 本子项仅补充 D-001 的金额舍入和 15 位大整数自动化往返证据；真实客户端聚合值及 T-005 首开/刷新/保存重开仍未验收。T-005 保持 `BLOCKED_EXTERNAL`，不推进其依赖项。
- 未新增 workflow、job、矩阵、基线、依赖或版本变更；未提交、推送、发布或发送外部通知。
