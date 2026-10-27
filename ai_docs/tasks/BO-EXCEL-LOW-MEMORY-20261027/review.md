<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BO-EXCEL-LOW-MEMORY-20261027
AI_REVIEWED_AT: 2026-10-27T07:05:52Z

# 低内存 Excel 计划对照审查

## 审查结论

主要能力已经接入真实调用链：SXSSF 多 Sheet 顺序导出、ZIP/XML 单 Sheet 逐行导入、共享字符串数据与索引落盘、有界缓存、回调背压、独立服务注册，以及双 TFM/API/包消费者交付均有实现和证据。但不能维持“全部完成且无遗留问题”的验收结论：本轮确认 7 项 P2（MEDIUM / SHOULD_FIX），涉及数据语义、列选择、停止原因、错误定位、导出异常和临时文件额度/测量。

IMPLEMENTATION_STATUS: PARTIAL
TEST_STATUS: PARTIAL
PERFORMANCE_EVIDENCE_STATUS: PARTIAL
RESOURCE_EVIDENCE_STATUS: PARTIAL
EXTERNAL_GATE_STATUS: NOT_APPLICABLE
RELEASE_STATUS: NEEDS_FIX
OPEN_ACTIONABLE: 7

`TEST_STATUS: PARTIAL` 表示已有测试通过，但未覆盖本轮确认的缺陷；不是指本轮现有测试执行失败。`execution.md` 的 COMPLETED/PASS 是上轮执行记录，本轮没有修改该记录，当前独立验收以本报告为准。

## 范围与方法

- 对照用户批准的执行计划、既有 plan.md、execution.md、实际源码、工作区差异、新增文件、测试、API 成员追溯及容量/包验证产物。HEAD 为 `84958c1cc5eec5afbae51d76b0d07b4ab133b748`，工作区有未提交修改，暂存区为空。
- 本轮仅审查低内存任务，不把 CI、其他历史任务及原有未提交改动归入本任务缺陷。
- 使用 review-code 技能；主代理审查实现与边界，luna_worker 独立核验 LM-008 证据和文件提交配额。
- 未修改生产代码、正式测试、plan.md 或 execution.md。仅生成本报告、复现证据及 ignored artifacts 下的临时验证程序。
- 沿用明确批准的 1 万/10 万行、10 列验收范围及无固定 MB 门槛，不把 50 万/100 万行、DOM 模板/图片/图表等排除项列为缺陷。

## 计划对照

|计划项|判定|已交付|欠缺/关联问题|
|---|---|---|---|
|LM-001 基础契约与固定 NPOI 能力|PASS|NPOI 2.7.4、窗口/压缩/样式/清理测试及默认配置|本轮未发现阻塞架构的问题|
|LM-002 接口与 Sheet 快照|PASS|独立导入/导出接口、构建器、行结果/摘要、枚举追加、显式 DI|双 TFM 增量一致，未删除旧成员|
|LM-003 共享逻辑|PARTIAL|列绑定、转换/校验、单元格写入、样式、唯一值字符预算可复用|新旧日期文本和动态读取异常语义不一致：FIX-002、FIX-006|
|LM-004 低内存导出|PARTIAL|单次枚举、inline strings、写入即赋样式、窗口释放、流/文件出口|文件出口没有执行库自有临时文件总额度：FIX-001；动态读取异常：FIX-006|
|LM-005 ZIP/XML 与共享字符串|PARTIAL|实际关系解析、有限 XML、磁盘数据/定长索引、有界缓存、日期与公式缓存|日期值的文本表示不兼容：FIX-002|
|LM-006 导入转换、校验与回调|PARTIAL|逐行回调、背压、失败行 Item=null、有限错误样本和统计|列选择、额度错误位置、Stop 优先级：FIX-003～FIX-005|
|LM-007 生命周期与服务集成|PARTIAL|真实异步外围 IO、调用隔离、取消保护已有文件、清理主异常保留|动态读取异常绕开既有异常观察契约：FIX-006；文件总配额：FIX-001|
|LM-008 容量、API、文档与包|PARTIAL|8 个容量场景、真实字符串落盘、快照、成员追溯、文档和包消费者|导出临时峰值及残留检查范围漏掉提交文件：FIX-007；部分原方案补充指标未验证|

使用公开 `FlushRows()` 在字符预算紧张时清空窗口，替代固定版本中非公开的重载，是有依据的实现调整（DEVIATED_OK）。仍保留 netstandard2.0 Abstractions 和 NPOI 2.7.4，没有新增 Open XML SDK / IAsyncEnumerable 依赖。

## 本轮验证与证据复用

### Change Impact Analysis

本轮不改产品实现。验证范围是流式导入/导出、共享字符串、唯一值字符预算和生命周期职责测试，以及针对疑似缺陷的小型公开 API 复现。复现程序通过公开 DI 解析服务，与原 DOM 接口读取相同 XLSX 对照；临时盘验证装饰默认提交器，只观察同步写入时两个目录的真实文件长度。无需重跑完整单测、包验证和 10 万行矩阵。

|证据|结果|范围说明|
|---|---|---|
|本轮职责测试 net6.0|58/58 PASS|Streaming / XlsxSharedStringStoreTest / UniqueTrackerCharacterBudgetTest|
|本轮职责测试 net8.0|58/58 PASS|相同筛选，非零匹配|
|本轮小型复现 net6/net8|均确认 FIX-001～FIX-006 的实际行为|原始输出见 review-evidence.json；不是以“进程退出成功”代替功能验收|
|git diff --check|PASS|存在原工作区换行提示，无 whitespace 错误|
|复用上轮完整单测|每 TFM 1033/1033 PASS|完整测试后两处修正仅有每 TFM 47 项影响测试；不宣称最终全量 1034 已重跑|
|复用上轮集成/文档|集成各 TFM 40/40；文档 12/12 PASS|本轮未修改相关产品源码|
|复用 API/包证据|每 TFM Abstractions +67、NPOI +1、删除 0；4 个包消费者组合 PASS|有实际本地包 restore/执行证据，不等于所有边界行为正确|
|复用容量结果|8 个场景 success；高唯一导入确实落盘|耗时/内存与逐值校验可复用，导出临时文件峰值见 FIX-007|

复现入口：

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj --no-restore --filter "FullyQualifiedName~Streaming|FullyQualifiedName~XlsxSharedStringStoreTest|FullyQualifiedName~UniqueTrackerCharacterBudgetTest" --logger "trx;LogFilePrefix=low-memory-review" --results-directory artifacts/low-memory/review/test-results
dotnet run --project artifacts/low-memory/review/ReviewProbe.csproj -f net6.0 --no-restore --no-launch-profile
dotnet run --project artifacts/low-memory/review/ReviewProbe.csproj -f net8.0 --no-restore --no-launch-profile
```

原始 TRX 位于 `artifacts/low-memory/review/test-results/`；程序和项目位于 `artifacts/low-memory/review/`。初次同时构建复现程序与职责测试触发共享 obj 文件占用；职责测试自行重试后通过，复现程序随后串行构建运行通过。这是验证调度冲突，没有作为产品缺陷。

## 修复项

### FIX-001：文件出口绕过库自有临时文件总额度

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：`MaxTemporaryBytes` 只计 ExcelStreamingTempSession 中的完整暂存包。文件出口还通过默认提交器在目标旁创建另一份 `.tmp`，复制期间两份并存；第二份未计入额度。批准计划及公开选项只排除了 NPOI 自建临时文件，没有豁免提交文件。
- 证据：[NpoiExcelStreamingExporter.cs:38](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelStreamingExporter.cs:38)、[文件出口:82](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelStreamingExporter.cs:82)、[AtomicFileCommitter.cs:98](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs:98)。net8 小样本配置 76,694 字节，实际同时存在 145,196 字节库临时文件，输出 72,598 字节，调用仍成功；net6 同样复现。
- 影响：文件出口真实临时盘用量可接近已配置总额度的两倍，调用方无法依靠该选项限制库自有临时占用。
- 修复目标：同步/异步文件出口对所有本次库自有临时文件执行统一总额度。
- 修复要求：将提交文件纳入计量，或复用同一份可原子提交的暂存文件；保持默认/自定义提交器的兼容性、目标保护和 NPOI 自建文件排除规则。不能仅改文档缩小已批准的总额度语义。
- 验证方式：分别直接调用两个文件出口，设置介于单包大小和双包总量之间的额度；验证受限失败或单副本提交，完整断言原文件不变、清理无残留，并覆盖流出口不误算调用方目标流。

### FIX-002：日期单元格导入字符串时静默变成序列号

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：XML reader 识别日期格式后只更新 `ExcelCellValue.Value/Kind`，保留 `Text=raw`。字符串属性以及使用文本的校验/转换器拿到 Excel 序列号，与既有 DOM 语义不一致。
- 证据：[XlsxRowReader.cs:137](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Streaming/XlsxRowReader.cs:137)、[日期分支:163](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Streaming/XlsxRowReader.cs:163)，对照 [CellExtensions.cs:46](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Extensions/CellExtensions.cs:46)。同一个日期格式单元格映射到 string：DOM 返回 `2026-10-27 13:15:00.000`，流式返回 `46322.552083333336`，摘要仍为成功。两个 TFM 均复现。
- 影响：切换低内存入口后产生看似成功的不同业务数据；日期字符串、文本规则和基于文本的转换器受影响。
- 修复目标：XML 和 DOM 的日期文本与结构化值保持相同契约。
- 修复要求：统一标量文本材料化规则，覆盖普通日期和日期公式缓存，保留 1900/1904 日期系统；不得引入完整正文 DOM。
- 验证方式：相同 XLSX 分别由两入口导入 string/DateTime，并验证原始值校验和自定义转换器接收的完整上下文；覆盖两种日期系统、内置/自定义日期格式及缓存公式。

### FIX-003：未选中列的单元格错误使所选数据整行失败

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：表头绑定应用 `ReadColumns`，正文却先把 reader 在所有上限内列产生的错误加入行错误，未按绑定列筛选。
- 证据：[NpoiExcelStreamingImporter.cs:168](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelStreamingImporter.cs:168)。表头 Count/Unused，数据 7/#DIV/0!，请求 `ReadColumns(0, 1)`：DOM 得到 Count=7 且成功，流式得到 Item=null 和 B2 错误。两个 TFM 均复现。
- 影响：调用方无法通过列选择避开无关计算列，未映射列会导致有效业务数据丢失。当前直接构造的错误也丢失 PropertyName、ColumnKey、Header、RawValue；选中错误单元格与 DOM 对照已经观察到这些字段全部变为 null。
- 修复目标：只让参与当前绑定的值错误影响实体材料化，同时保持包结构和资源检查。
- 修复要求：区分全局 XML/资源错误与可定位单元格转换错误；不因忽略列而关闭资源限制。所选列的 Excel 错误表达应保留既有结构化诊断字段。
- 验证方式：同步/异步分别覆盖显式范围外、未映射列、所选列中的错误单元格和无缓存公式；断言完整实体与结构化错误，并验证忽略列仍受 XML/文本额度约束。

### FIX-004：MaxRows 错误定位偏移一行

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：先执行 `expected=row.Index+1`，再检查 MaxRows，catch 又用 `expected+1` 作为一基错误行号。
- 证据：[NpoiExcelStreamingImporter.cs:154](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelStreamingImporter.cs:154)、[错误创建:160](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelStreamingImporter.cs:160)。表头第 1 行、数据第 2/3 行、MaxRows=1：DOM 报第 3 行，流式报第 4 行。两个 TFM 均复现。
- 影响：错误位置可能指向不存在的行，业务无法准确定位或提示额度边界。
- 修复目标：超限诊断使用实际触发的物理一基行号。
- 修复要求：将当前行错误位置与下一期望行位置分离；读取期间的额度错误也不能错误借用上一次游标，尤其要处理稀疏行。
- 验证方式：MaxRows 边界、非默认 DataRowStartIndex 和稀疏行小样本，断言准确行号、交付数量、ResourceLimit、IsComplete=false；另覆盖 reader 在行内触发文本额度的定位。

### FIX-005：消费者 Stop 被错误阈值覆盖

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：消费者返回之后，代码先返回 ResourceLimit/ErrorLimit，再判断 Stop；与已批准的“返回 Stop 时使用 ConsumerStopped”不一致。
- 证据：[NpoiExcelStreamingImporter.cs:187](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelStreamingImporter.cs:187)。MaxErrors=1，首行转换失败，回调返回 Stop，实际摘要为 ErrorLimit。两个 TFM 均复现。
- 影响：调用方不能稳定区分主动停止和库主动终止，依赖 StopReason 的续传/提示流程出现歧义。
- 修复目标：已交付回调明确返回 Stop 时，摘要遵循 ConsumerStopped 契约并保留已有错误统计。
- 修复要求：统一普通行错误、错误阈值和行级资源错误与消费者决定同时出现时的优先级；取消及消费者异常仍直接传播。
- 验证方式：Continue/Stop 分别与错误阈值、唯一值额度相交，精确断言原因、计数、样本和后续行未读取，覆盖同步/异步入口。

### FIX-006：动态导出读取器异常丢失公共异常上下文

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：流式 WriteSheet 直接调用 DynamicGetter，没有复用原导出对用户扩展错误的转换。普通 getter 异常原样漏出，也绕过仅观察 BingOfficesException 的观察器。
- 证据：[NpoiExcelStreamingExporter.cs:196](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelStreamingExporter.cs:196)，对照 [NpoiExportSheetWriter.cs:61](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExportSheetWriter.cs:61)。预声明动态列的 getter 抛 InvalidOperationException：DOM 返回 BingOfficesExportException，流式返回原 InvalidOperationException；两个 TFM 均复现。
- 影响：新入口不满足计划要求的既有转换/错误语义；统一异常处理失去 UserExtensionFailed、Sheet、行号、属性及 InnerException 结构，观察器漏报。
- 修复目标：动态值读取采用相同公共错误边界，并保持每实体只读取一次。
- 修复要求：复用小型共享帮助器；保留已有 BingOfficesException 和取消异常，包装普通用户扩展异常，不重复通知观察器。
- 验证方式：同步/异步和流/文件入口直接覆盖 getter 异常；断言类型、Code、Stage、Sheet/行/属性、InnerException、观察器一次及资源清理；补 HSSF/XSSF 共享帮助器回归。

### FIX-007：导出临时峰值和残留检查漏掉提交文件

- 严重程度：MEDIUM（P2）；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：ResourceProbe 的计时覆盖提交，但临时盘采样和最终残留统计只扫描库 session 目录与 NPOI poifiles，遗漏目标文件旁的原子提交 `.tmp`。这与 FIX-001 的额度实现问题相关，但即使修复额度，测量器也必须覆盖最终实现拥有的提交文件。
- 证据：[ExcelStreamingResourceMatrix.cs:602](/H:/Bing_Framework/Bing.Offices/tests/Bing.Offices.ResourceProbe/ExcelStreamingResourceMatrix.cs:602)、[残留统计:141](/H:/Bing_Framework/Bing.Offices/tests/Bing.Offices.ResourceProbe/ExcelStreamingResourceMatrix.cs:141)、[AtomicFileCommitter.cs:43](/H:/Bing_Framework/Bing.Offices/src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs:43)。容量表的导出库临时峰值不包含第二份完整包。
- 影响：目前不能以该表证明“全部库临时文件的全生命周期峰值/无残留”；API 耗时、进程内存及已监测目录结果仍有价值，不需全盘作废。
- 修复目标：准确报告本次调用的全部库临时文件占用和清理结果。
- 修复要求：包含提交目录中本次创建的文件，排除原目标和其他调用文件；同时记录报告所需机器信息。原方案的 DOM 对照、分配量、GC 和取消完成延迟目前没有证据，应补充或明确标注未验证，不能由现有 8 份 JSON 推导。
- 验证方式：用可控提交窗口的小型测试验证采样范围与残留检出；额度/探针修改后仅重跑受影响导出场景并更新原始 JSON/报告。未受影响的导入场景、API 和包证据继续复用；不因补机器说明重跑重型矩阵。

## 分组状态与后续验收

|分组|内容|
|---|---|
|Completed|主要能力实现已核对；当前职责测试每 TFM 58 项通过；6 项行为复现及资源探针审查完成；报告已落盘|
|Open Actionable|FIX-001～FIX-007；默认修复范围 SHOULD_FIX，共 7 项|
|Blocked Approval|无；已批准的无固定 MB 门槛无需重新审批|
|Blocked External|无本地审查阻塞；外部 CI/发布不在本次检查范围|
|Not Applicable|模板、图片、图表、复杂合并、公式计算、跨 Sheet 关系不是本次实现缺口|
|Accepted Limitations|NPOI 打包/内部写入同步；NPOI 系统临时文件不纳入库自有额度；外部流与消费副作用不回滚|
|Verified Boundaries|高唯一导入真实落盘；已有字符/唯一值/ZIP 等职责额度测试通过；容量只覆盖 10k/100k × 10 列|
|Deferred|50 万/100 万行与生产容量 SLA；无新增扩展目标|

后续修复应先补这些缺陷的职责级直接测试，再执行受影响的双 TFM 回归。资源报告应在 FIX-001/FIX-007 完成后更新。当前请求为检查，审查到此结束；本轮没有自动进入修复。
