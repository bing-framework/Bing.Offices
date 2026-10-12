<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001
AI_EXECUTION_FINISHED_AT: 2026-10-11T00:20:00+08:00

AI_EXECUTION_UPDATED_AT: 2026-10-11T00:20:00+08:00

# 实施执行报告

## 当前结论与范围

REVIEW_FIX_COMPLETED_WITH_ISSUES。P1 的 OPC 部件清单、Manifest 与默认拒绝已有双 TFM 门禁；P2 的公开 Builder/Request 描述、深快照和内部更新 Plan 已完成双 TFM 定向验证。P3 的库侧范围已完成：两个 Provider 均支持原样保留、业务 Sheet 写入、全部或部分模板透视来源更新、共享缓存拆分、布局保留、目标冲突预检、来源收缩清理和新建/更新组合；未知扩展及命名来源按安全边界保留或拒绝更新。更新缓存写入 `refreshOnLoad=true` 请求，但 Provider 不执行客户端刷新。按用户确认，本任务排除 Excel/WPS/LibreOffice 客户端行为；因此不把首开、刷新显示和保存重开列为代码 TODO，也不声称客户端行为已验证。PTPL-004 库侧业务/失败合同已完成；PTPL-005 的候选身份、包消费者和独立 Review 已完成，本任务仍因正式 API 门禁中 19 个其它任务成员审批未完成而保持部分完成。

F-003 前序候选 Release +27常规单元 5113/5113、包/API及NPOI基础客户端矩阵已有证据并完成记录同步；高级 NPOI 首开 Count 与 F-002 同样为空，完整数值未通过。F-002/003、F-005～007 和 E-001/E-002 仍属于完整 Goal。新生产变更使源身份进入 F-004 候选，前序通过只对应前序快照，最终新候选须重新绑定。

## 阶段矩阵

| 阶段 | 状态 | 当前结果 |
| --- | --- | --- |
| PTPL-001 清单与依赖图 | PASS | 两 Provider 清单/关系图及默认拒绝双 TFM 验证通过；关系部件读取贯通 XML 字符数与深度预算，ClosedXML 导出、实体模板导入、Forward 批次和失败工作簿预检均有直接测试 |
| PTPL-002 公共快照与预检 | PASS（子集） | Builder/Request 保留与更新描述、磁盘 Snapshot 已实现；NPOI/ClosedXML net6/net8 快照职责均通过；全部或部分模板透视更新在模板创建阶段完成预检，共享缓存更新进入拆分路径 |
| PTPL-003 保留/更新输出 | PASS（用户范围内） | NPOI/ClosedXML 的结构保留、业务 Sheet 写入、全部/部分更新、缓存拆分、布局保留、冲突预检、来源收缩及组合请求均有双 TFM 合同覆盖；客户端刷新/显示按用户范围排除且未验证 |
| PTPL-004 业务合同 | PASS（库侧） | 采购/仓储业务组合、来源变化、共享缓存、请求复用、失败重试和未知扩展边界均有双 Provider/双 TFM 合同；不包含客户端验收 |
| PTPL-005 交付 | PARTIAL（审批阻塞） | 普通/第三方 NuGet 消费者、Docs 示例、Release 构建和独立 Review 通过；正式双 TFM API snapshot 仍受 19 个非本任务成员审批阻塞 |

## 修改

- 新增 provider-shared/ExcelPivotTemplateParts.cs：依据准确ContentType、过渡/严格关系Type及关系拥有者URI发现真实Pivot/CacheDefinition/CacheRecords；标准目录保留兼容发现。元数据逐元素读取，DTD/外部解析禁用，取消保留令牌，清单稳定只读；缺失/外部关系和重复部件抛InvalidDataException。调用方ZipArchive保持持有。
- 两个Provider项目编译链接内部实现，不新增生产友元；P1 仅内部变更，P2 的公共 SPI 单列在成员审批及下方记录。
- ClosedXmlTemplatePreflight 复用清单，保持原Chart/Pivot/其它默认拒绝顺序，恢复模板输入位置/所有权。NpoiReportWriter 使用引擎公开ContentType识别非标准路径的透视；旧新建和XLS拒绝边界不变。
- Luna独占两个Provider的P1直接测试，主线程负责集成与dotnet验证。

## 状态与验证

任务注册首次在受限 .agents/runtime 临时文件遇EPERM，仅注册命令升级后成功：active=true、mode=plan-execution、startedAt=2026-10-05T11:57:23.224Z。不改任务脚本、发送通知或Finalizer。F-003历史执行保持IN_PROGRESS及真实未完成项。

生产编译、P1直接职责及后续门禁正在执行；无真实终态的项目不写PASS。计划不改写。C# UTF-8 BOM/LF，项目CRLF，文档无BOM/LF，最小引用修改。

未自动commit/push/PR/发布/修改版本/通知。

## 2026-10-05 后续：P1 验证与 P2 请求/Plan

- P1 两个 Provider 的常规 Release 门禁 `p1-results-01.json`：NPOI net6/net8 各 1154/1154，ClosedXML net6/net8 各 466/466，合计 3240/3240，四条命令均退出 0。
- P1 部件/Manifest 直接职责 `p1-npoi-net6.0-01.trx`、`p1-npoi-net8.0-01.trx`、`p1-closedxml-net6.0-01.trx`、`p1-closedxml-net8.0-01.trx` 各 42/42，合计 168/168。`p1-budget-npoi-net8-01.trx` 与 `p1-budget-closedxml-net8-01.trx` 各 32/32；Docs `p1-docs-net8-01.trx` 19/19。以上证明清单/预算/文档示例，不证明模板透视保留或更新输出。
- 新增 P2 请求模型和 Provider 拒绝合同后，首次 `p2-contract-net8-01` 编译失败，原因是 `TemplatePivotBoundaryRow` 重复 public 无参构造；已删除重复声明。失败记录保留。后续 `p2-contract-net8-02.trx` 退出 0，63/63（模板模型、各 Provider 明确拒绝模式及 F-002 模型回归）。只证明 net8 当前合同。
- `p2-plan-npoi-net8-01.trx` 与 `p2-plan-closedxml-net8-01.trx` 各 43/43，验证内部 Plan 的精确名称/Sheet 匹配、来源与坐标、已有名称冲突和失败重用。net6 对应验证仍待运行。
- `TemplatePivotModelTest.Request_ShouldIsolateBuilderMutationsMadeAfterCreation` 补充 Builder 返回后继续 AddSheet/Preserve/Update 时，既有 Request 的 Sheet/Updates/bool 快照不变及集合只读；已包含于 `p2-contract-net8-02.trx` 的 63/63，net6 对应验证待完成。
- 主线程正在实现内部模板 Snapshot 的真实磁盘读写、预检和资源保护。此路径尚未进入公开 Provider 输出，P2 不标记完成；P3 保留/更新序列化、P4 客户端和 P5 交付仍待实施。
- 当前任务继续 IN_PROGRESS；没有声明模板能力位，没有把 F-002/F-003 或 F-005～007、环境验收并入 F-004 完成范围。

## 2026-10-06 后续：透视部件保留输出

- 新增 `ExcelPivotTemplatePreserver`：仅保留模板且没有业务 Sheet、字典、元数据或更新定义时，NPOI/ClosedXML 先完成磁盘 Snapshot、Manifest/Plan 预检，再以可取消复制保留原始 ZIP 字节。
- 新增 `ExcelPivotTemplateBridge`：带业务 Sheet 的保留请求使用剥离透视关系的中间模板，Provider 完成普通 Sheet 写入后，将原始透视表、缓存、关联关系和内容类型按工作表名称重新合并；模板输入位置和所有权保持原契约。
- `p2-preserve-output-net8-03.trx`：ProviderContract 保留/拒绝边界 20/20；`p3-preserve-business-net8-02.trx`：NPOI/ClosedXML 同步和异步业务 Sheet 往返 4/4。断言包含实际 pivotTable、pivotCacheDefinition、pivotCacheRecords 部件、业务数据读回和输入流位置。
- `p2-p3-preserve-net6-followup.trx`：ProviderContract net6 保留/更新拒绝边界与完整业务往返 24/24；`p1-npoi-manifest-net6-followup-serial.trx`、`p1-closedxml-manifest-net6-followup-serial.trx` 各 57/57。此前并行执行造成临时文件基线互相干扰，改为串行后两 Provider 资源场景均通过。
- `p3-preserve-net8-followup.trx`、`p3-preserve-net6-followup.trx` 各 12/12：NPOI/ClosedXML 同步与异步预取消均保持模板位置、目标内容、流所有权和取消令牌。
- `p2-model-net6-followup.trx`：模板模式公共模型与快照隔离测试 8/8；P2 的双 TFM 模型证据已补齐，后续已接入独立缓存更新子集。
- NPOI 与 ClosedXML net8 Provider Release 构建均成功（0 警告、0 错误）；`git diff --check` 无差异错误。仓库既有 `ClosedXml.csproj`、`Npoi.csproj` 和 ProfileFixtures XML 行尾提示仍存在，未在本任务中重整。
- 该阶段保留实现不刷新原缓存；当时尚未消费 `UpdatePivotTable` 定义。后续已接入独立缓存更新子集，仍未新增 `TemplatePivotTables` 能力位，也未以局部输出证据替代 P3 完整更新验收。

## 2026-10-06 后续：缓存更新子集

- `ExcelPivotTemplateBridge` 在模板 Snapshot 完成后即校验临时 Sheet 名称；Provider 生成的每个缓存只能被本次请求中的更新透视引用。共享缓存的被更新消费者由生成缓存承接，未更新消费者继续使用原缓存，临时 Sheet 名称冲突在 Provider 枚举业务数据前返回 `UnsupportedFeature`。
- NPOI 与 ClosedXML 均通过临时业务 Sheet 生成同名透视和新缓存，删除临时 Sheet 后把新透视关系回挂到原目标 Sheet；未更新透视及其缓存按原始部件复制，生成部件路径和 `cacheId` 会避开保留部件。
- `p3-template-pivot-net8-shared-split.trx` 与 `p3-template-pivot-net6-shared-split.trx`：`TemplatePivotPreserveExportTest` NPOI/ClosedXML 同步和异步共 36/36、36/36；验证保留、单透视更新、双独立缓存透视更新、部分更新保留、共享缓存拆分及反向更新路径的缓存来源值、目标透视名称和临时 Sheet 清理。`p3-update-boundary-net8-shared-final.trx` 与 `p3-update-boundary-net6-shared-final.trx` 另以 12/12、12/12 验证无实际模板透视时更新在业务数据枚举前拒绝。
- Provider 限制和透视文档已改为明确说明该更新子集。未新增能力位；刷新、任意扩展更新和 Excel/WPS/LibreOffice 最终显示验收仍列为后续 TODO，未知扩展的原样保留与安全拒绝已有直接证据。

## 2026-10-06 后续：模板依赖安全边界

- `ExcelPivotTemplateManifest` 现在为每个透视和缓存记录未验证依赖标记：未知 `ext` 扩展、非 Provider 契约关系、缓存命名来源以及缓存记录扩展均进入安全拒绝集合；ClosedXML 自身已知的 `x14:pivotTableDefinition` 扩展按精确 URI、命名空间和元素形状放行。
- `ExcelPivotTemplatePlan` 在定位更新后、业务数据枚举前拒绝带未验证依赖的透视；原样保留模式不重写这些部件，因此扩展内容继续保留。该边界不声称已支持切片器、时间线、外部连接或任意扩展更新。
- `p3-template-dependency-final-net8.trx` 与 `p3-template-dependency-final-net6.trx` 各 14/14：NPOI/ClosedXML 同步和异步更新在扩展模板、命名来源模板及未声明关系模板上保持输入、目标流和未枚举状态；保留模式完整保留扩展标记。`p3-template-pivot-net8-dependency-final.trx` 与 net6 对应文件各 46/46，另外确认普通保留、单/多/部分更新及共享缓存拆分未受误伤。
- 该阶段仍不执行客户端刷新；未知扩展的最小支持范围是“保留或明确拒绝更新”，后续如需更新带切片器/时间线的模板，必须先建立对应部件依赖和客户端验收。

## 2026-10-06 后续：更新缓存打开时刷新请求

- `ExcelPivotTemplateBridge` 在合并全部或部分模板更新时，只对 Provider 本次生成的 `pivotCacheDefinition` 写入 `refreshOnLoad=true`；模板中保留的未更新缓存及其 XML 不改写。
- `TemplatePivotPreserveExportTest.UpdateSinglePivot_ShouldWriteNewCacheRecords` 已增加缓存定义属性断言，NPOI/ClosedXML 同步和异步 net8.0、net6.0 各 4/4 通过，证据见 `p3-template-refresh-on-load-net8.trx`、`p3-template-refresh-on-load-net6.trx`。
- 该标记只表示客户端打开时的刷新请求，不能替代 Excel/WPS/LibreOffice 的真实刷新、首开显示或保存重开验收；没有新增公共刷新 API，也不放宽未知扩展、切片器、时间线和外部连接的更新边界。
- 模板更新回归 `p3-template-pivot-refresh-regression-net8.trx`、`p3-template-pivot-refresh-regression-net6.trx` 各 36/36；Docs `Category!=Large` 使用已有程序集运行 `p3-template-refresh-docs-net8.trx` 19/19。一次 Docs 重新构建因 Abstractions XML 文档文件被占用返回访问拒绝，未作为代码失败；随后 `--no-build` 验证通过。

## 2026-10-06 后续：模板透视能力声明

- `ExcelProviderFeatures.TemplatePivotTables = 1 << 26` 已加入隐藏 Provider SPI；NPOI XLSX 与 ClosedXML XLSX 的导出器在能力描述中声明该成员，MiniExcel 与 SpreadCheetah 保持未声明。
- `TemplatePivotProviderBoundaryTest` 在 net8.0、net6.0 均通过 12/12，覆盖 NPOI/ClosedXML 的保留与更新请求以及 MiniExcel/SpreadCheetah 的结构化拒绝，并验证能力位与实际边界一致。
- 同一职责在 Release 测试资产下再次通过：net8.0 12/12、net6.0 12/12；net6.0 仅出现既有 Microsoft.Bcl.Memory TFM 支持提示。
- 能力位的语义限定为 XLSX OPC 透视部件的原样保留、部分或全部更新、缓存拆分及打开时刷新请求；不承诺 Excel、WPS 或 LibreOffice 的刷新、首开显示、保存重开结果。
- 第五项 API 成员已纳入当前候选范围；正式双 TFM snapshot、NuGet 候选身份和维护者基线审批仍属于 P5，未修改正式基线。客户端环境缺少 Excel、WPS 和 LibreOffice，P4 的真实客户端验收继续保留为待办。
- 候选快照最终保存在 `artifacts/api-template-pivot-20261006-final`，双 TFM 的结构差异保留在 `artifacts/api-template-pivot-20261006-compare/api-diff.json`；差异包含第五项能力位和本工作树中尚未合并的前序公共成员，未发现删除成员。比较命令按现有正式基线执行时因候选源清单、程序集身份和审批路径仍绑定前序任务而退出 1，未用覆盖基线的方式掩盖差异。

## 2026-10-07 后续：当前 NPOI 缓存修复后的模板回归

- 在 `NpoiSheetContentWriter` 增加缓存定义关系定位及 `sharedItems` 文本类型标记重写后，重新执行 Provider Contract 模板筛选：`net8.0` 70/70、`net6.0` 70/70，命令均退出 0。
- 本次筛选覆盖模板模型快照、Provider 能力边界、完整透视部件保留、业务 Sheet 写入、单/多/部分更新、共享缓存拆分、依赖安全拒绝和 `refreshOnLoad` 请求；结果表明 NPOI 缓存定义修复未回归既有模板保留/更新结构。
- 本次证据仍只证明生成包的 OPC 结构和 Provider 合同，不证明 Excel、WPS 或 LibreOffice 的实际刷新、首开显示、保存重开，也不解决高级文本 `Count` 客户端首开结果为空的问题。F-004 客户端验收、F-005/F-007 外部客户端、F-006 解析期低保留和 E-002 商业 Provider 仍保持 TODO。

## 2026-10-07 后续：模板合并资源与取消边界

- `ExcelPivotTemplateBridge` 的关键 XML 解析路径现在统一接收模板快照的 `MaxXmlCharacters`；缓存 ID 重映射、部件重命名、内容类型合并、临时 Sheet 移除和关系回挂不再退回固定默认字符预算。
- ZIP 条目读取和 XML 读取均按块检查取消令牌；XML 解析通过可取消的读取流连接到 `XmlReader`，在解析中取消时保持既有异常、目标流和输入流所有权契约。
- NPOI `ExcelPivotTemplate` 筛选在 net6.0/net8.0 各 75/75 通过；ClosedXML 同筛选在 net6.0/net8.0 各 75/75 通过；Provider Contract `TemplatePivot` 在 net6.0/net8.0 各 70/70 通过。Release 编译两个 Provider 均成功，除既有 NU1900/TFM 支持提示外无错误。
- 本次只收紧内部资源和取消边界，不改变模板透视的保留、部分/全部更新、共享缓存拆分或 `refreshOnLoad` 语义；`ReadArchive` 仍将 ZIP 部件加载到内存，外部可配置导出资源限制、Excel/WPS/LibreOffice 客户端验收、高级文本 `Count` 首开结果和任意扩展更新继续保持 TODO。

## 2026-10-07 后续：最终归档临时文件落盘

- `ExcelPivotTemplateBridge.CreateArchive` 不再使用完整 `MemoryStream` 组装最终 ZIP，改为使用 `DeleteOnClose`、异步和顺序访问选项创建临时文件流；透视保留路径和全部/部分更新路径共用该实现，复制到调用方目标流前保持定位在文件开头。
- Provider Contract `TemplatePivotPreserveExportTest.PreserveWithBusinessSheet_ShouldCleanArchiveTemporaryFile` 在 NPOI/ClosedXML 同步和异步组合下验证归档临时文件在成功返回后全部清理；模板部件、业务数据、输入位置和流所有权由同一筛选继续覆盖。net8.0、net6.0 `TemplatePivot` 筛选各 74/74 通过，证据为 `pivot-archive-net8.trx`、`pivot-archive-net6.trx`。
- 该增量只消除最终 ZIP 的一份完整内存副本；`ReadArchive` 及其部件字典仍按现有资源预检后驻留内存，尚未完成按部件的磁盘/惰性存储重构。取消、临时文件和调用方目标流行为已保留；Excel/WPS/LibreOffice 客户端验收、高级文本 `Count` 首开结果、任意扩展更新和完整懒加载继续保持 TODO。

## 2026-10-07 后续：透视桥接按需部件存储

- 新增内部 `ExcelPivotTemplateArchive`：原始/生成 ZIP 先复制到 `DeleteOnClose` 临时包，只建立部件索引；`TryGetValue` 首次访问时读取单个部件，修改或新增部件单独保存在覆盖集合，最终 ZIP 按名称逐条流式写出。部件删除、改名、缓存标识重映射及 XML 覆盖继续使用原有桥接流程。
- `ExcelPivotTemplateBridge.ReadArchive`、`ReadPackage`、透视保留/全部更新/部分更新路径已切换到该存储；中间去透视模板 `Strip` 也改为 `DeleteOnClose` 临时文件流，NPOI 与 ClosedXML 仍通过可定位的 `Stream` 加载，调用方模板流位置和所有权不变。
- `TemplatePivotPreserveExportTest.PreserveWithBusinessSheet_ShouldCleanArchiveTemporaryFile` 同时检查 `archive-*`、`parts-*` 和 `stripped-*` 三类临时文件。ProviderContract `TemplatePivot` net8.0/net6.0 各 74/74，NPOI net8.0 77/77，ClosedXML net8.0 75/75 通过；证据为 `pivot-lazy-stripped-net8.trx`、`pivot-lazy-stripped-net6.trx`、`pivot-lazy-stripped-npoi-net8.trx`、`pivot-lazy-stripped-closedxml-net8.trx`。
- 本次将原始和生成包的全量部件 `byte[]` 字典移除，但 Provider 生成阶段仍按现有实现使用 `MemoryStream`，XML 单个部件及修改覆盖仍会短暂进入内存；Excel/WPS/LibreOffice 客户端验收、高级文本 `Count` 首开结果、任意扩展更新及 Provider 生成阶段外置暂存继续列为 TODO。

## 2026-10-07 后续：Provider 生成包外置暂存

- NPOI 与 ClosedXML 在启用透视模板桥接时，Provider 生成目标从 `MemoryStream` 改为 `DeleteOnClose`、异步和顺序访问的临时文件；无模板或无需桥接的普通导出路径保持原有目标流行为。
- 资源回归同时检查 `generated-*` 临时文件；ProviderContract TemplatePivot net8.0/net6.0 各 `74/74`、NPOI net8.0 `77/77`、ClosedXML net8.0 `75/75` 通过，证据为 `pivot-generated-file-net8.trx`、`pivot-generated-file-net6.trx`、`pivot-generated-npoi-net8.trx`、`pivot-generated-closedxml-net8.trx`。
- ClosedXML 异步 API 外层仍使用既有外围缓冲流，NPOI 异步 API 仍由既有 staging 管理；本次只移除桥接内部生成包的一份额外内存副本。若要让 ClosedXML 异步整包也完全外置，需要单独处理其外围缓冲生命周期和异步复制契约。
- ClosedXML Workbook `ExportAsync` 在启用透视模板桥接时也改用 `async-*` 临时文件作为外围缓冲，普通导出、实体异步导出和保留-only 快速路径不变；资源回归已把该文件纳入清理集合。ProviderContract TemplatePivot net8.0/net6.0 各 `74/74`，ClosedXML net8.0 `75/75` 通过，证据为 `pivot-outer-staging-net8.trx`、`pivot-outer-staging-net6.trx`、`pivot-outer-staging-closedxml-net8.trx`。

## 2026-10-07 后续：导出模板资源限制接入

- `ExcelWorkbookExportBuilder.ResourceLimits(ExcelResourceLimits)` 与隐藏的 `ExcelWorkbookExportRequest.ResourceLimits` 已加入当前候选；保留旧的内部构造重载，未改变未配置时的默认行为。构建阶段验证限制值，Request 保存引用快照，模板透视路径将其传给 `ExcelPivotTemplateSnapshot`、`ExcelPivotTemplatePreserver` 和 `ExcelPivotTemplateBridge`，不再硬编码默认限制。
- `TemplatePivotPreserveExportTest.ResourceLimits_ShouldRejectTemplateBeforeBusinessEnumeration` 与 `ResourceLimits_ShouldRejectPreserveOnlyTemplate` 在 NPOI/ClosedXML 同步与异步组合下验证 `MaxInputBytes`：限制错误在业务集合枚举或保留复制前抛出，模板输入位置、调用方目标流内容及流所有权保持不变。`pivot-resource-limit-net8.trx` 与 `pivot-resource-limit-net6.trx` 各 `48/48` 通过。
- 公共 API 全量门禁 `pivot-resource-api-all-net8.trx` 仍按预期报告当前工作树中前序批次的 24 条未批准成员差异，其中包含本轮新增的 ResourceLimits 两条；没有更新正式基线。Excel/WPS/LibreOffice 客户端验收、高级文本 Count 首开结果、任意扩展更新、F-005、F-006 以及 F-007 的复杂布局/低保留边界仍为 TODO。

## 2026-10-07 后续：Forward 实体失败工作簿模式

- NPOI 与 ClosedXML 的 XLSX Forward 实体批次现在允许 `ErrorRowsOnly` 和 `AnnotatedOriginal`。正常路径仍使用 XML 前向解析并按原有批次回调交付；只有解析和回调完成且存在错误时，才重新打开已缓冲的原始 XLSX，复用 Provider 失败工作簿写入器生成错误行产物或原始批注产物；未知模式仍在计划阶段结构化拒绝。
- NPOI 同步和异步分别复用既有文件提交服务的同步/异步边界，并通过 `NpoiNonDisposingStream` 保持调用方输入流所有权；ClosedXML 通过导入器已有失败工作簿入口复用 DOM 准入、布局行预检和提交行为。没有新增公共 API 或生产友元。
- 直接职责测试：NPOI `EntityLayoutBatchImportTest` net8.0/net6.0 各 76/76（`forward-annotated-npoi-net8.trx`、`forward-annotated-npoi-net6.trx`）；ClosedXML net8.0/net6.0 各 58/58（`forward-annotated-closedxml-net8.trx`、`forward-annotated-closedxml-net6.trx`）。测试覆盖普通与模板同步/异步、错误 Sheet、业务 Sheet、错误行、原始单元格批注、未知模式拒绝、模板流和调用方流保持打开；此前 `ErrorRowsOnly` 矩阵仍为 NPOI 各 72/72、ClosedXML 各 54/54。
- 文件提交职责测试：NPOI `EntityFailureWorkbookCommitterTest` net8.0/net6.0 各 27/27，ClosedXML net8.0/net6.0 各 17/17；`ErrorRowsOnly` 与 `AnnotatedOriginal` 的 Forward 带路径输出均验证同步只调用 `Commit`、异步只调用 `CommitAsync`，提交失败只观察一次、无错误时不调用提交器并保留原目标。证据为 `forward-annotated-committer-npoi-net8.trx`、`forward-annotated-committer-npoi-net6.trx`、`forward-annotated-committer-closedxml-net8.trx`、`forward-annotated-committer-closedxml-net6.trx`。
- 本次仍不宣称 HSSF Forward、复杂 Entity Layout 或完整低保留失败产物能力；这些仍需单独设计原始工作簿保真、模板/多区域布局和资源矩阵后再开放。

## 2026-10-10 后续：ClosedXML 模板预检 XML 资源预算

- `ExcelPivotTemplateParts.Find` 与 `ReadMetadata` 新增接收资源限制、Provider 和操作上下文的内部重载，逐 XML 节点累计名称、值和属性字符数，并检查嵌套深度；超限抛出带 Provider、操作和 Preflight 阶段的 `BingOfficesResourceLimitException`。旧内部入口继续不设置显式 XML 字符数和深度预算。
- ClosedXML Workbook 导出、实体模板导入、失败工作簿输入预检和 Forward 实体模板导入均传入调用方已有的 `ExcelResourceLimits`。实体导出没有资源限制选项，继续使用默认预算。新增测试覆盖字符数、深度、公开导出异常上下文、模板位置及目标流保全。
- `production-test-map.md` 已补资源预算符号和测试映射，并注明导入与 Forward 传播目前缺少直接用例。
- Release 验证受现有 NuGet 资产阻塞：ClosedXML 与测试项目 `dotnet build --no-restore` 均因 `NETSDK1064` 找不到 `Microsoft.Extensions.Logging.Abstractions 6.0.0` 而未进入源码编译；测试项目资产仍指向不可访问的 `C:\Users\CodexSandboxOffline\.nuget\packages`。本机 NuGet 缓存的离线还原未能生成新资产。`dotnet test` 返回的“生成成功”没有发现或运行测试，不计作通过。
- 本轮改动文件通过严格 UTF-8、BOM、LF、末尾换行检查；`git diff --check` 退出 0。Git 同时报告两个未修改文件存在既有行尾转换提示：`Bing.Offices.ClosedXml.csproj` 与 `Bing.Offices.ProfileFixtures.xml`，未调整这些文件。
- 当前代码和测试仍待在具备有效双 TFM NuGet 资产的环境中编译与运行；本次不把 PTPL-001 阶段状态升级为 PASS，也不更新 API 基线或门禁配置。

## 2026-10-10 后续：XML 资源预算验证补齐

- 使用本机 NuGet 缓存完成 ClosedXML 与测试项目还原；Release 构建通过。Abstractions 下存在未跟踪的旧生成输出，普通 glob 会重复包含程序集属性，本次构建仅通过命令行 `DefaultItemExcludesInProjectFolder=output/**` 排除此输出，没有删除或改写该目录。
- `ExcelPivotTemplatePartsTest` 的 `MetadataResourceLimits` 与 `TemplateXmlCharacterLimit` 在 net8.0、net6.0 各 2/2 通过；net8.0 `Category!=Large` 常规筛选 765/765 通过。
- net6.0 `Category!=Large` 全量筛选两次各出现一个不同的共享临时目录快照竞态失败；对应资源用例单独运行 2/2 通过。设置单核环境的再次运行在测试宿主无输出且 CPU 不再增长后中断，因此 net6.0 全量筛选仍未验证通过。
- XML 资源预算在导出、实体模板导入、Forward 模板批次导入及失败工作簿输入预检中的同步/异步职责用例，在 net8.0、net6.0 各 8/8 通过；覆盖资源错误上下文、输入未读取、无批次交付、根关系元数据预算及目标流保全。模板部件资源与失败工作簿预检组合筛选在两个 TFM 各 29/29 通过。PTPL-001 保持 PARTIAL，完整阶段仍有其它清单与依赖图验收；未更新 API 基线、CI 配置或测试矩阵。

## 2026-10-10 后续：两 Provider 关系 XML 资源预算

- `ExcelPivotTemplateManifest.Read(Stream, limits, provider)` 现在将限制和上下文贯穿部件发现及 `ExcelPivotTemplatePackage` 关系/结构 XML 读取；旧无资源限制入口保留无界字符/深度行为。字符数和深度超限均保留 Provider 与预检异常上下文。
- 共享职责测试对 NPOI、ClosedXML 的 net6.0/net8.0 分别覆盖一般 ZIP/XML 限制、超长根 `.rels` 字符数及超深根 `.rels`，各 10/10 通过；`TemplatePivot` Provider Contract 在 net6.0、net8.0 各 82/82 通过。ClosedXML 普通导出、实体模板导入、Forward 批次及失败工作簿输入预检预算直接用例仍各 TFM 8/8 通过。
- PTPL-001 标记 PASS。PTPL-002 仍为 PASS（子集）；PTPL-003 仍因实际客户端刷新/首开/保存重开验证未完成而保持 PARTIAL。未更新 API 基线、CI 配置或测试矩阵。
- 本次更新文档均为 UTF-8 无 BOM、LF、末尾单换行；`git diff --check` 通过。

## 2026-10-10 后续：Excel 客户端会话复验

- Excel 安装存在于 `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`；本轮开始时没有既存 Excel 进程。NPOI/ClosedXML 同步和异步模板更新夹具共四份已生成于系统临时目录 `BingOfficesPivotClient-d37592d55b0446e4a8ae88bdf9608aee`，未写入仓库。
- PowerShell `New-Object -ComObject Excel.Application` 返回 `0x80080005 (CO_E_SERVER_EXEC_FAILURE)`。显式 `/x` 与 `/x /automation` 启动均未形成可自动化的桌面会话：进程无主窗口且不在 ROT 注册；后一实例在启动等待期后自行退出。四份夹具均未交给 Excel 打开，故无首次打开、源数据修改后刷新、保存重开的客户端结果。
- 自动化尝试期间未发现或连接用户既有 Excel 进程；当前确认没有遗留 Excel 进程。PTPL-003 保持 PARTIAL，PTPL-004 仍依赖 PTPL-003。此次环境复验不标记客户端验收通过，也不以结构测试替代。
- 本轮只追加环境诊断记录；没有修改生产代码、测试、CI、API 基线或任务计划。本记录为 UTF-8 无 BOM、LF、末尾单换行。

## 2026-10-10 后续：Windows PowerShell STA 会话复验

- Codex 桌面自动化返回 `apps: []`，没有可控制的 Excel 桌面应用。当前会话普通启动 Excel 后，进程保持运行但没有主窗口；Windows PowerShell 5.1 `-STA` 的 COM 激活超过 25 秒仍未返回，期间产生的无窗口 Excel 进程已结束。
- 此次未打开模板夹具，没有执行源数据变更刷新、保存或重开；未产生客户端验收产物。PTPL-003 继续保持 PARTIAL，阻塞类型为外部桌面会话不可用；PTPL-004 仍依赖 PTPL-003。
- 没有修改生产代码、测试、CI、API 基线或任务计划；PowerShell COM 探测未建立客户端会话。

## 2026-10-10 后续：无业务 Sheet 的模板透视更新

- 修正 `ExcelPivotTemplatePreserver.IsPreserveOnly`：请求含 `PivotTableUpdates` 时不再进入原样复制捷径，继续经过模板桥接和 Provider 更新缓存路径。修复了仅调用 `UseTemplate(...).UpdatePivotTable(...)`、没有 `AddSheet` 时更新定义被忽略的问题；符合计划中显式保留/更新模式允许不添加业务 Sheet 的约定。
- `TemplatePivotPreserveExportTest.UpdateOnly_ShouldRebuildTemplatePivotWithoutBusinessSheet` 覆盖 NPOI 与 ClosedXML 的同步、异步输出，断言原模板不被原样复制、缓存记录完整、缓存要求打开时刷新、目标透视保留且没有临时 Sheet。net8.0 4/4、net6.0 4/4 通过。
- 按当前用户范围，Excel 客户端启动、刷新显示及保存重开不作为本轮库功能开发的前置或验收门槛；本轮只依据 Provider 文件结构与缓存合同判断，不宣称客户端行为已验证。
- PTPL-003 仍有计划内实现待完成：更新时保留未显式覆盖的透视布局/外观；目标区域冲突预检与旧正文处置；来源收缩时清除受管理范围内的旧数据。新建与更新组合请求仍在 Provider 入口拒绝。PTPL-004 的完整业务组合及失败/重试矩阵、PTPL-005 的候选 API/消费者/常规门禁收口也未完成。
- 本轮修改 `src/provider-shared/ExcelPivotTemplatePreserver.cs` 与 `tests/Bing.Offices.ProviderContract.Tests/Contracts/TemplatePivotPreserveExportTest.cs`；未改 CI、项目配置、API 基线或版本，未提交、推送或发布。

## 2026-10-10 后续：更新目标占用与旧正文

- `ExcelPivotTemplateBridge.PrepareUpdatedTargets` 现在解析原模板及 Provider 输出透视的 A1 `location ref`，拒绝与未更新透视、另一项更新或声明目标范围外非空单元格重叠；确认所有目标通过后，清除原透视声明范围中的值、公式及内联文本，同时保留单元格样式。
- 新增 ProviderContract `Update_ShouldClearOldPivotBodyAndKeepOutsideCells` 与 `Update_ShouldRejectOccupiedTargetAndKeepDestination`，覆盖 NPOI/ClosedXML 同步和异步、更新锚点冲突、旧单元格清理、区域外内容保留、结构化 Preflight 错误及目标流不变。完整 `TemplatePivotPreserveExportTest` net8.0 与 net6.0 均为 60/60 通过。
- 冲突/清理范围以 OOXML `location ref` 声明为准。当前 Provider 生成的引用可能只有目标起始单元格，因此不代表 Excel 客户端刷新后透视表的完整渲染占地；客户端显示和保存重开仍不作为本轮库功能门槛，也没有宣称已验证。
- 首次常规测试构建发现已有未跟踪 `src/Bing.Offices.Abstractions/output/summary-forward-net8-obj` 中生成的程序集属性被 SDK 编译，产生重复属性错误；未清理或改写该目录。通过临时 MSBuild 排除现存 `output/**` 后构建及测试通过。未修改 CI、项目配置、API 基线或版本。

## 2026-10-10 后续：来源范围收缩清理

- `ExcelPivotTemplateBridge.PrepareShrunkSources` 在全部或部分更新合并前，比较模板缓存中原有的字面工作表来源范围与请求的新范围。仅当来源工作表和起点相同、且新范围是旧范围的严格子集时，清除旧范围中新范围之外的单元格值、公式与内联文本；保留单元格样式及旧范围外内容。
- 若保留透视或同批其他更新仍引用将被清理区域，则跳过清理，以免破坏其来源数据。命名来源、不同工作表、起点变化、扩张或不相含区域不做推断式清理。
- 新增 `UpdateSourceShrink_ShouldClearOldManagedRows`，使用模板原来源进行无业务 Sheet 更新，覆盖模板原始尾部数据、范围外数据和重建缓存记录。NPOI/ClosedXML 同步/异步 net8.0、net6.0 共 8/8；完整 `TemplatePivotPreserveExportTest` net8.0 与 net6.0 各 64/64。
- 本轮修改共享桥接器、ProviderContract 测试及生产符号追溯表。首次构建因 XML/obj 输出访问拒绝未运行测试；使用已有工作区授权完成后续构建。一次早期测试断言对整个缓存 XML 搜索数值，造成与清理行为无关的误报；已改为断言来源工作表范围及缓存记录结构，最终精准及合同测试通过。
- PTPL-003 仍待更新时保留未显式覆盖的透视布局/外观，以及支持“新建与更新组合请求”；后者目前仍在 Provider 更新入口拒绝。PTPL-004 的业务/失败重试组合尚未收口，外部 Excel 客户端验收按当前用户范围不作为本轮库实现前置；PTPL-005 的候选 API/消费者/常规门禁及独立 Review 未完成。
- 未修改 CI、项目配置、API 基线或版本；未提交、推送或发布。验证命令临时设置 `GenerateDocumentationFile=false`、排除现存 `output/**`，没有改动这些输出或配置。

## 2026-10-10 后续：布局保留、来源收缩与组合更新

- `ExcelPivotTemplateBridge.PreservePivotLayout` 在更新输出中恢复模板未由更新定义描述的根级显示/格式属性、`formats`、`conditionalFormats`、`chartFormats` 和 `pivotTableStyleInfo`。字段索引格式仅在新旧缓存字段顺序一致时复制；顺序改变时结构化拒绝，避免带错索引的格式继续输出。轴、度量、筛选、缓存和目标仍采用本次更新生成结果。
- `PrepareShrunkSources` 仅对同工作表、同一起点、严格缩小的有界字面来源范围清理旧单元格值、公式和内联文本；新来源范围、单元格样式与范围外内容保持。保留透视、同批更新或新建透视仍引用待清理范围时跳过清理。
- `ExcelPivotTemplatePlan` 保存本次新建透视描述；NPOI/ClosedXML Provider 将新建和模板更新定义一起交给现有 Pivot Writer。桥接器按请求名称区分两类生成部件，更新缓存只对涉及模板更新的缓存设置 `refreshOnLoad`；部分更新继续恢复未更新透视及原缓存。更新入口仍精确定位已有模板透视，不隐式新建替代。
- 修正文档注释，说明未指定更新的模板透视保持原样。新增组合覆盖全部旧透视更新和部分更新并保留旧透视两种模式，NPOI/ClosedXML 同步/异步 net6.0、net8.0 共 16/16。布局保留职责用例双 TFM 共 8/8；来源收缩职责用例共 8/8；最终 `TemplatePivotPreserveExportTest` 双 TFM 各 76/76。
- 按用户确定的当前范围，本轮不启动 Excel/WPS/LibreOffice；刷新、首开显示及保存重开外部验收不作为源码实现阻塞。这里的测试只证明生成工作簿结构与缓存合同。
- 本轮改动共享透视桥接、模板执行计划、模板保留判断、NPOI/ClosedXML 报表导出、ProviderContract 测试及本任务执行/追溯文档。首次普通构建遇到现有 obj/XML 输出访问拒绝，未清理/终止既有进程或输出；之后使用当前工作区授权完成构建与测试。命令行只临时设置 `GenerateDocumentationFile=false` 并排除既有 `output/**`，没有改 CI、项目配置或基线。
- 当前 PTPL-003 源码实现待办中的布局/外观保留、目标冲突和旧正文处理、来源收缩、组合新增与更新均已实现并通过合同回归。PTPL-004 的业务组合及失败/取消/重试矩阵、PTPL-005 的当前候选 API/包消费者/常规门禁与独立 Review 仍待完成；客户端证据单列为当前用户范围之外。
- 未修改版本号、API snapshot、CI、项目配置；未提交、推送或发布。

## 2026-10-10 后续：OOXML 布局属性名称核对

- 对照 Open XML `pivotTableDefinition` 属性定义，修正布局白名单中的 `showEmptyCol`、`showDataTips` 拼写，并移除误列为根属性的样式子元素属性；`pivotTableStyleInfo` 仍整体保留。
- `Update_ShouldPreserveUnconfiguredPivotLayout` 夹具明确写入并断言 `showEmptyCol` 与 `showDataTips`。最终属性白名单修正后，net8.0、net6.0 各 4/4 通过，覆盖 NPOI/ClosedXML 同步和异步；完整合同类在本次名称核对前各为 76/76。
- 本轮不运行 Excel 客户端；测试仅验证生成包中的 XML 布局属性及样式结构。
- 修改共享桥接器、ProviderContract 测试和生产符号追溯表；未修改 API、CI、基线或项目配置。

## 2026-10-10 后续：PTPL-004 库侧业务合同

- 增加采购/仓储双 Sheet 模板合同：基础 Sum、多轴五聚合、状态筛选和自定义数字格式在同一请求中更新两个独立缓存；断言新源记录、轴字段、聚合顺序、筛选字段、格式和目标透视结构。NPOI/ClosedXML 同步/异步 net6.0、net8.0 各 4/4。
- 增加同一更新 Request 连续成功导出两次、来源从三行扩张到四行、首次目标流写入失败后复用同一 Request 重试，以及未知 `extLst` 在既有 `ExcelPivotTemplatePlan.Create` 预检中结构化拒绝的直接合同；分别核对缓存记录、目标流和模板输入完整性。新增用例双 Provider/同步异步/双 TFM 各 4/4。
- 来源范围收缩、共享缓存拆分、已知布局保留、更新目标冲突和失败后重试均已有对应直接用例。最终 `TemplatePivotPreserveExportTest` net8.0、net6.0 各 96/96 通过，包含本轮所有新增测试。
- 本轮 PTPL-004 的库侧业务与失败合同已完成；实际 Excel/WPS/LibreOffice 打开、刷新、保存重开按当前用户范围排除，不据此宣称客户端行为通过。PTPL-005 的 API 成员审批/双 TFM 差异、NuGet 消费者、Release/Category!=Large、Docs 及独立 Review 仍未收口。
- 修改生产符号追溯表和本任务执行记录；未增加 CI job/matrix、API 基线或项目配置，未修改版本号，也未提交、推送或发布。

## 2026-10-10 后续：PTPL-005 包消费者与文档验证

- 新增共享包消费者合同 `AdvancedPivotPackageContract.VerifyTemplateUpdateAsync`，仅通过公开 Workbook API 及已打包 NuGet 引用运行；普通 .NET 6、.NET 8 消费者和第三方 Provider 的公开 API 消费者都调用 NPOI、ClosedXML 两个 Provider。合同覆盖同步/异步更新、模板流保持打开且字节不变、透视名称、缓存字段、全部缓存行值及 `refreshOnLoad`。最终输出记录为 `package-consumer-ok`（.NET 6.0.36、.NET 8.0.31）和 `third-party-public-only-provider-ok`；消费者没有引用内部类型或生产友元。
- 八个 Provider 包以本任务候选版本 `2.0.0-pivottpl-20261010` 打包；普通和第三方消费者均从候选包编译、运行通过。打包保留既有目标框架与缓存警告，没有新增依赖、测试矩阵或 CI 配置。
- 文档恢复基础透视历史客户端夹具的完整数值证据，并明确该结果不覆盖模板更新；新增公开模板更新示例与消费者验证映射。Docs 项目 `--no-build --list-tests` 没有列出测试用例，不能计为通过；当前源码的 Release 构建因离线缓存缺少 `Microsoft.CodeAnalysis.Analyzers 3.3.4`（`NETSDK1064`）失败，因此本轮 Docs 示例测试未验证。
- 已捕获 net6.0 与 net8.0 API 候选。Abstractions 相对正式基线各有 27 个新增、0 个删除成员；经成员级核对，7 个成员属于本任务 `api-approval.md` 批准范围，其余 20 个来自当前工作树中其它实体批次/映射改动。候选身份仍绑定前序任务，比较器无法将其作为本任务干净候选验收；未更新正式 API 基线。需在隔离且只含已批准 API 差异的候选状态下重新比较。
- 全解决方案 Release 门禁未完成：无还原构建受现存 `output/summary-forward-net8-obj` 生成源码重复包含及离线 NuGet 缓存缺少 `Microsoft.CodeAnalysis.Analyzers 3.3.4` 阻塞。没有删除或改写该输出，也没有更换环境包源。受影响 Provider/测试项目包构建及 .NET 6/.NET 8 消费者验证通过，`TemplatePivotPreserveExportTest` 双 TFM 各 96/96；这不等价于完整解决方案门禁通过。
- 独立 Review 仍未完成。仓库 `review-code` 技能解析结果为 `CLI_NOT_FOUND` / review route unavailable，当前没有可运行的独立只读审查会话；因此没有创建或伪造 Review 通过结果。应在可用的独立审查会话中对照本计划、源码和验证证据完成审查。
- 本轮实际客户端范围仍排除 Excel/WPS/LibreOffice；此前基础透视客户端证据只保留为历史事实，不外推至模板更新。没有新增公共 API，未修改正式基线、CI、项目配置或版本，也未提交、推送或发布。
- 交付格式检查：任务涉及的 C# 文件均为严格 UTF-8 BOM、LF、单个末尾换行；Markdown 为严格 UTF-8 无 BOM、LF、单个末尾换行。`git diff --check` 退出码为 0；Git 另提示当前工作区修改的 `Bing.Offices.ClosedXml.csproj` 将在下次 Git 写入时转换为 CRLF，该文件不在本任务修改范围内。

## 2026-10-10 后续：PTPL-005 Release 与常规门禁复验

- Docs 示例适配已完成；全解决方案 Release 构建编译了 Docs 测试程序集。`dotnet test Bing.Offices.sln -c Release --no-build --no-restore -m:1 --filter 'Category!=Large'` 中 `DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` 所在测试程序集通过 `19/19`，新增的透视模板更新示例也随全部文档代码围栏编译并执行。
- 全解决方案 Release 构建命令 `dotnet build Bing.Offices.sln -c Release --no-restore -m:1 -p:UseSharedCompilation=false -p:DefaultItemExcludesInProjectFolder=output/**` 通过，`0` 错误、`24` 个警告。临时排除仅针对工作区已有的 `output/**` 生成文件，没有删除或修改输出、项目配置或 CI。
- 上述 `Category!=Large` 全量测试除公开 API snapshot 外通过。Integration 的 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 在 net6.0 和 net8.0 均失败：Abstractions 基线为 `1455` 个成员，当前为 `1480` 个，差异 `+25/-0`。本任务 `api-approval.md` 仅批准其中七个透视模板成员；其余差异来自当前工作树其他功能。没有更新正式 snapshot，也没有把 API 门禁记为通过。
- PTPL-005 中 NuGet 普通/第三方公开 API 消费者和 Docs 示例现已验证；Release 编译门禁通过。剩余工作是将 API 差异与其它工作树变更隔离并完成成员级核对/候选身份绑定，再通过双 TFM API 门禁；独立 Review 仍因当前 Review 执行器不可用而未完成。Excel/WPS/LibreOffice 客户端验收按用户确认的范围排除，不作为本轮代码功能阻塞，也不声称客户端刷新已验证。
- 本次复验未新增 CI workflow/job/matrix、基线快照或依赖；未修改版本，未提交、推送或发布。`Category!=Large` 全量命令退出码为 `1`，唯一已观察到的失败为上述双 TFM API snapshot 不匹配。

## 2026-10-10 后续：PTPL-005 候选 API 身份重绑

- 在当前 Release 输出与源码状态上重新生成候选，保留在 `artifacts/api-candidate-pivottpl-20261010/api-candidate-rebound.json`、`api-snapshot-net6-rebound.json`、`api-snapshot-net8-rebound.json` 和 `api-diff-rebound.json`。新候选源码清单 SHA-256 为 `6F5E9484A07565232D92B04DF30EBD8FEC0AC38414D20D68E3902EEB45EA45B6`；包含 14 个程序集身份和 8 个生产包身份。
- 八个 `2.0.0-pivottpl-20261010` 包从本轮 Release 输出使用 `dotnet pack --no-build --no-restore` 重新打包到系统临时目录。逐个检查包内 `lib/` 的 14 个 DLL 与 `output/release` 对应文件，SHA-256 全部一致；新旧候选包规范化身份 8/8 一致，已归档候选包内 DLL 也与当前 Release 输出 14/14 字节一致。因此既有普通 .NET 6/.NET 8 与第三方公开 API 消费者结果对应同一包载荷，无需重跑相同消费者。
- 对新候选双 TFM 成员清单重新计算差异：Abstractions 各新增 27、删除 0。成员核对结果为：本任务已批准 7 个；ExcelDataReader 失败工作簿任务已批准 1 个；Entity Forward Batch 任务的 15 个成员仍标记待审批；Entity Failure Workbook 任务的 2 个成员仍标记待审批；两个 `TargetPropertyName` 成员在任务审批记录中未找到批准项。正式 API 基线及审批文件均未修改。
- PTPL-005 的 API 门禁仍受 19 个未批准/缺少批准记录的成员阻塞；全量 `Category!=Large` 仍不能记为通过。独立只读 Review 当前不可用：现有 Review 工具仅提供 PR 检查诊断，仓库 Review 路由仍返回 `CLI_NOT_FOUND`，没有独立审查结果。Excel/WPS/LibreOffice 客户端行为继续按用户范围排除。
- 本轮未修改生产代码、测试、CI、版本或正式基线；仅新增上述候选快照/差异及本执行记录。未提交、推送或发布。

## 2026-10-10 后续：模板关联透视功能边界

- 补齐 P3/D-005 的关联功能边界识别。`ExcelPivotTemplateManifest` 沿工作簿、工作表、Drawing 和 Chart 的实际关系链识别带 `pivotSource` 的透视图，并识别工作簿关系中的透视切片器缓存与时间线缓存；不扫描无关包关系。
- 更新计划在数据枚举前以具体功能名称拒绝当前不支持的透视图、切片器或时间线更新。显式保留模式继续复制关系与部件原始字节；普通图表和未关联透视表的表格切片器不触发此拒绝。
- `TemplatePivotDependencyBoundaryTest` 在 net8.0、net6.0 各 40/40 通过，覆盖 NPOI/ClosedXML 同步与异步更新拒绝、目标与输入流保全、三个关联功能在保留输出中的关系和部件保留，以及普通图表和表格切片器不误触发拒绝。
- 更新透视文档和生产符号到测试方法追溯表。该增量只修改内部预检、合同测试和文档，没有新增公共成员；不影响双 TFM API 差异或审批状态，也未更新正式 API 基线、CI、项目配置或版本。
- `TemplatePivotDependencyBoundaryTest` 的普通功能放行检查在调整为 XML 结构比较后，net8.0/net6.0 均为 40/40 通过；NPOI 重写 XML 命名空间前缀的序列化差异不再误判为图表结构丢失。
- P3/D-005 标记为完成。PTPL-005 仍因 19 个未批准/缺少记录的 API 成员及独立 Review 不可用而未收口；Excel 客户端验证继续按用户范围排除。

## 2026-10-10 后续：D-007 库侧语义与 D-009 能力矩阵

- 透视文档新增 Provider/格式矩阵，分别说明基础与高级新建、模板保留和模板更新范围；另列保留、更新刷新标志、公式计算、客户端刷新/呈现及数据权限语义。
- D-007 的库侧边界在文档中明确：保留不刷新来源缓存；更新重建缓存并设置 `refreshOnLoad=true`，Provider 不执行客户端刷新或公式计算。更新旧正文和来源收缩清理由既有桥接实现负责。
- D-009 完成库侧文档部分：筛选、隐藏工作表及保护不能替代数据权限，完整来源仍在工作簿与缓存中。NPOI/ClosedXML XLSX 支持模板保留和更新；其他 Provider/格式不支持该模式。
- 本次仅更新文档与执行记录，没有公共 API、生产代码或测试代码变化。Excel/WPS/LibreOffice 模板客户端显示仍按用户确认的范围排除；不据此宣称客户端结果通过。
- PTPL-005 仍受 19 个未批准或缺少成员审批记录的 API 差异，以及当前没有独立 Review 执行器阻塞；未修改正式 API 基线、CI、依赖或版本。

## Review 修复记录

### Round 1

- 初审状态：`NEEDS_FIX`；修复后独立复审：`PASS_WITH_ISSUES`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001/review.md`

#### FIX-001

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/provider-shared/ExcelPivotTemplatePreserver.cs`
  - `src/provider-shared/ExcelPivotTemplateBridge.cs`
  - `tests/Bing.Offices.ProviderContract.Tests/Contracts/TemplatePivotPreserveExportTest.cs`
- 根因：模板保留预检把“仅保留已有透视并新建透视、没有模板更新”的合法请求误判为不支持；桥接无更新路径也未将新建部件与原模板部件做部件名和缓存标识协调。
- 修复：允许有 XLSX 模板的显式保留请求进入主链；合并时为生成部件分配不冲突的部件名和关系，重映射缓存 ID，并将缓存声明合并到唯一的 `pivotCaches` 集合。保留原有透视缓存和关系，同时加入新透视。
- 验证：
  - `PreserveAndCreate_` 职责用例：NPOI/ClosedXML、同步/异步，net6.0 `12/12`、net8.0 `12/12`。
  - `FullyQualifiedName~TemplatePivot`：net6.0 `168/168`、net8.0 `168/168`。
  - NPOI 与 ClosedXML Release Provider 双 TFM 编译及 ProviderContract 测试项目双 TFM 编译通过，0 个错误。

#### FIX-002

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
  - `artifacts/packages-pivottpl-reviewfix-20261010/`（八个候选 NuGet 包）
  - `artifacts/api-candidate-pivottpl-reviewfix-20261010/`（候选身份及双 TFM 快照）
- 根因：先前候选生成后继续修改了模板 Manifest，旧包消费者结果无法证明当前最终源码对应的包载荷。
- 修复：从本轮 Provider 编译产物重新打包 NPOI、ClosedXML，其余六个包沿用相同候选版本的既有包；按当前源码和八个包重新捕获候选身份及双 TFM API 快照。包内 14 个程序集与候选程序集逐字节匹配，两个本轮更新 Provider 的包也与双 TFM 编译产物匹配；消费者隔离缓存中的八个 `.nupkg` 均与候选 feed 字节一致。
- 候选源码清单 SHA-256：`739E38686F2CAD0023D2FD36546FEE5F3A9E524423EFCE6E338A7355381795B9`。正式基线相对差异仍为 net6.0 `+27/-0`、net8.0 `+27/-0`；仅本任务七个成员已批准，其余成员的审批阻塞不因本轮包身份重绑而改变。
- 验证：
  - 普通 NuGet 消费者：.NET 6 `package-consumer-ok`，.NET 8 `package-consumer-ok`。
  - 第三方公开 API 消费者：net6.0、net8.0 均为 `third-party-public-only-provider-ok`；未引用 Provider 内部类型。
  - 候选快照：`artifacts/api-candidate-pivottpl-reviewfix-20261010/api-candidate.json` 及两个 TFM 快照；未更新正式 API baseline。

### Round 1 汇总

- MUST_FIX：`0`
- 已完成：`FIX-001`、`FIX-002`
- PARTIAL：`0`
- BLOCKED：`0`（本轮 Review Fix 项）
- FAILED：`0`
- 回归验证：模板透视合同双 TFM、Provider 编译、普通与第三方 NuGet 消费者通过。
- 仍独立跟踪：正式 API snapshot 门禁因其它 19 个未批准或缺少审批记录的成员保持 `BLOCKED_APPROVAL`；不属于本轮 FIX，未修改基线。Excel/WPS/LibreOffice 客户端行为按用户确认范围排除。
- 独立复审：`ai_docs/tasks/BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001/review.md` 已关闭 FIX-001、FIX-002，`OPEN_ACTIONABLE=0`；正式 API 门禁仍为 `BLOCKED_APPROVAL`。复审期间额外并行测试发生共享临时目录竞争，重编译发生共享 obj 文件锁定；不替代也不推翻本轮串行双 TFM 验证。
- Review Fix 终态：`COMPLETED` 仅表示本轮 recommended 范围内的 FIX 已处理并完成独立复审，不代表正式 API 审批门禁或整个更大 Goal 完成。
- 后续：维护者处理正式 API snapshot 中 19 个非本任务成员的审批/隔离；不扩展本任务的客户端验收范围。
