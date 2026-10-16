<!-- AI_EXECUTION_STATUS: IN_PROGRESS -->
AI_TASK_ID: BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001
AI_EXECUTION_STARTED_AT: 2026-10-10T19:10:56.710Z

# 实施执行报告

## 当前结论与完整范围

IN_PROGRESS。用户持续授权按 TODO 推进。NPOI XLSX 基础及高级透视已进入公开同步/异步导出，保存完整来源缓存和原生字段；XLS/HSSF 透视继续提前拒绝。当前候选 Bing.Offices.sln Release 隔离构建通过，0 错误、145 条警告。`Category!=Large` 常规测试退出码为 1：NPOI.Tests 双 TFM 各 1745/1745、ProviderContract.Tests 各 679/679、Docs.Tests 19/19 通过；ClosedXML.Tests 双 TFM 各有 3 项失败，涉及前向批次暂存清理断言及 ErrorRowsOnly 列元数据，属于 F-006/F-007 范围。提权测试没有复现 `File.Replace` 沙箱异常。当前候选 API 快照复核中，NPOI、ClosedXML、Core 等 Provider 均为零成员差异；Abstractions 相对 F-003 批准基线新增 27 条 API 记录（25 个成员、2 个类型），与本阶段零公共成员批准不符，未更新基线。高级文本 Count 的 Excel 数值仍未通过；按用户的非 Office 客户端范围，本轮未启动任何 Office 客户端。F-003 仍未收口，F-004～007 和 E-001/E-002 不缩小、不宣告完成。

以上早期测试、包和客户端结果仅属于此前冻结的 F3 候选。后续候选及 F4 等阶段必须以各自批准的 API、包和门禁证据验收，不得借用旧结果替代。

生产 Provider 已声明 PivotTables/AdvancedPivotTables，明确仅支持 XLSX 新建；结构及来源公开重开不能替代客户端数值验收。以下旧验证与修复记录保留历史时点事实，当前进度以本节、阶段矩阵及新增验证记录为准。

## 阶段矩阵

| 阶段 | 当前状态 | 内容 |
| --- | --- | --- |
| NPIVOT-001 来源及容量预检 | PARTIAL | 已接入主链；类型、坐标、日期、筛选和预算职责通过；剩余客户端实际布局验收 |
| NPIVOT-002 完整缓存及原生字段 | PARTIAL | 完整共享项/记录/关系、字段角色、五聚合和索引零筛选通过；多透视隔离 net8 通过，客户端待验 |
| NPIVOT-003 主链与资源 | PARTIAL | 公开导出、普通模板及同步/异步文件资源通过；排除 SXSSF，既有 100K 定向探针通过；当前候选常规矩阵未通过，ClosedXML 下游失败见本轮记录，客户端待验 |
| NPIVOT-004 业务与客户端 | PARTIAL | 当前 NPOI.Tests 双 TFM 各 1745/1745、共享 Provider Contract 双 TFM 各 679/679；客户端高级 Count 未通过，完整客户端矩阵未执行 |
| NPIVOT-005 API、包与交付 | IN_PROGRESS | 当前 Release 构建通过；常规门禁因 ClosedXML 下游测试失败，Abstractions 快照有 27 条新增 API 记录；当前候选包消费者及最终独立 Review 未闭环 |

## 原始证据及计划局部调整

原 plan 的 P2 路线是在 NPOI 公开 CT schema 中补齐 sharedItems/records 再序列化。实际 roundtrip-cache-04.json/原 XLSX 显示 cacheFields=6、recordCount=3，但 records 部件实际 0 条，CompleteCacheRoundTrip=false。退出 0 只是诊断执行完成。run-05 后续因探针读取重开对象的 NullReferenceException 终止，未形成新的完整结果；其实际退出码 3221225477 和旧产物保持原样。

实际调整：NPOI 公开接口建立原生透视、缓存定义与 OPC 关系；新增 NpoiPivotCacheWriter 建立类型化共享项与 NpoiPivotCacheRecords 不可变索引快照；NpoiPivotCacheRecordWriter 使用 XmlWriter 写出全部 r/x 记录。在 NpoiSheetContentWriter 的既有磁盘暂存与 ZIP 转写中替换实际记录部件。旧的内联列表引号修正与新增缓存共用一次暂存，不加载业务数据行 XML，不使用反射、生产友元或复制引擎内部实现；不计算静态汇总或裁剪来源。

风险：合法共享项/记录不能单独证明 pivot field.items、轴/度量/筛选以及首次显示正确；后续仍需专用 TableWriter、包关系、完整业务及客户端矩阵。现有透视模板仍按 F-004 边界，不在此阶段假装保真。

## 修改及职责

主代理拥有三个生产内部单类型文件及 NpoiSheetContentWriter 的最小重载/写入接入：

- Exports/PivotTables/NpoiPivotCacheWriter.cs：先构建和验证临时 cacheFields，再替换定义；非法输入或取消不留下半份缓存。全部来源类型保留，真空与空文本分开。
- Exports/PivotTables/NpoiPivotCacheRecords.cs：校验完整字段宽度与索引范围，隔离调用方行数组和共享项数量集合。
- Exports/PivotTables/NpoiPivotCacheRecordWriter.cs：无 BOM UTF-8 记录 XML，逐行/字段检查取消，调用方流保持开启。
- Exports/NpoiSheetContentWriter.cs：原重载继续兼容；新增缓存重载仅 XLSX，预先核对部件存在，再向目标 ZIP 写入。未覆盖部件保留原内容，临时文件 DeleteOnClose。

Luna 独占新增 NPOI/PivotTables 直接测试及其必要单类型辅助文件。所有 dotnet 命令主代理串行运行，避免并发构建竞争。

## 验证记录

- cache-production-build-01.command.json：三个缓存生产类型的双 TFM Release 构建退出 0，0 错误/0 警告。该结果在 NpoiSheetContentWriter 新重载加入之前，不作为新重载已编译通过的证据。
- 新直接职责及最终编译尚待执行，不将测试文件存在视为通过。
- SharedItems 类型标志依照微软公开文档：containsMixedTypes 的单类型字段仍可含空白；containsSemiMixedTypes 表示含文本。参考：[SharedItems](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shareditems?view=openxml-3.0.1)。将实际 XML 与完整值矩阵验证，不只断言属性赋值。

## 任务状态与环境

首次 task-state start 在沙箱中写 .agents/runtime 临时文件遇 EPERM，execution 启动元数据已写入；随后仅该注册命令通过授权升级正常执行。实际 runtime 当前 active=true、taskId=F-003、mode=plan-execution、status=implementing。原运行游标另存 artifacts/runtime-before-f003.json；未修改 task-state 工具或发送通知。

正常 Office 验证此前的自动审批超时/沙箱 WinError 1312 与 Count 失败仍单列。本轮未新启动 Office，待确认不作为授权或成功。

## API、编码、回归与交付

全部新类型 internal，公共接口/旧签名/能力位暂不修改；双 TFM 最终 API 成员差异预计为零，但 SourceManifest/程序集/包身份因实现发生变化，旧 API 身份比较结果不能当作当前新候选门禁。最终捕获、审批内身份绑定和 NuGet 消费者尚待执行，不覆盖基线掩盖差异。

C# 使用 UTF-8 BOM/LF/单末尾换行，Markdown/JSON 无 BOM/LF；保留前序大量未提交改动，未全仓库格式化。后续必须完成共享序列化默认实现的独立 XLS/XLSX 成功/失败、完整 XML/工作簿结构、取消/提交/临时文件/流所有权，以及原内联校验回归。

未自动 commit、push、创建 PR、发布、修改版本或发送外部通知；未执行 task-finish，完整 Goal 继续进行中。

## Review 修复记录：Round 1

Review 状态 NEEDS_FIX，fixScope=recommended。review.md 保留为检查时点证据，不修改其结论。2026-10-05T08:51:15.908Z 注册为 review-fix；沙箱对 runtime 临时文件再次 EPERM，仅注册命令通过升级执行成功，不启动 Office 或通知。

### FIX-001：缓存部件定位与真实持久化

IN_PROGRESS。首轮 cache-duties-net8-01 原始 TRX 为 12 项中 4 通过、8 失败，七个实例停止于 GetPivotCacheRecords().GetPackagePart()。该对象创建阶段未携带实际部件，不代表包中没有 records。

修复为读取缓存定义 GetPackagePart()、其 CT r:id、内部 pivotCacheRecords 关系及 GetRelatedPart 的实际地址；检查关系类型和 Internal 目标，不猜编号、不反射。cache-production-build-04 双 TFM Release 退出 0、0 警告/错误。随后 cache-duties-net8-02 为 13 项中 12 通过，完整 records/关系/源值职责通过，但混合长文本的换行在原生 cache definition 属性中被规范化为空格。

最终文本策略依微软 [ST_Xstring 契约](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d)：新增 NpoiPivotTextEncoding，在 CT_String.v 及 cacheField.name 原生赋值时一次编码属性 CR/LF/tab、XML 非法字符并保护字面 _xHHHH_；合法 Unicode 代理对保留。共享项去重和 longText 使用原始值/长度，不依据转义后的字符串膨胀长度。原始来源、records 索引和类型不改动。

此前尝试新增 DefinitionSnapshot/DefinitionWriter 的 XML 属性恢复路线已移除，仅移除本轮自身尚未验收的临时文件和追加重载；保持缓存定义的原生序列化，避免不必要的第二次定义转写。records 的实际 XML 输出仍复用既有一次 ZIP 暂存。文本协议直接测试及完整缓存测试继续保留原文语义，不将换行改为空格来通过。

### FIX-002：职责层级、轴项目与类型矩阵

IN_PROGRESS。无 records 的旧快速路径仍由原生 Write 控制流，其公开 Exporter 原本就使用 NonDisposingStream；测试改用同样包装并完整重开 XLS/XLSX、断言原流可用，不修改业务路径掩盖回归。RecordWriter 取消直接构造合法快照，避免被 Complete 前置失败遮挡；保留真实写入中取消、预取消和流所有权。增加 row 字段 data(x=0)、data(x=1)、default 的完整断言及混合类型/长文本 XML 矩阵。

cache-duties-net8-03 为 13 项中 12 通过：换行实体暂存路线已通过原文断言，剩余测试 helper 把所有缺失布尔属性当 false。改用公开 CT_SharedItems.Parse 读取 schema 默认值，而不强制生产写入默认属性迁就夹具。该结果早于最终 ST_Xstring 实现，后续复验另记，不冒充最终文本候选通过。

### NPIVOT-001：继续计划内来源与容量预检

新增 NpoiPivotSourceValidator、NpoiPivotSourcePlan、NpoiPivotLayoutBudget 内部单类型文件。读取业务 Sheet 的完整有界来源；拒绝任何字段的公式/Excel 错误/非有限数值，表头非空且大小写不敏感唯一；数值度量排除文本/布尔/日期且至少一项数值，Count 保留全部来源类型和空项。按实际工作簿 1900/1904 日期体系转换，异常保存原始一基坐标。

筛选匹配保留类型、文本忽略大小写且不 Trim，正负零归一；解析实际来源选项，但不减少来源记录。不可变来源记录/表头/列映射/筛选快照供后续 TableWriter 使用。完整轴层级及非叶自动小计、度量、总计、筛选/表头占用纳入 checked long 保守预算，既有单行 Sum 的两列边界保留。cache-production-build-05 双 TFM 编译退出 0、0 警告/错误；来源职责由 Luna 独立补充，尚未作为主流程能力启用。

### FIX-003 与 FIX-004

- FIX-003：DEFERRED 到前序 F-002；其文本 Count 原合同及真实客户端失败仍保留，不能因 F-003 缓存职责进展宣告解决。本轮不启动此前待确认的正常 Office 入口。
- FIX-004：PARTIAL。缓存候选先完成职责及主链，再执行当前源码/程序集/包身份、消费者、回归/常规门禁和最终独立复核。既有 API 形状差异为零仅是捕获时证据，未更新基线或审批当前新身份。

Round 1 尚未收口，execution 保持 IN_PROGRESS；不得 task-finish 或将完整 Goal 标记完成。

## 主链接入与索引零修复：当前实施

前序审查之后，NpoiReportWriter.ApplyPivotTables 已在全部业务 Sheet 写入后预检所有来源/目标，再调用 NpoiPivotTableWriter 创建原生透视。Exporter 声明既有 PivotTables/AdvancedPivotTables，仅 XLSX 新建；有透视时排除 SXSSF。来源预检、完整缓存、字段角色、类型化筛选、排序、小计及数字格式进入公开同步/异步文件/流入口。字典/级联辅助 Sheet 分配避让透视目标。

夹具先区分单选 pageField.item、多选隐藏项及 XML 布尔表示；随后新的冷进程实际文件复现 NPOI 2.7.4 省略 item=0。pivot-typed-cold-net8-02 保留真实失败；新增内部 NpoiPivotPageFieldWriter 捕获真实部件的不可变零索引选择，复用原有 ZIP 暂存补齐该属性。只处理明确单选零，不把全选/多选改成单选。三个局部 XML/快照/取消职责及完整类型矩阵通过，pivot-typed-cold-net8-03 新宿主冷启动并发 1/1。

默认文件提交服务的真实暂存写入失败、写入中取消、序列化后失败/取消、原目标保全、同目录清理、一次异常观察及调用方流所有权已覆盖同步/异步十二项。普通模板未覆盖样式、行列设置、名称、合并和公式保留通过；已有透视模板明确拒绝并保全输入/输出，仍不支持 F-004。

| 验证 | 真实终态 | 范围 |
| --- | --- | --- |
| pivot-page-fields-build-01 | 退出 0 | NPOI 双 TFM Release 生产构建 |
| pivot-public-net8-05 / pivot-public-net6-01 | 各 61/61，退出 0 | 来源、缓存、文本、原生字段、公开主链、资源和模板 |
| pivot-contract-net8-01 / pivot-contract-net6-01 | 各 158/158，退出 0 | 共享透视模型、Provider 及真实文件合同 |
| pivot-typed-cold-net8-03 | 1/1，退出 0 | 全新宿主冷并发、单选/多选/全选及完整类型化记录 |
| pivot-streaming-boundary-net8-01 | 1/1，退出 0 | 100K 阈值：100001 行、200002 单元格逐值；仅指定区域 3 条缓存 |
| pivot-composition-net8-02 | 8/8，退出 0 | 仅列轴、单/多度量、双筛选、同源/异源多透视、格式隔离、重复导出、后续预检失败重试 |

100K 探针实际输出 1,044,068 字节，夹具耗时 4,536ms，进程 PeakWorkingSet=295,219,200 字节；这些是本次环境事实，不是恒定内存或性能承诺。只运行计划指定的一项 Large 探针，没有重跑无关 500K/1M 矩阵。

组合首轮 pivot-composition-net8-01 为 2/8，失败是新夹具依赖原生重开对象的空 cache getter。改为按 pivot cacheId、workbook 关系和 cacheDefinition 关系解析实际 OPC 部件，完整记录和隔离断言保留，未修改生产逻辑掩盖问题。第二轮 8/8；原失败保留。

## 当前包消费者与 API 进展

pivot-feed-01 新生成 Abstractions、Core、Npoi、ClosedXml、MiniExcel、ExcelDataReader、SpreadCheetah、AsposeCells、EntityProviderContracts 九个 2.0.0 包，全部真实 pack 退出 0；未修改项目版本文件。三个独立包缓存通过 packageSourceMapping 限制 Bing.Offices.* 仅从新 feed 解析，其他依赖使用本地 NuGet 缓存。

新增 BasicPivotPackageContract/Row，普通 net6/net8 和第三方双 TFM 实际运行 NPOI 基础及高级同步/异步公开导出；原 ClosedXML 合同保留。完整来源、缓存每行每字段、OPC 关系/内容类型及公开重开成立。package-consumers-results-01.json 的十条 restore/build/run 均退出 0，其中四次实际消费者运行通过；这不是客户端聚合数值验收。

api-pivot-01 保留比较前 api-before-f003.json、原始候选和 api-zero-member-diff.json；双 TFM 各八个生产程序集公共成员差异为空。pivot-api-verify-01.command.json 实际退出 0；仅在源码/程序集/包身份比较成功后同步根基线。成员内容和候选身份与 api-approved-candidate.json 一致，未以覆盖基线掩盖成员差异。新增 api-approval.md 仅批准零成员变化与当前实施身份；生产友元仍仅向测试程序集开放。

当时剩余：新增组合/业务 net6.0、Release/Category!=Large/Docs、真实 NPOI 客户端完整数值矩阵及最终独立 Review。上述常规门禁、组合双 TFM、业务组合双 TFM 和基础 Excel 客户端现均已有后续通过证据；高级 Count 仍失败。F-004～007、E-001/E-002 与最终独立 Review 保留；未提交、推送、发布或改版本。

## 第十二次检查后的继续实施

上一 Goal turn 是进度：新增只读报告、核实原始 TRX/JSON 与当前调用链，明确下一步，不把审查当实现完成。本轮恢复 fix-review，沿用现有 active/review-fix 状态，不重复注册或修改历史 review。

- pivot-business-net8-01.trx 2/2：真实公开同步/异步动态列 Converter、Table、保护、完整业务来源/缓存和 OPC 关系通过。
- pivot-api-verify-01.command.json 退出 0：双 TFM 成员零差异、源码/程序集/包/审批身份比较通过；根基线成员与批准候选相同，身份同步成功。
- client-fixture-results-01.json 十条 restore/build/生成命令均退出 0；实际八次生成覆盖 NPOI 基础/高级、net6/net8、同步/异步。结构成功不代表客户端通过。
- 当时启动当前候选 solution Release 与 27 个 Category!=Large 项目/TFM 门禁，原始结果增量写入 regular-results-01.json；当时尚未终态的命令不计 PASS。该运行随后在“当前候选常规门禁终态”中记录为 28 条命令全部通过，不复用之前高级候选的 4975 结果。
- npoi-net8-advanced-sync-client-01.json/harness.json：真实 Excel 16.0 Build 20430 首开 Failed，V9 expected=3 actual=空；完整显示及实际筛选 Yes/Maybe 已记录，五聚合总计为 150/30/10/50/空。显式刷新及原生重建 Count 仍为空。65.24 秒、未超时、父进程在 COM 前取得专用句柄并发布身份；自身实例已退出，cleanupError=null。输入只读、未请求修复，诊断副本保留。该结果进一步证明问题不只出现在 ClosedXML 文件；未宣告根因或数值通过。

当时基础生命周期工具正在补齐同一父进程所有权入口。随后四个 net6/net8 同步/异步基础客户端运行均 Passed，basic-owned-runner 12/12；高级 Count 完整断言仍失败。继续处理 F-004～007 与 E-001/E-002；不因本地通过缩小 Goal。

### 当时候选常规门禁终态（历史时点）

当时 regular-results-01.json 共 28 条实际命令全部退出 0：1 次 Bing.Offices.sln Release 构建及 27 个项目/TFM Category!=Large 测试单元。矩阵 5113/5113，failed=0、notExecuted=0；NPOI net6.0/net8.0 各 1112/1112，包含当时的列轴/多透视/失败重试和业务组合。Docs 19/19，实际示例测试通过。Release 0 错误、4 项存量 Microsoft.Bcl.Memory net6 支持警告，未 suppress 或升级依赖。该历史结果不代表 2026-10-11 当前候选，当前结果见文末最新记录。

该门禁只证明当前本地候选，不关闭高级客户端 Count、已有透视模板、前向失败输出、有界 Entity、Entity 失败输出及环境验收。客户端工具变更另行验证实际进程所有权与真实工作簿。

## 2026-10-07 增量：Count 数据字段序列化约束

本轮继续处理高级透视表文本 `Count` 的结构化 TODO，未把客户端仍为空的结果伪装为通过。`NpoiPivotTableWriter.ApplyValues` 现在为每个数据字段明确写入 `showDataAs=normal`、`baseField=0` 和 `baseItem=0`，并关闭度量字段自身的默认小计；`NpoiPivotCacheWriter.CreateFieldItems` 仅为行、列和筛选轴建立共享项目，不再给数据字段附加轴小计 `countSubtotal` 或 `items`。数据区的 Count 语义只由 `<dataField subtotal="count">` 表达，同时保留完整缓存来源与记录索引。实现仅修改 NPOI XLSX 原生透视 XML，不改变公开 API、XLS/HSSF 边界、筛选裁剪语义或静态聚合策略。

`NpoiPivotTableWriterTest.Create_ShouldPersistAllRolesAndCompleteUnfilteredCache` 新增最终 ZIP 部件断言：Count 数据字段使用 `subtotal="count"` 且不携带轴字段 `countSubtotal`，全部数据字段包含 `showDataAs`/`baseField`/`baseItem`，缓存 records 的完整行列索引与内存 CT 模型一致。该直接测试在 NPOI `net8.0` 和 `net6.0` 均为 4/4；完整 `FullyQualifiedName~NpoiPivot` 职责集双 TFM 均为 72/72，全部退出 0；仅出现离线环境无法读取 NuGet 漏洞源的既有 `NU1900` 警告。

本轮生成的 `pivot-client-net8.0-advanced-counta-*` 探针副本及误建任务目录已清理。真实 Excel 客户端高级文本 Count 仍未验收：此前首开、刷新和原生重建均为空，当前 COM 环境也未形成新的可重复成功证据。因此 NPIVOT-002/004、F-002/F-003 继续保持 `PARTIAL`，下一步仍需稳定的真实客户端矩阵或明确兼容边界；不得用静态汇总、数值替代字段或仅结构化重开关闭该 TODO。

## 2026-10-07 增量：缓存定义文本类型标志

继续验证上一项 XML 收口时发现，NPOI 的 `CT_SharedItems` 对 `containsString`、`containsSemiMixedTypes` 和 `containsNonDate` 使用了默认值为 `true` 的 schema 属性，原生 XmlSerializer 因此会省略这些显式标志。缺少标志时，包内虽然已有完整的 `<s>` 共享文本和 `<dataField subtotal="count">`，真实 Excel 仍可能无法正确识别文本 Count 的缓存类型。

`NpoiSheetContentWriter.Write` 现在在同一次 ZIP 暂存中读取缓存定义关系部件，按 `pivotCacheDefinition -> pivotCacheRecords` 的 OPC 关系精确定位本次写入的缓存定义，仅对对应 `sharedItems` 显式写入实际需要的 `containsString`、`containsSemiMixedTypes` 和 `containsNonDate` 标志。未关联本次记录部件的模板缓存及其他包部件保持原字节复制；解析禁止 DTD，所有循环和 XML 写入均观察取消令牌，目标流仍由调用方持有。

新增 `NpoiPivotTableWriterTest` 的最终 ZIP 断言，核对文本缓存标志和 Count 数据字段；相关回归在 NPOI `net6.0`、`net8.0` 均通过 `28/28`，覆盖 `NpoiPivotTextEncodingTest`、`NpoiPivotTypedFilterWriterTest`、`NpoiPivotCacheWriterTest` 和 `NpoiPivotResourceTest`。此前本轮的 NPOI 公开导出与组合测试仍为双 TFM 各 `19/19`。测试仅出现离线环境无法读取 NuGet 漏洞源的既有 `NU1900` 警告。

这次修复补齐了缓存 XML 的结构契约，但没有新的真实 Excel 高级文本 Count 成功证据；现有首开、刷新和原生重建结果仍为空。因此 F-002/F-003、NPIVOT-002/004 继续保持 `PARTIAL`，不得以结构化包测试替代客户端聚合验收。

## 2026-10-07 增量：公开导出路径固化 Count 哨兵

为避免底层 writer 测试与公开出口之间出现覆盖断层，`NpoiPivotPublicExportTest.AssertWorkbook` 现在对同步、异步、基础和高级公开导出生成的每个 `dataField` 都断言 `baseField="-1"`、`baseItem="1048832"`，并继续核对 Count 的 `subtotal="count"`。该断言覆盖公开请求经过 `NpoiReportWriter`、`NpoiSheetContentWriter` 和最终 ZIP 转写后的结果。

验证结果：NPOI Pivot 过滤集 net8.0、net6.0 各 `73/73`；Provider Contract 的 `FullyQualifiedName~PivotTable` net8.0、net6.0 各 `113/113`。首次并行执行的 NPOI net6.0 命令曾因共享 `output\\debug` 依赖 DLL 锁竞争失败，随后串行重跑并通过；该环境失败不计入测试失败。真实 Excel 高级 Count 仍没有有效会话成功证据，F-002/F-003 继续保持 `PARTIAL`。

## 2026-10-07 增量：当前候选客户端启动边界

为避免复用修复前的包，本轮将当前 Abstractions、Core、NPOI 和 ClosedXML 重新打包到系统临时 feed，并用该候选构建 `PivotClientProbe`。NPOI `net8.0` 异步高级夹具生成及完整结构校验通过，输出包含 `Data`、`Summary`、完整缓存记录和本轮显式文本类型标志。

随后仅执行一次独立 Excel 生命周期入口，结果为 `status=Failed`、`startupError=[WinError 1312] 指定的登录会话不存在`、`childStarted=false`、`ownedExcelHandleObtained=false`。该失败发生在创建专用 Excel 实例之前，不能判断本轮缓存标志是否改变文本 Count；临时夹具和结果留在系统临时目录，未覆盖历史 artifacts。后续需要有效的交互登录会话再复验，仍不得用结构检查替代真实客户端验收。

## 2026-10-07 增量：对齐 NPOI 原生 Count 基准哨兵

离线提取 NPOI 2.7.4 原生 `AddColumnLabel(DataConsolidateFunction.COUNT, ...)` 的透视 XML 后，确认普通数据字段使用 `baseField="-1"`、`baseItem="1048832"` 表示没有基准字段和项目；当前自写路径此前强制写入 `0/0`。`NpoiPivotTableWriter.ApplyValues` 已改为使用同一原生哨兵值，保留 `showDataAs="normal"` 和 `subtotal="count"`，不改变公开 API、筛选语义或来源缓存。

新增直接对照测试 `NativeCountField_ShouldExposeExpectedPivotFieldShape`，并更新最终 ZIP 断言。NPOI Pivot 直接职责在 net8.0、net6.0 各 73/73 通过；Provider Contract 的 `FullyQualifiedName~PivotTable` 在 net8.0、net6.0 各 113/113 通过。该修复尚无真实 Excel 客户端成功证据；WinError 1312 的独立启动失败仍保留，F-002/F-003 继续为 `PARTIAL`，必须在有效客户端会话中复验首开、刷新和保存重开。

## 2026-10-07 增量：显式缓存类型标志完整化

继续核对最终透视缓存 XML 时发现，仅补写文本相关的三个标志仍会让 NPOI `CT_SharedItems` 的其他 schema 默认值泄漏到数值字段。`NpoiSheetContentWriter.RewritePivotCacheDefinitionXml` 现在按实际 `s`、`n`、`b`、`d`、`m`、`e` 子项计算并显式写出 `containsString`、`containsSemiMixedTypes`、`containsNonDate`、`containsDate`、`containsBlank`、`containsMixedTypes`、`containsNumber`、`containsInteger` 和 `longText`；文本长度在计算前解码 SpreadsheetML 转义。缓存记录、关系、数据字段和公开 API 均未改变，未关联本次记录部件的模板缓存仍按原字节复制。

新增 `NpoiPivotCacheWriterTest.ContentWriter_ShouldWriteExplicitNumericSharedItemFlags`，通过最终 ZIP 断言数值字段的 `containsNumber`、`containsInteger`、`containsNonDate` 为 `1`，其余不适用标志显式为 `0`。NPOI `NpoiPivotCacheWriterTest` 过滤集在 net6.0、net8.0 各 `14/14` 通过，只有离线环境无法读取 NuGet 漏洞源的既有 `NU1900` 警告。该修复改善缓存 XML 的结构确定性，但没有新的真实 Excel 高级文本 Count 成功证据；F-002/F-003、NPIVOT-002/004 继续保持 `PARTIAL`，仍需有效客户端会话完成首开、刷新和保存重开验收。

## 2026-10-09 后续：当前源码 NPOI net6.0 组合冒烟

- 使用一次性 MSBuild 参数 `DefaultItemExcludesInProjectFolder=output/**` 排除 Abstractions 项目中已有的 `output/summary-forward-net8-obj` 生成源码；未修改项目配置或清理现有输出目录。
- Release net6.0 当前源码编译通过 Abstractions、Core、NPOI Provider 和 NPOI 测试项目。运行 `NpoiPivotCompositionTest` 与 `NpoiPivotBusinessCompositionTest`，10/10 通过，覆盖列轴、多筛选、多透视缓存隔离、失败后重试，以及动态 Converter/保护/表格组合的同步与异步路径。
- 测试编译出现 `EntityLayoutBatchImportTest.cs` 中既有的 3 条 xUnit2013 集合断言警告；构建及测试命令退出码为 0。此次仅运行 NPOI net6.0 目标过滤集，不代表全解方案 Release、常规测试、NuGet 消费者或真实客户端门禁通过。
- 本次结果不能关闭 F-002：文本 Count 的真实 Excel 首开、显式刷新和保存重开仍无通过证据；F-003 仍待完整当前候选客户端验收及最终独立 Review。F-004～F-007 的客户端验证与未审批公共成员的正式双 TFM snapshot 继续保留为外部/审批阻塞。

本轮未修改生产源码、公共 API、版本号，未提交、推送或发布。

## 2026-10-11 增量：缓存构建与 ZIP 重写取消

- 新增 `Complete_CanceledDuringCacheBuild_ShouldLeaveCacheDefinitionUnchanged`，在输入行完成预检后、缓存字段仍由局部状态构建时触发取消，并断言原缓存字段引用、记录数、名称和共享项未改变。
- 新增 `ContentWriter_CancellationDuringZipRewrite_ShouldKeepCallerStreamOpenAndDeleteStagingFile`，在 ZIP 重写首个输出写入后取消，验证取消令牌传播、调用方流仍可写，并确认本次创建的 `bing-npoi-validations-*.tmp` 暂存文件已清理。
- NPOI `NpoiPivotCacheWriterTest` Release net8.0、net6.0 当前源码各 16/16 通过。验证使用单节点、按 TFM 隔离的临时输出，并通过一次性 MSBuild 属性排除既有 `output/**` 生成源码、禁用 XML 文档输出和共享编译器；三条 xUnit2013 集合断言警告来自 `EntityLayoutBatchImportTest.cs`。共享 TFM 输出目录曾导致运行配置混用，最终以独立目录结果为准。
- 本轮未改生产实现、公共 API、依赖、CI 或 API snapshot；双 TFM 定向职责测试通过。D-004 仍部分完成：缓存构建及 ZIP 取消的原子性/清理已有直接证据；其余阶段的取消响应与资源矩阵仍需逐项核对，故保持未完成。

## 2026-10-11 增量：页眉取消暂存清理证据

- `EntityRepeatedPageHeaderTest.PageHeader_MidWriteCancellation_ShouldPreserveOutput` 覆盖 XLS/HSSF 与 XLSX/XSSF 的第二页页眉值工厂触发取消；本轮补充断言，确认异步模板导出保留原输出字节、调用方输出流和模板流保持打开，且本次创建的 `bing-offices-entity-template-async-*.tmp` 暂存文件已清理。该用例 Release net8.0、net6.0 各 2/2 通过。
- 更新 `production-test-map.md`，将重复页眉取消资源行为映射到 `NpoiEntityExportExecutor.WritePageHeader` 和直接测试方法。
- 构建仅出现 `EntityLayoutBatchImportTest.cs` 中既有的 3 条 xUnit2013 集合断言警告。D-004 继续部分完成：来源预检已有预取消/失败重试证据，缓存构建与 ZIP 重写有双 TFM 直接测试，页眉源范围复制完成后的值工厂取消有 XLS/XLSX 资源断言，文件提交已有同步/异步写入中和序列化后取消矩阵。来源扫描和页眉复制循环中途取消尚无确定性运行时注入；`NpoiNonDisposingStream` 会在同步序列化写入/刷新边界检查令牌，但该阶段的检查间隔和取消响应耗时尚未独立测量，待后续补证。

## 2026-10-11 增量：来源扫描中途取消

- `NpoiPivotSourceValidator.Validate` 保留原有生产入口，并增加内部行读取观察重载。观察器在数据行读取完成后触发；生产调用传入空观察器，并在行边界检查取消令牌，未改变公开 API。
- 新增 `NpoiPivotSourceValidatorTest.Validation_ShouldHonorCancellationDuringSourceScan`：观察器在完整读取前两条数据行后同步取消；断言随后抛出携带原令牌的 `OperationCanceledException`、未继续读取第三行，且工作簿来源未被修改。
- 更新 `production-test-map.md`，将扫描中途取消行为映射到验证器与直接测试。
- 双 TFM Release 定向测试及冒烟结果记录在下方。该证据只关闭来源扫描中途取消子项；D-004 仍部分完成，因为页眉复制循环的中途取消和同步序列化取消响应间隔仍待验证。
- 首次常规构建命令因 `src/Bing.Offices.Abstractions/output/**` 中的预存生成源码触发重复程序集特性；随后仅在命令行增加 `DefaultItemExcludesInProjectFolder=output/**`，未修改项目配置。新测试 Release net6.0、net8.0 各 1/1 通过；完整 `NpoiPivotSourceValidatorTest` 类双 TFM 各 19/19 通过。首次构建还报告既有 `ExcelPivotTemplateManifest.cs` XML 参数注释警告及 `EntityLayoutBatchImportTest.cs` 的 3 条 xUnit2013 警告。

## 2026-10-11 增量：页眉复制循环中途取消

- `NpoiEntityExportExecutor.WritePageHeader` 在复制页眉源区域时对每个单元格、区域复制结束、合并及 marker 收尾前检查取消令牌；新增内部复制进度观察参数仅用于确定性职责测试，不改变公开 API。生产路径仍通过原有调用点，且不提供观察器。
- 新增 `EntityRepeatedPageHeaderTest.PageHeader_CancellationDuringSourceCopy_ShouldStopBeforeFinalization`，分别在页眉首格和末格复制后同步取消；XLS/HSSF 与 XLSX/XSSF 均断言异常携带原令牌，复制按预期停止，合并区域和 marker 收尾未执行，来源单元格未被修改。与既有 `PageHeader_MidWriteCancellation_ShouldPreserveOutput` 组合，覆盖复制循环取消与实际异步模板输出/模板流/暂存文件资源行为。
- 新用例首格/末格取消矩阵 Release net6.0、net8.0 各 4/4 通过；完整 `EntityRepeatedPageHeaderTest` 类双 TFM 各 92/92 通过。构建出现的 `ExcelPivotTemplateManifest.cs` XML 参数警告和 `EntityLayoutBatchImportTest.cs` 三条 xUnit2013 警告均为既有警告。
- D-004 仍部分完成：来源预检、缓存构建、页眉复制、ZIP 重写及提交阶段均已有取消证据；`NpoiNonDisposingStream` 同步序列化阶段的检查间隔与取消响应耗时仍未独立测量。

## 2026-10-11 增量：同步序列化写入边界取消

- `NpoiNonDisposingStream.Write` 已在每次底层同步写入前后检查令牌。本轮未修改生产实现；新增 `NpoiPivotResourceTest.StreamExport_CanceledDuringSynchronousSerialization_ShouldObserveAtWriteBoundary`，通过真实 `NpoiExcelExporter.Export` 路径在首个非空底层写入完成后取消，断言同一写入边界抛出携带原令牌的取消异常、没有第二次非空写入、调用方流保持可写且异常观察器未收到取消异常。
- XLS/HSSF 与 XLSX/XSSF 新用例在 Release net6.0、net8.0 各 2/2 通过；完整 `NpoiPivotResourceTest` 类双 TFM 各 14/14 通过。构建中仅出现 `EntityLayoutBatchImportTest.cs` 的 3 条既有 xUnit2013 警告。
- 同步引擎序列化的取消粒度以目标流写入边界为准；NPOI 内部两次写入之间的工作不可由此包装器抢占，不承诺墙钟响应上限。来源扫描、缓存构建、页眉复制、序列化写入边界、ZIP 重写和文件提交均已有取消/资源证据，D-004 完成。

## 2026-10-11 后续：当前候选本地编译与 D-004 冒烟复核

- 当前 `Bing.Offices.sln` Release 构建因仓库内多个既有 `obj`、XML 文档及 `output/release` 路径返回 `Access Denied` 失败；19 个错误均为文件写入失败，不能据此判为源码编译错误。没有修改 ACL 或清理现有输出。
- NPOI Provider 当前源码通过系统临时目录隔离输出，分别以 Release/net6.0、Release/net8.0 构建，均为 0 警告、0 错误。
- 使用现有 Release 测试程序集执行 `NpoiPivotResourceTest`、`NpoiPivotSourceValidatorTest`、`NpoiPivotCacheWriterTest` 与 `EntityRepeatedPageHeaderTest`；net6.0、net8.0 各 141/141 通过。该命令为 `--no-build`，不作为本轮测试程序集重新编译证据。
- 尝试将 NPOI 测试项目还原到隔离 artifacts 目录，命令以退出码 1 结束且未提供诊断；因此本轮没有声称隔离测试项目编译完成。全解方案本地门禁仍未完成。
- 未运行 Excel/WPS/LibreOffice；未改变源码、公共 API、依赖、CI、API snapshot 或版本，未提交、推送或发布。临时 MSBuild 输出映射文件已移除。

## 2026-10-11 后续：隔离构建环境诊断

- `dotnet --version` 正常返回 `10.0.401`。尝试对 Abstractions `netstandard2.0` 执行隔离 Release 构建时，命令超过 20 秒无输出、无 CPU 增长且未生成目标程序集；仅结束本轮启动的构建 PID `12624`，未结束其他进程。
- 同一时点宿主可见 839 个 `dotnet` 进程，读取进程命令行被系统拒绝；无法证明这些进程均属于本任务，故未批量终止。全解当前候选构建仍未通过，本轮未获得新增编译或测试通过证据。
- 清理本轮创建的两个隔离 MSBuild 临时文件；未更改 ACL、仓库输出、项目配置或代码。NPOI 功能 TODO 状态不变。

## 2026-10-11 后续：T-004 文档问题修复与复核

- 根据 T-004 Review，修正 `docs/excel/nuget-migration.md` 的 Provider 格式能力说明：NPOI XLSX 与 ClosedXML XLSX 支持基础及高级透视新建；NPOI XLS/HSSF 不支持；高级 Count 的客户端数值验收仍未完成。
- 按 follow-up review 分别修正文档范围：迁移说明中的 Excel 历史验收仅限定为基础透视；`docs/excel/pivot-tables.md` 明确较早候选的门禁结果不能代表 2026-10-11 当前候选已通过 Release、API 或包身份检查。
- 对产品文档执行针对性全文检索，确认能力和客户端验收表述没有相反声明；本轮改动的 Markdown 文件均通过严格 UTF-8、无 BOM、LF 和末尾换行检查。
- 最后一次 Docs.Tests 冒烟运行使用已有 Release 输出及 `--no-build --no-restore --filter Category!=Large`，net8.0 为 19/19 通过。该结果只作为文档测试冒烟证据，不代表本轮重新编译，也不代表当前候选完整 Release 门禁通过。
- 当前候选完整 Release/API/包身份门禁仍为 `NOT_VERIFIABLE`：此前全解 Release 构建在既有输出路径遇到 `MSB3491 Access Denied`；隔离测试项目还原未提供诊断。本轮未通过旧候选结果替代当前候选证据，也未重跑已知受阻的全解门禁。
- 高级文本 Count 客户端行为仍为 `BLOCKED_EXTERNAL`。按用户明确的非 Excel 客户端范围，本轮未启动 Excel、WPS 或 LibreOffice，未将此项标记完成。
- 任务状态登记脚本在写入 `execution.md` 时返回 `EPERM`；核对 `.agents/runtime/current-task.json` 后确认未发生状态变更，当前任务仍为 `plan-execution / implementing`。本记录通过工作区文件编辑追加。
- 审查产物中提到的 `review-20261005-historical.md` 当前不存在；本轮不声称旧审查报告已归档。现有 `review.md` 与 `review-followup-t004-fix001-20261011.md` 保持审查记录，最终复核结论以本轮 follow-up review 为准。
- 本轮未修改生产代码、测试代码、CI、公共 API、API snapshot 或版本；未提交、推送或发布。

## 2026-10-11 后续：T-004 最终独立复核

- 最终独立复核记录：`review-followup-t004-final-20261011.md`，结论 `PASS`。T-004 FIX-001、FOLLOWUP-001、FOLLOWUP-002 均为 `CLOSED`，全文检索未发现未限定的高级客户端通过声明或当前候选门禁已通过的旧说法；本轮文档范围 `OPEN_ACTIONABLE: none`。
- 复核明确限定 PASS 只覆盖文档修复。Docs.Tests 19/19 是 `--no-build --no-restore` 冒烟，不代表当前候选重新编译；当前 Release/API/包身份仍为 `NOT_VERIFIABLE`，Release 门禁为 `BLOCKED`。
- 高级文本 Count 客户端验收仍为 `BLOCKED_EXTERNAL`。按用户范围未启动 Excel、WPS 或 LibreOffice；该外部门禁不转换为代码缺陷，也不标记完成。
- 复核确认 `review.md` 所述 `review-20261005-historical.md` 实际不存在；执行记录已如实说明，未声称已归档。
- 最终复核未修改业务代码、测试、API、执行记录或运行状态。任务后续仍取决于独立的当前候选门禁和外部客户端条件；未提交、推送或发布。

## 2026-10-11 后续：隔离 artifacts-path Release 重试

- 为重新验证当前候选 Release，使用 .NET SDK 10.0.401 的 `--artifacts-path` 将工件及还原输出定向到系统临时目录；31 个解决方案项目均成功还原。
- 构建仍以退出码 1 失败：`GenerateDepsFile` 尝试写入 `output/release/Bing.Offices.Abstractions.deps.json` 时返回 `UnauthorizedAccessException`。`common.props` 的 Release `OutputPath` 覆盖了隔离 artifacts 输出，因此没有形成完整编译证据；该结果是现有输出路径访问失败，不作为源码编译错误或通过。
- 仅删除本次创建的唯一临时 artifacts 目录，删除前确认其规范化路径位于系统临时目录且名称带本次生成前缀。未修改 ACL、项目配置、仓库输出、代码或 API；Release/API/包身份仍为 `NOT_VERIFIABLE`。

## 2026-10-11 后续：隔离输出 Release 构建与常规测试

- 使用临时 MSBuild target 按项目和 TFM 重定向 `OutputPath`、`OutDir`、`TargetDir`、程序集/依赖清单路径及 `DocumentationFile`，并将 restore artifacts 放入系统临时目录；没有修改仓库构建配置。
- 当前工作区 `Bing.Offices.sln` Release 全解方案构建成功，退出码 0、0 错误、145 条警告。该命令编译了当前工作区项目及测试项目；警告包括既有 XML 注释、nullable、xUnit 及 net6.0 依赖支持提示。
- 紧接着运行 `dotnet test Bing.Offices.sln -c Release --no-build --no-restore --filter "Category!=Large"`，退出码 1，常规测试门禁未通过。多组文件提交测试在 `File.Replace` 处遇 `UnauthorizedAccessException`；Docs `SourceLayoutTests` 从临时程序集目录无法定位仓库根；ClosedXML `EntityLayoutBatchImportTest.ImportEntityBatches_Forward_SummaryOnly_ShouldWriteOnlyErrorRowsWorkbook` 的同步与异步场景因共享失败工作簿写入器拒绝未支持的列元数据而失败。
- `File.Replace` 及 Docs 当前工作目录失败按环境/隔离运行路径归类，不转换为生产代码缺陷。ClosedXML Forward SummaryOnly 失败工作簿列元数据行为保留为 F-006/F-007 的后续可执行检查；当前活动优先级仍在 F-003，不在本轮跳转修复。
- 当前候选 Release 编译门禁现有通过证据；常规测试门禁未通过，API/包身份比较仍未绑定本次编译候选，故 NPIVOT-005 仍未完成。没有运行 Office 客户端或 Large 性能矩阵。
- 清理本轮创建的临时 MSBuild target 与唯一临时输出目录；删除前确认输出目录位于系统临时目录且名称为本轮专用值。未修改生产代码、测试、项目配置、CI、API snapshot、版本或访问控制。

## 2026-10-11 后续：当前候选全解与 API 快照复核

- 为使 Docs 源布局测试能从程序集目录定位仓库根，并避开 `common.props` 固定的 `output/release` 写入限制，使用仓库内一次性 MSBuild target 和 `.codex-p5-validation-20261011` 隔离输出；全解 Release 构建退出 0、0 错误、145 条警告。
- 以 `Category!=Large` 运行当前 Release 测试，退出 1。NPOI.Tests net6.0/net8.0 各 1745/1745，NPOI.Tests.Integration 各 28/28，ProviderContract.Tests 各 679/679，Docs.Tests 19/19 通过。ClosedXML.Tests 两个 TFM 均为 792/795：`EntityForwardResourceTest.MaxInputBytes_ShouldRejectBufferedNonSeekableInput(asynchronous: true)` 的暂存文件集合断言失败；`EntityLayoutBatchImportTest.ImportEntityBatches_Forward_SummaryOnly_ShouldWriteOnlyErrorRowsWorkbook` 同步、异步均因 `ExcelDataReaderFailureWorkbookWriter.ValidateColumnMetadata` 拒绝额外列元数据。后两项属于 F-006/F-007 后续范围，本轮不越序修改。
- 提权运行后未复现先前的 `File.Replace` `UnauthorizedAccessException`；Docs `SourceLayoutTests` 也能从隔离输出找到仓库根。Integration API 合同测试读取固定的 `output/release`，因此另用当前隔离构建程序集运行 API snapshot capture 与成员比较，避免覆盖默认输出或快照。
- 当前候选与 F-003 批准基线比较：Abstractions net6.0/net8.0 各新增 27 条 API 记录、删除 0；Core、NPOI、MiniExcel、ClosedXML、ExcelDataReader、SpreadCheetah、AsposeCells 均为 0 新增/0 删除。Abstractions 新增项涉及 Entity Batch 与透视模板契约，超出 F-003“零公共成员变化”批准；保持基线不动，需先依据所属任务的成员审批状态处理，再重新执行当前候选比较。
- 本轮未重跑 NuGet pack/consumer；既有 consumer 结果保留为此前候选证据，不绑定本次完整编译候选。未运行 Excel/WPS/LibreOffice 或 Large 性能矩阵；未修改生产源码、测试、CI、API snapshot 或版本，也未提交、推送或发布。

## 2026-10-11 后续：当前候选包门禁重试

- 复核 `build/PivotClientProbe/README.md` 已提供 `--provider NPOI` 的双 TFM、同步/异步基础及高级入口，并明确 `--number-format-probe` 仅支持 ClosedXML；`execution.md` 与 `production-test-map.md` 已记录当前候选组合测试及验证边界。T-003 的文档同步项完成。
- 为避免覆盖任务目录中旧 feed 和旧结果，在 `.codex-temp/package-run-20261011` 使用既有 `run-package-gates.py` 的隔离副本及临时 MSBuild 路径。解决方案 restore 到隔离 assets 成功；当前候选 pack 在 Abstractions 首个包阶段未完成。
- 当前候选编译 DLL 写入隔离目录，但 `GenerateDepsFile`/输出复制仍尝试写入仓库固定 `output/release`，随后因 `UnauthorizedAccessException` 失败。失败发生在包创建前，没有运行任何当前候选 NuGet consumer；不以早前九包及四次 consumer 成功替代。
- 临时尝试未修改访问控制、项目配置或仓库输出。常规 Release 与 `Category!=Large` 结果仍以本节前述当前候选记录为准；API Abstractions 27 条新增记录及双 TFM零删除差异仍未更新正式基线，需按所属任务逐成员审批后重新比较。
- 未运行 Office 客户端或 Large 性能矩阵；未修改生产源码、测试、CI、正式 API snapshot 或版本，也未提交、推送或发布。

## 2026-10-11 后续：隔离候选包、消费者与常规测试复验

- 使用本轮专用 `.codex-temp/package-run-f3-isolation-20261011` 及临时 `Directory.Build.props/targets` 将 restore assets、编译输出和包 feed 隔离于仓库固定 `output/release`。31 个解决方案项目还原成功；Release 全解方案构建退出 0、0 错误、145 条警告。
- 当前候选九个包均已生成：八个生产包及 EntityProviderContracts 测试支持包。普通 Consumer net6.0/net8.0 的 restore/build/run 均通过；第三方 Consumer 在沙箱内因 `File.Replace` 权限失败，提升文件访问权限后 net6.0/net8.0 均通过。包和 API capture 已保存至 `artifacts/current-candidate-20261011/`。
- API capture 相对 F-003 批准候选在 Abstractions 双 TFM 各发现 27 条新增记录、0 条删除；其他生产程序集为 0 差异。新增项为 F-004 的 7 项、F-005 的 1 项、F-006 的 17 项及 F-007 的 2 项，其中 19 项仍待成员级审批。正式 API snapshot 未更新；常规集成 API 合同因此双 TFM 各有 1 项失败。
- 提升文件访问权限后运行完整 `Category!=Large`：NPOI.Tests 双 TFM 各 1745/1745、NPOI 集成各 28/28、Provider Contract 各 679/679、Core `Bing.Offices.Tests` 各 204/204、MiniExcel.Tests 各 57/57、MiniExcel 集成各 7/7、ExcelDataReader 各 103/103、SpreadCheetah 各 69/69、AsposeCells 各 2/2、Analyzers 各 20/20 通过；Docs.Tests net8.0 为 19/19。`Bing.Offices.Tests.Integration` 双 TFM 各 49/50，唯一失败为正式 API snapshot 合同；ClosedXML.Tests 双 TFM 各 792/795，分别有 1 个暂存文件集合断言和 2 个 SummaryOnly ErrorRowsOnly 列元数据失败。常规门禁退出 1。
- 未启动 Excel、WPS 或 LibreOffice，也未运行 Large 性能矩阵。T-001 已有当前候选构建、包和消费者证据，但常规门禁及 API 差异未闭合，仍为部分完成；T-002 客户端生命周期与 T-004 最终独立 Review 未完成。没有修改生产源码、测试、CI、正式 API snapshot 或版本，也未提交、推送或发布。

## 2026-10-11 后续：Forward 暂存文件测试隔离复验

- `EntityForwardResourceTest` 会对进程级临时目录中的 `bing-offices-forward-*.xlsx` 做前后快照；ClosedXML 的其他 Forward 测试也使用相同文件名前缀，测试并行时会造成快照误判。本轮将资源测试放入禁并行集合，隔离该全局文件快照。
- 重新编译受影响项目后，定向 Forward 测试在 net6.0、net8.0 各通过 33/33，覆盖资源快照、页脚写入和非定位输入批量导入用例。
- 本轮只修改测试隔离与本执行记录；未修改生产代码、公共 API、CI、API snapshot 或版本。完整 `Category!=Large` 门禁未重跑，之前记录的 API snapshot 差异及 F-007 `style="0"` 列元数据失败仍未闭合。未运行 Excel、WPS、LibreOffice 或 Large 性能矩阵；未提交、推送或发布。

## 2026-10-11 后续：SummaryOnly 列元数据夹具边界

- ClosedXML 的 SummaryOnly 错误行测试使用 ClosedXML 生成带列宽和隐藏状态的来源；生成的 `<col>` 同时包含 `style="0"`，而 F-005 已批准契约明确拒绝所有列样式属性。本轮只在该测试用例夹具中移除默认样式索引，仍保留列宽、隐藏状态、失败行和错误摘要断言；生产校验及对真实样式的拒绝行为未放宽。
- `ImportEntityBatches_Forward_SummaryOnly_ShouldWriteOnlyErrorRowsWorkbook` 从本轮隔离构建程序集执行，net6.0/net8.0 各 2/2 通过。ExcelDataReader `Import_ErrorRowsOnly_XlsxColumnMetadata_UnsupportedAttribute_ShouldRejectBeforeWriting` 双 TFM 各 1/1 通过，确认不支持的非默认样式仍在写出前被拒绝。
- 首次标准构建因仓库 `obj` 写入拒绝失败；后续使用项目/TFM 隔离的临时 MSBuild 输出重新编译并运行，成功结果来自隔离目录中的测试程序集。离线 NuGet 漏洞源及既有 XML/xUnit 分析器警告保留。
- 本轮未修改生产代码、公共 API、CI 或 API snapshot；未重跑完整 `Category!=Large` 门禁，因此不据此宣告整体验收通过。未运行 Office 客户端或 Large 性能矩阵，未提交、推送或发布。

## 2026-10-11 后续：ClosedXML 常规项目门禁复验

- 使用隔离的项目/TFM MSBuild 输出编译当前 ClosedXML Provider 和测试程序集，运行 `Category!=Large` 完整测试集。net8.0 在默认沙箱权限下为 793/795；唯一两例是原子文件提交测试的 `File.Replace` 返回 `UnauthorizedAccessException`。提升文件访问权限后，两个原失败用例 2/2 通过，完整 net8.0 集合 795/795 通过。
- 以同一隔离配置编译并运行 net6.0 完整集合，`Category!=Large` 为 795/795 通过。Forward 临时文件竞态与 SummaryOnly 列元数据夹具问题均未在完整项目集合中复现。
- 本轮只更新本执行记录；未修改生产代码、公共 API、CI、正式 API snapshot 或版本。该结果只证明 ClosedXML 项目双 TFM 常规测试通过；整个解决方案门禁及当前 API 候选仍需按成员审批流程收口。未运行 Office 客户端或 Large 性能矩阵，未提交、推送或发布。
