<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001
AI_EXECUTION_FINISHED_AT: 2026-10-15T14:49:00Z

# 实施执行报告

## 结果与范围

本轮优先实现本地 F-006 资源 TODO 的临时磁盘预算。代码及直接测试完成；正式发布验收仍受成员审批和仓库定位测试环境阻塞，任务保持 PARTIAL，不把 F-006 或历史 TODO 全部标为完成。

新增 ExcelEntityBatchImportOptions 的五参数构造与 MaxTemporaryStorageBytes。NPOI、ClosedXML Forward/SummaryOnly 共享单次调用预算，统计输入暂存、共享字符串正文/索引；NPOI XLS 还计入转换 XLSX 和逐 Sheet XML。null 保留原行为。Buffered+Complete、模板导入和失败工作簿组合在指定预算时读取前拒绝。

## 计划与完成情况

| 项目 | 结果与直接证据 |
| --- | --- |
| 增量选项和默认兼容 | 完成；保留旧构造；默认、非正数与 consumer 测试 |
| 聚合计费和释放 | 完成；增长前预留、覆盖写、Seek、SetLength、空写、同步/异步及并发隔离 |
| NPOI/ClosedXML 主链路 | 完成；真实 DI 入口精确预算及少一字节拒绝、三种支持模式 |
| NPOI XLS 中间文件 | 完成；HSSF 成功/转换超限、同步/异步 |
| 取消和异常清理 | 完成；输入 IO 中途取消/失败、回调取消/失败、调用方流所有权 |
| 共享字符串清理异常 | 完成；一个 Dispose 抛错仍释放另一文件，保留原始 XmlException |
| 使用说明及追溯 | 完成；batch-import.md、symbol-test-map.md、api-candidate.json |
| 正式 API/全解决方案门禁 | 未关闭，见验证与风险 |

## 实现与变更影响

新增共享 ExcelTemporaryStorageBudget，使用 Stream 包装普通 FileStream；源文件链接到两个 Provider。修改两个 EntityBatchImporter、Forward 输入与共享字符串、NPOI HSSF 转换链路。公共 API 增量 2 成员；旧签名、枚举和默认预算保持兼容。无新增依赖、产品版本或打包拓扑变化。测试文件和职责对应见 symbol-test-map.md，两个 Provider 各有独立公共入口测试。

预算为同时存活受管临时文件的逻辑长度，不含调用方流、文件系统块开销、引擎缓存或实体堆。Complete 仍保留完整结果。每次调用独立，不使用静态环境开关。

初版继承 FileStream 导致 Span 写虚调用锁重入，最终改为 Stream 包装并通过测试。独立审查发现异常清理可能覆盖解析异常，已修复并补故障注入。未扩大为无关重构。

## fresh 验证

证据目录：C:/Users/jianx/AppData/Local/Temp/bing-offices-entity-temp-budget-20261015。

| 验证 | 实际结果 |
| --- | --- |
| 全解决方案 Release 构建、no-restore、m:1 | PASS，145 warnings、0 errors；build.log |
| NPOI net6.0/net8.0 普通测试 | PASS，各 1,778 |
| ClosedXML net6.0/net8.0 普通测试 | PASS，各 826 |
| NPOI Integration 双 TFM | PASS，各 28 |
| ClosedXML Integration 双 TFM | PASS，各 4 |
| ProviderContract 双 TFM | PASS，各 679 |
| 本轮预算及新增清理案例 | PASS，NPOI 各 33、ClosedXML 各 31，共 128 次执行 |
| 四包本地 pack、真实 NuGet consumer 双 TFM | PASS，独立缓存，调用两个 Provider 的真实预算拒绝链路 |
| 双 TFM API 候选捕获 | 各 1,714 条；相对工作区 baseline 新增 29、删除 0；本轮 2、历史候选 27 |
| 全解决方案普通回归 | 总计 7,673，通过 7,660、失败 13；退出 1，不宣称全绿 |
| git diff --check | PASS；既有行尾提示保留 |
| 本轮文件严格 UTF-8 | 27 文件解码通过，CS BOM 和末尾换行通过；存量 csproj 例外见后文 |

结果 TRX 位于 solution-results。普通回归筛选 Category!=Large；本轮没有重跑与临时磁盘额度无关的百万行性能矩阵。正确性采用精确字节边界、跨文件聚合、释放、取消直接证据，不用吞吐代替。

13 个失败：PublicApiContractTest 两方法各双 TFM 共 4 个，Docs SourceLayoutTests 共 9 个。隔离输出位于系统临时目录，测试按 AppContext.BaseDirectory/当前目录寻找仓库失败。没有修改测试或伪造 baseline。正式成员审批也仍未关闭。

早期沙箱 File.Replace 权限拒绝经允许本地原子文件测试的运行环境验证。旧 PageHeader 测试曾因全局临时目录快照受并行执行干扰而失败，定向两案例通过，最终 NPOI 双 TFM 普通套件均全绿。预算定向双 TFM 曾并行执行造成同类快照干扰，最终按 TFM 串行验证。

## 命令、环境与证据复现

SDK 10.0.401，运行时 net6.0/net8.0。仓库 obj 写入受限，采用临时 isolated.props/isolated.targets，显式设置 BaseIntermediateOutputPath、OutDir、TargetDir、TargetPath，并启用默认 Compile/资源项，避免环境响应文件关闭 Compile 项。未改仓库构建配置。

```text
dotnet build Bing.Offices.sln --no-restore -c Release -v:minimal -m:1 [隔离属性]
dotnet test Bing.Offices.sln --no-build --no-restore -c Release -v:minimal -m:1 [隔离属性] --filter Category!=Large --logger trx
dotnet pack [Abstractions/Core/Npoi/ClosedXml] --no-build --no-restore -c Release -o [临时 packages]
dotnet run --project [临时 Consumer] --no-restore -c Release -f net6.0/net8.0
git diff --check
```

隔离属性、build.log、四包、consumer 源码及 TRX 留在证据目录。完整快照为 api-net6.txt、api-net8.txt；仓库 api-candidate.json 保留相关类型候选。正式 baseline 本轮未更新。

## 风险、剩余与编码

新 2 成员与历史 27 条候选仍待审批，详见 api-approval.md；无法据此宣称正式发布门禁通过。剩余代码方向是堆内存/DOM/完整实体总预算，以及模板与 Failure Workbook 的预算组合。临时文件额度不等于进程恒定内存。历史客户端和商业 Provider 外部验收保持原状态。

严格 UTF-8 检查通过；本轮新增 CS 使用 BOM+LF。ClosedXML.csproj 有两条既有 LF 链接行，其余 63 行 CRLF，本轮保留，没有整文件重写；全工作区既有文档行尾偏差未混入本轮修复。

## 独立审查与 Git

独立 Reviewer 已复核清理修复，OPEN_ACTIONABLE=0；最终依据 review.md。存量无关工作区改动全部保留。未 commit、push、创建分支/PR、发布、升级依赖或版本，也未发送外部通知。

## 2026-10-15 续轮：隔离输出下的仓库定位

用户要求先读 TODO、按既有顺序推进且只做必要编译/冒烟。本轮读取第七次高级透视 TODO、原独立页锚点 TODO、对应后续 execution、F-005/F-006/F-007 和本任务 plan/TODO。不同任务复用 T/D 编号，汇总保留来源，未合并或重排。F-002 客户端数值合同仍阻塞，F-003/F-005 API 收口未完成；F-006 整体堆预算没有获批的总额/失败终态续阶段契约，不能自行扩范围。本轮处理既有验收待办中的可独立定位问题，没有新增业务能力。

### 改动与影响分析

ChangedFiles 为 SourceLayoutTests.cs、PublicApiContractTest.cs 及本任务记录；ChangedProjects 为 Docs.Tests、Tests.Integration；ChangedPublicContracts、ChangedRuntimePaths、ChangedProviders、ChangedTFMs、ChangedBuildPackaging、ChangedBenchmarkHarness 均无；ChangedDocsOnly 为否；AffectedDependents 无生产依赖；RiskLevel LOW。

两个测试私有定位方法保留程序集祖先搜索，并补编译时 CallerFilePath 和当前工作目录祖先搜索；未找到时抛出明确 DirectoryNotFoundException。没有静默跳过断言、复制基线、改程序集集合或变更固定 Release 输入目录。项目配置、CI、job、matrix、快照和依赖均未修改，也没有清理无关配置。

### 必要验证与实际结果

- Docs.Tests Release 编译通过，0 warnings/0 errors；SourceLayoutTests 在系统临时输出执行 9/9 通过。
- Tests.Integration Release 双 TFM 编译通过，0 errors、2 条既有 net6 支持提示。
- 两 TFM 的 PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies 各 1/1 通过。
- 两 TFM 的 PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot 仍失败，现已找到真实仓库及 baseline；失败明确是统一 output/release 输入缺少程序集：net6 缺 NPOI/MiniExcel/ClosedXML/ExcelDataReader，net8 缺 MiniExcel/ExcelDataReader。正式成员审批也保持原状态，没有把定位修复等同于 API 门禁通过。
- 改动文件 git diff --check 通过；UTF-8/BOM/LF/末尾换行检查通过。复用此前未变更生产源码的预算测试证据；不重跑全解、包、容量或客户端校验。

沿用前轮 isolated.props/targets 和 restore 缓存，增加 BuildProjectReferences=false、UseSharedCompilation=false，仅重新编译两个实际改动的测试项目；不引入工具链。实际测试用 dotnet vstest 从系统临时路径启动，覆盖本轮回归原因。TRX 位于证据目录 root-resolution-results/{docs-root-net8,api-root-net6.0,api-root-net8.0}.trx。

### 本轮终态与下一项

Completed：测试仓库定位缺陷修复（两方法）。Open Actionable：本轮选定缺陷无剩余。Blocked Approval：历史公共成员及新增预算两成员。Blocked External：固定统一 Release 输入尚不齐全、真实 Office/商业环境。Accepted Limitation：预算只统计临时磁盘，不等于总内存。Deferred：堆预算续阶段契约、可选扩展；Not Applicable：本轮没有生产性能变化；Verified Boundary：定位成功后仍正确拒绝缺失 Release 输入。No-Progress Check CHANGED；Goal Status STOPPED_BLOCKED。原完整 TODO 均按实际状态保留，详见 todo-status-20261015.md。未 commit/push/PR/发布/升级版本或依赖；Finalizer 禁用外部通知。
