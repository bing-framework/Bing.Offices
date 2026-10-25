<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BO-API-USABILITY-20261025
AI_REVIEWED_AT: 2026-10-25T10:25:04Z

# 易用性改进复查

## 结论

主体功能已真实实现，接入、单表导入和状态摘要均进入主调用链；未发现空实现或旧 API 删除。但本轮确认两项 P2 问题需要修正：新增跨 Sheet 诊断的表名不正确，以及已有默认输入字节上限未在默认调用中生效。后一项由既有 Workbook 管线遗留，新单表入口继承了该行为，不能将其误记为本次新引入。

IMPLEMENTATION_STATUS: NEEDS_FIX
TEST_STATUS: PASS（既有套件；新增缺陷不在当前断言覆盖内）
OPEN_ACTIONABLE: 2
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL
RELEASE_STATUS: NOT_VERIFIED
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE
RESOURCE_EVIDENCE_STATUS: PARTIAL（发现默认输入限制接线缺口；未做新容量测量）

本轮只审查，未修改业务代码或正式测试，也未修改 plan/execution。新增本 review 和忽略目录中的独立复现程序。未发送外部通知、未提交或推送。

## 验收矩阵

| 计划项 | 验收 | 核实结果 |
| --- | --- | --- |
| U-01 正则示例 | PASS | 正则单反斜杠正确，文档原文模型参与真实好/坏数据导入 |
| U-02 模板流所有权 | PASS | 默认关闭、leaveOpen=true 保留、预取消不接管的说明与路径一致 |
| U-03 默认语义 | PASS | 表内列出的按行首错、100 列上限、空行和坐标均与实现一致；另发现表外的默认输入限额问题，见 FIX-002 |
| U-04 Core 注册 | PASS | 默认服务实际注册，TryAdd 保留预注册实现；NPOI 复用该入口 |
| U-05 单表六入口 | PARTIAL | 同步/异步复用既有 Workbook 管线，保留全局错误与源行；默认资源合同受 FIX-002 影响 |
| U-06 统一摘要及行限额 | PARTIAL | 普通行错误、结构、资源和错误上限摘要符合计划；跨 Sheet 诊断存在 FIX-001 |
| U-07 文档、契约、追溯 | PARTIAL | 快照仅增量、逐成员追溯存在；默认输入限制的文档承诺与实际接线需统一 |

## FIX-001：按索引选择时，行数上限错误未使用真实 Sheet 名称

- 严重程度：MEDIUM / P2。
- 处理要求：SHOULD_FIX。
- 状态：OPEN_ACTIONABLE。
- 来源：本次新增行数边界诊断。
- 位置：`src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs:404`。
- 问题：MaxRows 恰在前一个 Sheet 结束时耗尽，下一张表的 ResourceLimit 错误用 `GetSelectorDescription(resolvedSheet.Request.Selector)` 填写 SheetName。按索引选择时得到 `#1`，即使工作簿已解析出真实名称 `Orders`。
- 证据：独立程序 `TestResults/usability-review/probe/Program.cs` 创建 Data、Orders 各一条数据，按索引 0/1 请求且 MaxRows=1。XLS/XLSX × Sync/Async 四组均返回 `actualSheetName="#1"`，源工作簿不存在此名称；失败工作簿 `_ImportErrors` 的 Sheet 列同样为 `#1`。实际 StopReason=ResourceLimit、IsComplete=false 正确，错误在定位信息。
- 影响：依赖错误表名定位 Sheet 的消费者无法找到实际表，用户下载的错误摘要也无法直接对应表签。按名称选择且大小写不同也会回填请求文本而非实际名称。
- 修复目标：已解析存在的 Sheet 使用 `resolvedSheet.Name`；仅不存在时才使用 selector 描述作为定位回退。
- 修复要求：维持错误代码、行列 0、已有成功数据和只报一次的行为；不改变其他限额策略。
- 验证方式：增加 XLS/XLSX × Sync/Async 的按索引行限额边界测试，断言完整错误和失败工作簿摘要的实际表名；保留现有按名称测试。当前 `ImportResultStatusTest` 的 sheet-boundary 仅按名称请求，漏掉该组合。

## FIX-002：默认 128 MiB 输入上限没有接入默认调用

- 严重程度：MEDIUM / P2。
- 处理要求：SHOULD_FIX。
- 状态：OPEN_ACTIONABLE。
- 来源：既有 Workbook 管线遗留，新单表便利入口沿用；不是本次新引入。
- 位置：`src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs:209`、`:330`；新入口 `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelSingleSheetImportExtensions.cs` 的 CreateRequest。
- 问题：用户未显式传入 ResourceLimits 时，请求持有 null。同步/异步输入复制均传 `request.ResourceLimits?.MaxInputBytes`，于是 maxBytes=null，复制器不限制输入；`new ExcelResourceLimits()` 中的 128 MiB 初始化值不会自动执行。
- 证据链：`ExcelImport.cs:41` 的默认资源字段为 null；CreateRequest 原样传递可选 limits；`ExcelWorkbookImportRequest` 原样保存；上述 Copy/CopyAsync 调用读取 nullable 字段；`NpoiStreamCopier.Copy/CopyAsync` 仅在 maxBytes.HasValue 时检查；`ExcelResourceLimits.cs:9` 的初始化值为 128 MiB。架构文档 `docs/architecture/README.md:309` 和 `:604` 明确宣称此限制默认在复制阶段生效。XLSX ZIP 预检另有默认值，但发生在完整输入已缓冲之后，不能替代输入字节限制。
- 影响：快速开始及 `ImportSheet<T>(stream)` 的默认调用会持续复制超过文档默认上限的输入；调用者容易误以为已经具有默认的输入大小保护。XLS 与 XLSX 的输入复制都受影响，不等于已证明它们都能成功解析大文件。
- 修复目标：在 Provider 输入复制前统一取得有效默认 limits；同时保留显式 `new ExcelResourceLimits { MaxInputBytes = null }` 关闭限制的公开语义。不能仅对 nullable MaxInputBytes 使用 `?? 128MiB`，否则会破坏显式关闭。
- 修复要求：Workbook 与单表、Sync 与 Async 使用相同默认策略；文档、测试和异常分类同步；不扩大为 DOM 峰值内存保护承诺。
- 验证方式：针对默认、显式上限、显式 null 三种配置增加职责测试；覆盖 Sync/Async 真实 IO、边界超限及调用方流保持打开，验证 ResourceLimit 异常在解析前发生。复用受控合成流避免制造大型 Office DOM；默认配置需有实际默认阈值验证。本轮为源码调用链确认，未运行 128 MiB 资源探针，不将其记录为容量测量结果。

## 可选增强与外部边界

- OPTIONAL / LOW：快速开始 CSV 编译排除了 NPOI，能够验证样例不直接依赖 Provider。但 `QuickStartConsumerTest.cs:89` 仍从测试进程收集其他可用引用，没有从空目录恢复一个仅引用 Core 的 NuGet 消费者。当前未发现缺失依赖；发布前可补独立包消费者 smoke，避免将源码宿主编译等同于 NuGet 依赖闭包验证。这不计入 OPEN_ACTIONABLE。
- BLOCKED_EXTERNAL：远端 CI 和真实生产资源/并发部署证据仍未验证，本次本地检查不能代替它们；不因这些状态开启自动修复循环。
- 已接受兼容边界：旧 Workbook.IsSuccess 只检查 Errors，手工构造矛盾旧结果时还需检查 IsComplete；当前文档已明确。单表快捷入口不提供所有 Workbook 高级配置，复杂场景仍用 Workbook Request，属于计划内设计。

## 验证证据

1. 读取 plan、execution、当前源码和工作区 diff；暂存区无变更。主代理审查注册、适配、状态及限额管线；luna_worker 独立核对默认值、模板所有权、文档和示例测试。
2. 本轮实际执行 net8 精准测试：CoreServiceRegistrationTest、ExcelSingleSheetImportTest、ImportResultStatusTest、UsabilityDefaultsTest，共 98/98 通过、0 跳过；记录 `TestResults/usability-review/review-targeted.trx`。这证明既有覆盖通过，不抵消上述遗漏。
3. 本轮独立复现 FIX-001 四组均命中。命令：`dotnet run --project TestResults/usability-review/probe/Probe.csproj -c Release`。该程序使用真实 HSSF/XSSF 和失败工作簿写出，不使用 Mock Provider。
4. 复核前轮 TRX：Unit net6/net8 各 924；Integration 各 38；Docs 11，共 1,935。未重跑全量，因为本轮没有生产代码修改。
5. 复核实际 API baseline 与冻结旧基线：每个 TFM Abstractions +22、Core +9、NPOI +0，移除 0；既有最终 api-diff.json 两个 TFM 均为空。没有重新捕获或放宽基线。
6. `git diff --check` 通过；本轮未改生产、正式测试、计划或实施报告。

## 建议顺序

先修 FIX-002，使默认输入保护与公开合同一致；再修 FIX-001 的表名及对应组合测试。两项均可在本地解决，修复后按影响范围验证，无需先重跑重型 benchmark 或完整生产容量矩阵。
