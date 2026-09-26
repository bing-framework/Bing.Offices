<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001
AI_REVIEWED_AT: 2026-09-26T21:46:07.2401247+08:00

# 独立生产代码审查

## 结论

当前复核的 4 个生产 finding 均已关闭，`OPEN_ACTIONABLE=0`（已关闭 `MUST_FIX=3`、`SHOULD_FIX=1`）。F01 映射异常、F03 HSSF 条件格式、F04 报表预检与筛选/表格合并均有当前源码和职责级直接测试证据。实现审查通过；本地 F01-F06 的 Docker/字体/容量、全量构建/测试、package consumer 和双 TFM API compare 均已完成，剩余远程 CI 与正式阈值属于独立外部/审批边界。

## 变更影响分析

| 项目 | 结论 |
| --- | --- |
| ChangedFiles | Abstractions 图片/校验/报表定义；provider-shared 内容与报表预检；SpreadCheetah、NPOI、ClosedXML 写入器；对应合同测试、consumer、Docker/Probe 工具 |
| ChangedProjects | Abstractions、Core/provider-shared、SpreadCheetah、NPOI、ClosedXML、ProviderContract 和各 Provider tests |
| ChangedPublicContracts | `ExcelSheetImageDefinition`、`ExcelDataValidationDefinition`、报表定义和 Builder 增量成员 |
| ChangedRuntimePaths | 映射计划、流式批次写出、图片部件暂存、原生校验、条件格式/表格/筛选/打印写入、取消与文件提交 |
| ChangedProviders | SpreadCheetah、NPOI、ClosedXML；MiniExcel 对新增内容保持预检拒绝 |
| ChangedTFMs | Abstractions netstandard2.0；Provider/test net6.0/net8.0 |
| ChangedBuildPackaging | 新 Provider 包引用、测试 consumer、Docker 验证与容量 probe |
| AffectedDependents | Provider contract、结构化 XML 合同、双 TFM API/package consumer、最终 Docker gate |
| RiskLevel | HIGH |

## Findings

### FIX-F03-001：NPOI XLS 普通条件格式带颜色会在写入阶段抛异常

- `requirement`: `MUST_FIX`
- `status`: `CLOSED`
- `evidence`: 初始候选的问题位于 XLS 预检之后的条件格式颜色分支；当前 [`NpoiReportWriter.cs:90`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:90) 到 [`NpoiReportWriter.cs:166`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:166) 已按工作簿类型选择 HSSF 调色板或 XSSF 颜色，XLS 高级 ColorScale/DataBar/IconSet 仍在 [`NpoiReportWriter.cs:27`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:27) 到 [`NpoiReportWriter.cs:33`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:33) 的 Preflight 阶段拒绝。
- `observed`: 历史复现是 NPOI 2.7.4 的 `HSSFWorkbook` 将 `XSSFColor` 赋给普通条件规则时抛出 `ArgumentException: Only HSSFColor objects are supported`；当前实现已改为 HSSF 调色板路径，未再使用该不兼容 setter。
- `impact`: XLS 请求在已经创建工作簿并枚举数据后才失败，不能兑现“可表达能力成功、不可表达能力在 Preflight 结构化拒绝”的合同；文件目标虽可由外层提交器保全，但通用 Stream 会得到写入失败而非结构化预检结果。
- `fixTarget`: `NpoiReportWriter.ApplyConditionalFormatting` 的 HSSF/XSSF 分支与 `Validate` 的格式能力矩阵；补 XLS 普通 CellValue/Formula 带前景/背景色的直接回归测试。
- `acceptance`: XLS 普通 CellValue/Formula 条件格式带颜色时生成可重新打开的 HSSF 规则；若该版本确实不支持颜色，则在数据枚举和工作簿创建前返回 `BingOfficesUnsupportedFeatureException`，`Stage=Preflight`，且输出为空。XLSX 行为保持不变。
- `verification`: `NpoiReportContractTest.Export_XlsBasicConditionalFormat_ShouldRoundTripPaletteColors` 覆盖 CellValue/Formula、同步/异步、无色/前景色/背景色/双色组合；NPOI Report 合同当前 25 个用例/TFM 通过。
- `resolution`: `CLOSED`。HSSF 使用有限调色板近似颜色，文档不宣称 XLS 真彩等价；XLSX 高级规则保持原有能力和预检边界。

### FIX-F01-001：SpreadCheetah 映射转换异常未按公共异常契约结构化包装

- `requirement`: `MUST_FIX`
- `status`: `CLOSED`
- `evidence`: 初始候选的问题位于 [`SpreadCheetahStreamingExcelExporter.cs:272`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:272) 到 [`SpreadCheetahStreamingExcelExporter.cs:357`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:357) 的 Getter/Converter/类型转换路径；当前同一调用链已由 [`SpreadCheetahStreamingExcelExporter.cs:282`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:282) 到 [`SpreadCheetahStreamingExcelExporter.cs:503`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:503) 分层包装并保留取消、公共异常和致命异常传播。
- `observed`: 历史行为会直接暴露 `InvalidOperationException`、`FormatException` 或 `TargetInvocationException`；当前用户扩展失败使用 `UserExtensionFailed`，内建值/格式转换失败使用 `ExportFailed`，均携带 `Export/SpreadCheetah/Validate` 和 Sheet/行/列/属性位置。取消、已有公共异常和致命异常通过原始对象/EDI 传播。
- `impact`: 违反 F01 的映射转换异常一致性要求，也不符合 [`exceptions-and-observers.md:3`](D:/Bing_Framework/Bing.Offices/docs/excel/exceptions-and-observers.md:3) 至 [`exceptions-and-observers.md:5`](D:/Bing_Framework/Bing.Offices/docs/excel/exceptions-and-observers.md:5) 对不可恢复 Office 失败、用户扩展 InnerException 和位置诊断的公共约定。同步/异步调用方无法依赖结构化错误分类；通用 Stream 可能已经保留部分输出，但这不能替代操作级异常分类。
- `fixTarget`: 为 SpreadCheetah 的固定列/动态列 Getter、Converter、Formatter 和类型转换补充 `BingOfficesExportException` 包装，使用 `Provider=SpreadCheetah`、`Operation=Export`、`Code=UserExtensionFailed`，并携带 `Stage=Validate`（映射/用户扩展）及 Sheet/行/列/属性信息；保留 `BingOfficesException`、`OperationCanceledException` 和致命异常原样传播，不吞掉原始 `InnerException`。
- `acceptance`: 固定列和动态列的 Converter、Getter、格式转换失败均返回稳定的结构化异常，保留原始 InnerException 和位置；同步/异步结果分类一致，文件目标仍由提交器保全，通用 Stream 的部分输出边界继续按既有兼容说明处理。
- `verification`: `SpreadCheetahMappingContractTest` 的固定/动态 Converter、DynamicGetter、属性 Getter、类型转换、Formatter 和取消场景共 17 个用例/TFM，覆盖同步/异步、结构化元数据、InnerException 和同一 OCE 对象传播；当前结果通过。
- `resolution`: `CLOSED`。该 finding 不因通用 Stream 允许部分输出而重新打开；文件目标继续由原子提交器负责保全。

### FIX-F04-001：SpreadCheetah 没有执行公共报表预检，导致晚失败和筛选定义静默丢失

- `requirement`: `MUST_FIX`
- `status`: `CLOSED`
- `evidence`: 初始候选的问题是报表预检晚于部分输出；当前 [`SpreadCheetahStreamingExcelExporter.cs:106`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:106) 到 [`SpreadCheetahStreamingExcelExporter.cs:127`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:127) 先执行共享 `ExcelReportPreflight` 并构造 `preparedSheets`，随后才调用 `Spreadsheet.CreateNewAsync`；[`SpreadCheetahStreamingExcelExporter.cs:155`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:155) 到 [`SpreadCheetahStreamingExcelExporter.cs:169`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:169) 对实际数据区长度执行显式 Write 阶段检查。
- `observed`: 历史行为会晚失败或只消费第一项筛选；当前共享预检覆盖重复名称、样式、范围、筛选冲突，Spread 子集在工作簿创建前完成。表格 `EndRow` 与实际数据区不一致时按明确的 `Write` 阶段结构化失败，单次枚举且不静默忽略范围。
- `impact`: 违反 F04 的“请求级检查与实际写入分离”，跨 Provider 的名称/范围/冲突筛选合同不一致，并可能把用户配置静默降级为另一份工作簿。
- `fixTarget`: 在 `Spreadsheet.CreateNewAsync` 前增加与 `ExcelReportPreflight` 等价的 Spread 子集预检；明确表格范围、样式、名称、筛选数量/冲突和 Provider 不支持项；保证无法表达的请求在枚举和输出前结构化拒绝。
- `acceptance`: 所有静态报表错误均在 `Stage=Preflight` 返回，数据源枚举次数为 0、通用 Stream 长度为 0；多个独立筛选区域必须拒绝而不能只写第一项；有效的同范围 Table+AutoFilter 按统一合并语义成功写出。
- `verification`: `SpreadCheetahStreamingExcelExporterTest` 的无效样式/表格数量/重复名称/越界/筛选冲突/同范围合并覆盖同步异步，共 16 个用例/TFM；`EndRow` 过多/过少的 4 个直接用例通过，断言 `Stage=Write`、单次枚举和输出边界。
- `resolution`: `CLOSED`。静态报表错误在 Preflight，数据区长度错误在实际消费阶段显式失败；两者均不再静默丢失用户定义。

### FIX-F04-002：同范围 Table+AutoFilter 的合并语义仍未贯通所有 Provider

- `requirement`: `SHOULD_FIX`
- `status`: `CLOSED`
- `evidence`: 当前 Provider 文档要求同范围 Table 与 AutoFilter 合并。SpreadCheetah 在 [`SpreadCheetahStreamingExcelExporter.cs:770`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:770) 允许同范围组合，并在 [`SpreadCheetahStreamingExcelExporter.cs:790`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports/SpreadCheetahStreamingExcelExporter.cs:790) 由 Table 拥有筛选；ClosedXML 在 [`ClosedXmlReportWriter.cs:37`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.ClosedXml/Bing/Offices/Exports/ClosedXmlReportWriter.cs:37) 跳过重复筛选；NPOI 在 [`NpoiReportWriter.cs:51`](D:/Bing_Framework/Bing.Offices/src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:51) 到 [`NpoiReportWriter.cs:67`](D:/Bing_Framework/Bing.Offices/src/Bing/Offices.Npoi/Bing/Offices/Exports/NpoiReportWriter.cs:67) 将筛选写在 Table 上并跳过同范围 worksheet filter。
- `impact`: 历史的“Spread 拒绝、ClosedXML 合并、NPOI 重复应用”差异已消除；不同范围重叠仍按共享预检拒绝。
- `fixTarget`: 共享预检和各 Provider writer 的同范围处理；对不同范围重叠继续结构化拒绝。
- `acceptance`: 同范围 Table+AutoFilter 在适用 Provider 中写出一个等价有效的筛选语义；不支持的 Provider 必须在 Preflight 明确说明其限制并同步文档/Profile，不能与全局“合并”文档冲突。补充 OOXML/重新打开结构断言。
- `verification`: SpreadCheetah 同范围 Table+AutoFilter 的 4 个同步/异步与样式别名用例读回 Table 结构；NPOI/ClosedXML 报表直接测试和独立 Profile 读回单一有效 filter。NPOI 的 `ShowTotals` 会扩展 Table 到空汇总行，数据区保持 `A1:B3`，`totalsRowCount=1`，filter 仍绑定数据区；XLS/XLSX 同名 Workbook/Sheet 名称范围读回作用域为 `-1/0/1`。
- `resolution`: `CLOSED`。Provider 差异现在是可表达能力差异而不是静默降级；无效重叠/越界请求仍在 Preflight 结构化拒绝。

## 已验证项目

- 当前线程提供的双 TFM 定向证据：F01 映射异常直接测试 17/TFM；NPOI Report 25/TFM；ClosedXML Report 11/TFM；Spread 报表预检/合并直接测试 16/TFM，均通过。
- 内容合同最终扩展至 83 场景；1900 早期日期边界、NPOI/ClosedXML 1904 模板同步/异步、默认条件格式颜色同步/异步 XML 断言通过。
- 图片路径的 PNG/JPEG 部件、96 DPI EMU 锚点、列表 XML、取消清理和输入/输出流所有权已有职责级覆盖；当前四个 finding 均完成实现闭环，未发现新的同根生产缺陷。
- 共享报表预检此前发现的名称范围逆序/整型溢出、打印区域/重复标题边界、冻结可视起点边界已在当前源码中补齐并由新增 ReportBoundary 合同验证；不再重复列为 finding。
- 最终 Docker 候选 exit 0：`docker-final/docker-smoke.log` 共 352/352（每 TFM 83+25+11+57）；fontless JSON 中 NPOI 结构化失败和 Spread 2821 bytes/readback=true 均符合负向/正向预期；容量矩阵 9/9，最大工作集 81,268,736 bytes。该数据用于报告实测，不推导未批准的性能排名或阈值。
- 最终全量构建 0 errors；final-tests 2,820/2,820，0 fail、0 skip，覆盖 Core/Integration/NPOI/NPOI Integration/MiniExcel/MiniExcel Integration/ClosedXML/ClosedXML Integration/ProviderContract/ExcelDataReader/Spread/Aspose 既有预检及 Docs。Aspose 的 2 个既有预检仅作回归证据，不作为本轮商业 Provider 验收。
- 最终编码检查 160 个文本文件通过，`git diff --check` 通过。
- 最终 package consumer 的 Net6/Net8 restore/build/run 均 exit 0，使用 fresh NuGet cache 和 pinned final feed，真实执行 Spread -> NPOI Sheet 与完整行检查；两次日志均为 `package-consumer-ok`。
- 最终 API compare 使用 `artifacts/open-provider-20260926/api-final/compare/` 下的双 TFM capture，`api-diff.json` 为 `net6.0={}`、`net8.0={}`；未使用 `api-final` 根目录旧文件作为最终证据。

## 证据分类与边界

- `IMPLEMENTATION_STATUS`: `PASS`（`OPEN_ACTIONABLE=0`；4 个 finding 已关闭，其中 `MUST_FIX=3`、`SHOULD_FIX=1`）。
- `TEST_STATUS`: `PASS_FINAL_BUILD_TEST_CONSUMER_API`；final-tests 2,820/2,820、0 fail/skip，package consumer Net6/Net8 和 API compare 双 TFM 均通过。
- `PERFORMANCE_EVIDENCE_STATUS`: `BLOCKED_APPROVAL`；正式容量阈值尚未由负责人批准，只能报告实测。
- `RESOURCE_EVIDENCE_STATUS`: `PASS_FOR_FINAL_DOCKER`；最终 Docker/容量报告已绑定当前候选，容量阈值仍单独受审批状态约束。
- `EXTERNAL_GATE_STATUS`: `PASS_FOR_DOCKER_CONSUMER_API`；最终候选由 `BING_OFFICES_DOCKER_EVIDENCE=<repo>/artifacts/open-provider-20260926/docker-final` 显式指定并 exit 0；远程 CI 仍属于外部边界，不计入生产 `OPEN_ACTIONABLE`。
- `RELEASE_STATUS`: `BLOCKED_EXTERNAL_APPROVAL`；本地 F01-F06 交付证据已齐，剩余为远程 CI 和未批准性能/容量阈值。
- `GOAL_STATUS`: `STOPPED_BLOCKED`；实现 finding 和本地交付门禁已完成，停止自动修复；剩余阻塞仅是外部/审批边界。
- `artifacts/open-provider-20260926` 中来自旧源快照的报表/内容 Docker 失败标记为 `STALE_SOURCE_SNAPSHOT`，不能与当前源码测试拼接为最终 PASS。容量证据只有在相关生产源码范围未变化时才可按 scope 复用。

## 门禁待验证

1. 保留最终 API compare、package consumer 和 Docker/字体/容量证据的显式路径；旧源快照失败仍标记为 `STALE_SOURCE_SNAPSHOT`，不与当前候选拼接。
2. 远程 CI 结果由外部环境单独记录，不替代或否定当前本地 F01-F06 通过证据。
3. 正式性能/容量阈值仍需负责人批准，当前只报告 9/9 实测和最大工作集。
