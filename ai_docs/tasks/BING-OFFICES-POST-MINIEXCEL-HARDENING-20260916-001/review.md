<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001
AI_REVIEWED_AT: 2026-09-17T07:32:26.2160191+08:00

# 第六轮独立代码审查

## 结论与范围

结论：`PASS_WITH_ISSUES`。Round 5 的 `FIX-013` 已解决，未发现新的未解决 MUST_FIX/SHOULD_FIX；保留 `FIX-010` 和 `FIX-014` 两项 LOW OPTIONAL。这是当前代码整改的审查结论，不是正式发布批准；计划 H-800 仍为 `PARTIAL/NOT_VERIFIABLE`，不得据此将任务标记为 `COMPLETED`。

本轮对照完整 plan.md、execution.md（含 Round 5）、前轮 review、工作区/暂存区差异、实际调用链、测试、配置和产物。暂存区无变更；分支为 `feat/miniexcel-provider`。保留任务开始前的 dirty worktree，不把既有 version.dev.props diff 误归因于本轮。仅更新本报告及根 artifacts 下验证输出，未修改生产代码、测试、plan.md、execution.md、baseline、版本文件或任务运行状态；未 commit/push/tag/PR/发布。

## 本轮复核

- 当前候选源码身份与 Round 5 baseline 一致：使用当前 repository、Release 和 Round 5 packages 重新运行 API compare，双 TFM exit 0；`artifacts/review-round6/api-compare/api-diff.json` 两个 TFM 均为空。未发现新的 MUST_FIX/SHOULD_FIX。
- 只读检查共享行处理和 raw reader：`HeaderRowIndex + 2` 起点保持，FIX-013 继续 CLOSED；当前 benchmark 表头、缺失 XML 注释和符号表缺口未变，FIX-010/FIX-014 保留 OPTIONAL，不重复升级或开启修复循环。
- 本轮版本文件哈希匹配下文起点，嵌套产物扫描为 0，git diff --check exit 0（仅 LF/CRLF 提示）；资源报告审批仍 BLOCKED。
- 下文功能测试、Consumer 原始字节核验为第五轮及 Round 5 Executor 的已核验历史证据。本轮未重新执行测试、build、pack、Consumer、100K 或资源矩阵，未将历史结果记为本轮运行。发布门禁仍 PARTIAL。

## FIX-013 验收：CLOSED

- Requirement：H-200/H-201/H-500/H-501，原始日期 serial、SourceRows 和公共上下文必须对应当前输入物理行。
- Evidence：`MiniExcelExcelImporter.cs:347` 将起点改为 `HeaderRowIndex + 2`；`skipped` 仍为 `DataRowStartIndex - HeaderRowIndex - 1`，每个枚举项只推进一次物理行。`CreateConfiguration` 明确 `IgnoreEmptyRows=false`。同步 `Query` 和异步 `QueryAsync` 均进入同一 `ImportTypedSheetRows`，未复制第二套行号算法。
- 调用链：固定/动态转换和验证将当前 `rowNumber` 传入上下文；`CreateRawDateCell` 按当前物理行列读取 XML serial；成功行 `SourceRows` 使用 `rowNumber - 1`，保留零基来源行和一基公共上下文合同。
- Direct tests：`DataRowStartIndex_ShouldAlignRawDateSerialsAcrossProviders` 四组真实 XSSFWorkbook fixture 覆盖跳过一行/多行、表头 0/2、1900/1904、固定 DateTime/显式 +08:00 DateTimeOffset、动态 DateTime，同一文件/request 比较两个 Provider；预期日期独立列出，不从待测解析器生成。完整断言日期数组、空错误集合、SourceRows、converter 和 validation 行号。`DataRowStartIndexAsync_ShouldAlignRawDateSerials` 两组覆盖非零表头、多行跳过及两日期系统，共享路径异步回归通过。
- 默认布局、0/1/59/60/61、负小数、DTO 默认导出、物理列范围回归仍通过。符号表已新增共享行处理及 raw reader 的测试映射。
- 验证：第五轮重新执行 net6.0/net8.0 MiniExcelProviderTest + ExcelDateParserTest，各 `43/43 PASS`，包含六个新增 theory case；Round 5 保存的真实 probe 双 TFM 中 NPOI/MiniExcel 均输出正确日期及 `rows=2|3`。本轮调用链复核未发现原错位再现。

## OPTIONAL 项

### FIX-010：Benchmark allocated median 名称不准确

- Severity：LOW
- Fix Level：OPTIONAL
- Status：DEFERRED
- Requirement：H-401/H-701，统计口径可追溯。
- Problem/Evidence：`benchmark-report.md:35-38` 的“中位 allocated”取 elapsed 中位样本的 allocation，而不是 allocation 字段自身中位数；延续前轮低风险问题，不升级为阻断项。
- Impact：表头可能误导统计口径，但不改变原始样本，也不能据此宣称优化收益。
- Goal/Requirements：改为“elapsed 中位样本 allocated”，或逐字段独立计算中位数；保留原始数据，禁止填造性能结果。
- Validation：根据 `artifacts/review-fix/provider-comparison-100k-final.json` 重算并核对表格。

### FIX-014：修改范围内中文注释和最终符号追溯仍不完整

- Severity：LOW
- Fix Level：OPTIONAL
- Status：OPEN
- Requirement：H-700 和最终生产符号到测试方法的追溯规范。
- Problem/Evidence：`MiniExcelRawDateSerialReader.cs:13-15,74,86,125,147,170,172` 的常量/辅助方法仍缺独立中文 XML 注释；`MiniExcelExcelImporter.cs:485` 的原始日期查表辅助同样缺少 XML。`MiniExcelValueAdapter` 相关转换成员缺完整中文 XML 契约；不能因 private/internal 或 docs build 通过就称技能已全面满足。`symbol-test-map.md` 虽已加入 raw reader 和跳过行测试，但未单列 `ConvertTo/ConvertDynamicTo` 的 DTO 保 offset 测试、Core 新负数日期测试及列范围测试，部分证据计数仍为历史 25/25。
- Impact：维护和验收追溯不完整；当前功能有真实通过测试，本项没有发现运行时错误，不作为新的功能阻断。
- Goal/Requirements：仅治理本任务实际修改成员的准确中文 XML；补全最终符号到 `DateTimeOffsetExport_ShouldPreserveOffsetAsInvariantRoundTripText`、`DateTimeOffsetExportAsync_ShouldPreserveOffsetForDynamicColumns`、`TryParse_NegativeSerials_ShouldUseLinearExcelDays`、`ReadColumnRange_ShouldReportAbsolutePhysicalColumnIndexAcrossProviders` 的映射，标明测试项目和当前 TFM 证据。不得修改行为或公开签名。
- Validation：Docs/XML build、注释签名检查、最终映射与 TRX 方法名核对。在处理前 H-700/H-701 只可 PARTIAL，不能宣称完全符合技能。

## 前轮整改验收

| 项目 | 结果 |
| --- | --- |
| FIX-001/002/003/005/006 | 保留前轮已验收结果；配置校验门控、offset 政策、动态列上下文、取消和同负载对照未发现新的主路径问题 |
| FIX-004 | 默认及跳过布局 raw XML serial 均进入 Core；1900/1904 和负数合同目标测试通过 |
| FIX-007/008 | 包 DLL 独立身份校验存在；Round 5 包/cache/Release/Consumer 原始字节重新核验一致 |
| FIX-009/011 | 物理列号计入范围起点；CI 哈希比较统一大写且冻结版本未变；外部 CI 仍未执行 |
| FIX-012 | 默认固定/动态 DTO 使用 invariant O 文本；显式 converter/formatter 优先；现有 unsupported 门禁未扩大 |
| FIX-013 | CLOSED；共享同步/异步物理行起点修复和直接合同通过 |

## 计划覆盖与发布门禁

| 阶段 | 结果与边界 |
| --- | --- |
| H-000 | PASS；版本哈希匹配任务起点，保留既有 dirty worktree |
| H-001/H-100 | DEVIATED_OK；Core 保持 netstandard2.0；历史错误未复现，采用共享源文件链接到两个 Provider，不伪造兼容修复 |
| H-101 | PASS；Core 生产 IVT 为 0，测试友元保留，无新增 SPI |
| H-200/H-201 | PASS；当前日期、跳过行和行列上下文直接回归通过 |
| H-300 | PARTIAL；取消/流所有权/旧目标保护/清理测试存在；共享/NPOI 36 场景矩阵不等于 MiniExcel 新索引预算验证，资源审批仍 BLOCKED |
| H-400/H-401 | PARTIAL；历史 100K 同 workload 样本存在，关系仍 O(P×C)；新增无条件数值索引后的最终 100K 未重跑，最终性能 NOT_VERIFIABLE；before 不可重放 |
| H-500/H-501 | PASS/PARTIAL；功能职责和跨 Provider 合同通过，新增跳过行测试已映射；完整最终追溯仍有 FIX-014 |
| H-600 | PASS；根 artifacts 集中输出，src/tests/benchmarks 嵌套产物目录计数为 0 |
| H-601 | PASS；四个 Round 5 包同版本，双 TFM Consumer 来源和 DLL 字节一致，MiniExcel 无 NPOI 依赖 |
| H-602 | PARTIAL；当前双 TFM API compare 通过、成员 diff 为空；既有 approval 元数据不等于新增维护者审批，外部 CI 未执行 |
| H-700/H-701 | PARTIAL；新增行处理和 fixture 注释存在，主能力文档同步；XML 注释/最终追溯尚有 OPTIONAL 缺口，历史 benchmark 不能代表最终新 reader 性能 |
| H-800 | PARTIAL/NOT_VERIFIABLE；代码整改无未解决 MUST_FIX/SHOULD_FIX，但发布预算、最终性能及外部环境门禁未全部完成 |

## 第五轮验证证据与限制

- 第五轮重新运行：`dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0/net8.0 --no-build --no-restore --filter "FullyQualifiedName~MiniExcelProviderTest|FullyQualifiedName~ExcelDateParserTest"`，各 43/43 PASS；TRX 为 `artifacts/review-round5/net6.0/targeted.trx` 和 `net8.0/targeted.trx`。
- 第五轮重新运行 API compare：`output/release`、`build/api-snapshot-baseline.json`、`artifacts/review-fix-round5/packages`、repository 源码，exit 0；`artifacts/review-round5/api-compare/api-diff.json` 的 net6.0/net8.0 均为空。
- 第五轮委托 luna_worker 只读重算：8 行包/cache/Consumer/Release DLL SHA256 全部一致，四份 feed/cache nupkg 原始哈希一致；四包 nuspec 都为 2.0.0，MiniExcel dependencies 无 NPOI；八份 cache metadata source 都指向 Round 5 packages；net6/net8 run log 均含 package-consumer-ok/2.0.0 和 npoiExtensions=ok。
- 第五轮核验的 Round 5 TRX Counters 与实际结果一致：Unit 各 771/771、Integration 各 39/39、Docs net8 10/10，共 1,630 条 Passed；目标各 34/34 是额外运行记录，不作为额外不重复用例。第五轮未重新执行全量、pack 或 Consumer restore/build/run。
- 版本哈希：version.props=`77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1`；version.dev.props=`BABFE7703A15E1C11F46B45D34BC59D7913ECB1E3DCFDA901E72F4D60A8D134D`，匹配起点。
- 第五轮 src/tests/benchmarks 的 artifacts/TestResults/BenchmarkDotNet.Artifacts/packages 嵌套目录扫描为 0；git diff --check exit 0，仅 LF/CRLF 提示。
- 第五轮未重新执行 Solution build、identity self-test、100K、资源矩阵或外部 CI。500K/1M、生产 2 CPU/4 GiB、外部 CI 继续 NOT_VERIFIED；资源批准继续 BLOCKED。最终 reader 全量 numeric 索引带来的资源/性能边界不可从旧样本外推。

## 后续

FIX-013 的代码整改可结束，无需继续默认 recommended 修复循环。OPTIONAL 可由维护者单独安排；发布前须补最终 100K/资源验证与审批，继续保留未执行环境的 NOT_VERIFIED。当前 Review 通过不替代计划发布门禁，不修改 execution.md 的 PARTIAL 状态。
