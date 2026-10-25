<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BO-FUNCTIONAL-GAPS-20261024-001
AI_EXECUTION_FINISHED_AT: 2026-10-24T22:44:10+08:00

# 功能遗漏修复执行记录

> 以下首轮结果保留为历史证据；2026-10-24 复核发现的 FIX-001 及本轮处理见文末“Review 修复记录”。review.md 保持原样。

- 日期：2026-10-24
- 范围：本次功能检查确认的 3 项导入缺陷和 2 项 CI 验证缺口。
- 授权：用户要求解决已发现的遗漏功能点。
- 结果：Implementation PASS；Tests PASS；本地 5 项遗漏全部关闭。Release BLOCKED_EXTERNAL，不以本地结果代替外部证据。
- 状态汇总：OPEN_ACTIONABLE=0，CLOSED=5，BLOCKED_EXTERNAL=2；其余状态为 0。无下一轮自动修复。

## 变更影响分析

| 项目 | 影响 |
| --- | --- |
| ChangedFiles | Core 的 CsvHeaderBinder、CsvEntityImporter、CsvPipelineSupport；NpoiImportSheetExecutor；两个新增直接测试类；测试项目依赖；CI/Docs 工作流；相关说明文档 |
| AffectedProjects | Bing.Offices.Core、Bing.Offices.Npoi、Bing.Offices.Tests；Integration/Docs.Tests 参与最终回归 |
| AffectedPublicContracts | 未增删或改变公共签名；更正重复表头、无表头动态列和字段上限的行为；已更新迁移记录 |
| AffectedRuntimePaths | CSV 同步/异步导入、表头绑定、无表头位置绑定、所有字段长度检查；NPOI 表头绑定 |
| AffectedProviders / TFMs | CSV、NPOI 的 XLS/HSSF 与 XLSX/XSSF；net6.0、net8.0 |
| Build / Packaging | 测试项目增加 PrivateAssets=all 的 coverlet.collector 6.0.4；生产包版本及依赖不变 |
| Resource / Performance | 超长 CSV 字段在解析阶段提前终止；不涉及 Excel DOM 并发、文件提交或临时文件生命周期；不作新的性能或生产容量结论 |

先执行新增边界用例，再执行双 TFM 解决方案回归；覆盖率单独验证实际产物。公共 API 未变化，复用全量回归中的成员快照检查，不重建批准基线。未修改映射缓存键或计划缓存，不启动新的缓存矩阵或 Benchmark。

## Finding 状态

| ID | 状态 | 实现与结果 |
| --- | --- | --- |
| GAP-001 | CLOSED | CSV 和 NPOI 分别记录已绑定固定属性与已声明动态 Key；标题/别名指向同一目标时返回 InvalidHeader，避免覆盖。不同目标仍可正常导入。 |
| GAP-002 | CLOSED | 无表头 CSV 按映射属性顺序展开固定列及动态列，动态列按 Order、Key 排序，与导出一致。动态字典缺少定义时明确报错，覆盖空输入、列数上限和类型转换失败。 |
| GAP-003 | CLOSED | MaxFieldLength 检查所有解码字段，含表头与未映射列；解析器同时限制原始转义长度，超限转换为 ResourceLimit/IsTruncated。保留 UTF-16 长度语义，解析器提前终止时列号为 0。 |
| GAP-004 | CLOSED | 安装覆盖率采集器；CI 与 Docs 工作流在找不到非空 coverage.cobertura.xml 时失败，CI 归档覆盖率文件。 |
| GAP-005 | CLOSED | 将被 Git 忽略目录的无效 git diff 检查替换为生成示例与可编译源码的一致性检查；5 份 Markdown 生成成功，6 个示例一致；篡改临时副本的一个示例后门禁返回失败。 |
| EXT-001 | BLOCKED_EXTERNAL | 当前候选的远端 GitHub Actions 全绿证据，需外部实际运行。 |
| EXT-002 | BLOCKED_EXTERNAL | 目标 2C4G 生产入口的资源预算和 semaphore=1 部署证据，需目标环境验证。 |

外部项沿用 `ai_docs/tasks/BO-RC-20260908-002/BO-RC-20260908-002-final.md` 第 18～20 节边界；本地回归不替代生产或远端 CI。未修改历史独立 review.md。

## 验证证据

### .NET

新增 58 项参数化用例：CsvImportBindingRegressionTest 为 34 项，NpoiImportHeaderBindingTest 为 24 项。

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --filter "FullyQualifiedName~CsvImportBindingRegressionTest|FullyQualifiedName~NpoiImportHeaderBindingTest" --logger "trx;LogFileName=binding-regressions-net8.trx" --results-directory TestResults/audit-fixes
dotnet test Bing.Offices.sln -c Release --no-restore /m:1 --logger trx --results-directory TestResults/audit-fixes/full
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-build --no-restore --filter "FullyQualifiedName~CsvImportBindingRegressionTest|FullyQualifiedName~NpoiImportHeaderBindingTest" --collect:"XPlat Code Coverage" --results-directory TestResults/audit-fixes/coverage
```

| 验证 | 结果 |
| --- | --- |
| 新增用例定向运行 net8.0 | 58 passed / 0 failed / 0 skipped |
| Unit net6.0 | 787 passed / 0 failed / 0 skipped |
| Unit net8.0 | 787 passed / 0 failed / 0 skipped |
| Integration net6.0 | 38 passed / 0 failed / 0 skipped |
| Integration net8.0 | 38 passed / 0 failed / 0 skipped |
| Docs.Tests net8.0 | 10 passed / 0 failed / 0 skipped |
| 全量合计 | 1660 passed / 0 failed / 0 skipped |
| 定向覆盖率采集 | 58 passed，生成可解析的 2,540,890 字节 Cobertura XML；不是全仓覆盖率测量 |

TRX 位于 `TestResults/audit-fixes/full/`；覆盖率文件为 `TestResults/audit-fixes/coverage/c31a8dac-f9de-4f94-9593-35ed3536c868/coverage.cobertura.xml`。两份工作流的 YAML 解析通过；抽取实际 Bash 门禁执行，缺失报告/空报告/真实报告的退出码均为 `1 / 1 / 0`，见 `TestResults/audit-fixes/coverage-gates.json`。

### 文档

运行真实 docgen CLI 的 `all --format md` 与 `docs-check`：成功生成 5 份 Markdown，6 个源码示例全部匹配。在临时目录复制 API 文档并改动一个示例，再调用同一 `run_docs_check`，检测出 ExportWorkbook 不一致，返回码为 1。原始生成文档和示例源码未被负向验证改动。

本地 docgen 虚拟环境出现 Python 3.12 解释器与 cp313 原生扩展不匹配；验证改用 Codex bundled Python 3.12 和 `TestResults/audit-fixes/docgen-python` 中的隔离依赖，没有修改仓库 requirements 或用户全局包。脚本为 `TestResults/audit-fixes/verify-docs.py`；日志及返回码分别见 `docs-gates.log`、`docs-gates.json`（同目录）。本轮只验证受影响的 Markdown 与示例门禁，没有重跑全部 Office/PDF 渲染。

最终 `git diff --check` 通过。窄范围独立交叉核对没有发现新的可执行缺陷。

## 最终生产符号到直接测试的追溯

以下测试均位于 `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj`，并已纳入双 TFM 完整回归。表内类名及方法名可直接检索。

| 最终符号 | 关键行为 | 测试方法 |
| --- | --- | --- |
| CsvHeaderBinder.Bind | 固定目标冲突、动态目标冲突和独立动态别名 | CsvImportBindingRegressionTest.Import_DuplicateFixedTarget_ShouldReturnInvalidHeader；Import_DuplicateDynamicTarget_ShouldReturnInvalidHeader；Import_DistinctDynamicAliases_ShouldPreserveAllValues |
| CsvHeaderBinder.BindByPosition | 动态列排序及固定/动态混合顺序；缺少定义；展开后的列数 | CsvImportBindingRegressionTest.Import_NoHeaderDynamicColumns_ShouldRoundTripInExportOrder；Import_NoHeaderWithoutDynamicDefinitions_ShouldReject；Import_NoHeaderDynamicColumns_ShouldEnforceColumnLimit |
| CsvEntityImporter.Import / ImportAsync 及内部执行路径 | 同步/异步完整实体往返和动态值转换错误 | CsvImportBindingRegressionTest.Import_NoHeaderDynamicColumns_ShouldRoundTripInExportOrder；Import_NoHeaderDynamicColumns_ShouldReportConversionError |
| CsvRecordReader.Read / ReadAsync / ValidateRecord | 所有字段限制；转义、换行和中文边界 | CsvImportBindingRegressionTest.Import_FieldLimit_ShouldCoverEveryField；Import_FieldLimit_ShouldMeasureDecodedCharacters |
| CsvRecordReader.ReadNext / ReadNextAsync | 超长未映射字段提前停止读取；异步不退回同步读取 | CsvImportBindingRegressionTest.Import_FieldLimit_ShouldStopReadingOversizedUnmappedField |
| CsvResourceLimitException / CsvEntityImporter.CreateResourceLimitError | ResourceLimit、IsTruncated、记录/列坐标及已知属性信息 | CsvImportBindingRegressionTest.Import_FieldLimit_ShouldCoverEveryField；Import_FieldLimit_ShouldStopReadingOversizedUnmappedField |
| NpoiImportSheetExecutor.CreateColumns | XLS/XLSX、同步/异步下重复固定/动态绑定的结构化错误 | NpoiImportHeaderBindingTest.Import_DuplicateFixedHeaderTarget_ShouldReturnInvalidHeader；Import_DuplicateDeclaredDynamicHeaderTarget_ShouldReturnInvalidHeader |
| NpoiImportSheetExecutor.CreateColumns | 单一别名、独立动态列、固定属性和动态 Key 分离、未声明动态列正常导入 | NpoiImportHeaderBindingTest.Import_SingleAliasAndDistinctDeclaredDynamicColumns_ShouldImportValues；Import_DifferentUnspecifiedDynamicHeaders_ShouldRemainDistinct |
| 发布公共成员快照 | 现有双 TFM 公共契约保持批准基线 | PublicApiContractTest.PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline；PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot |

未提交、暂存或推送 Git；保留任务开始前已有 CI 工作，并未改动原有未跟踪的 commit-lint.yml。

## 剩余 TODO

- 在远端运行当前候选 CI，记录双 TFM、package consumer 和发布门禁结果。
- 在目标 2C4G 生产环境确认入口并行度为 1，并补齐资源预算与部署证明。

这两项为外部验证，未通过自动修改源码或重复本地测试伪造关闭。

## Review 修复记录

### Round 1：FIX-001（P2 属性身份误判）

- 模式：REVIEW_FIX；Fix Scope：recommended。
- 输入：本任务 review.md，原始状态 NEEDS_FIX；本轮仅处理其唯一 SHOULD_FIX 项 FIX-001。
- 授权：用户明确要求修复 P2 回归；本任务没有 plan.md，范围以用户指令和 FIX-001 为准。
- 执行状态：COMPLETED。未完成 FIX：无。原 review.md 保留为独立复核证据，不将其状态自行改写为 PASS。

#### 变更影响分析

| 项目 | 本轮范围 |
| --- | --- |
| ChangedFiles | NpoiImportSheetExecutor.cs、NpoiImportHeaderBindingTest.cs、本 execution.md |
| ChangedProjects | Bing.Offices.Npoi、Bing.Offices.Tests |
| ChangedPublicContracts | 无公共签名变化，无 API baseline 重建 |
| ChangedRuntimePaths | CreateColumns 的固定属性目标身份去重 |
| ChangedProviders / TFMs | NPOI XLS/HSSF、XLSX/XSSF；net6.0、net8.0 |
| ChangedBuildPackaging / BenchmarkHarness | 无 |
| ChangedDocsOnly | 否；执行记录补充验证与追溯 |
| AffectedDependents | Excel 同步/异步导入路径 |
| RiskLevel | LOW；单个 HashSet 比较器更正，保留表头匹配与动态列处理 |

#### 根因与修复

旧代码按 HeaderComparison 为固定属性名称集合选择比较器，默认 OrdinalIgnoreCase 把 CLR 的 Name/name 合并。现改为 `new HashSet<string>(StringComparer.Ordinal)`，明确固定属性身份区分大小写；表头仍遵从请求中的比较规则。不同标题映射到 Name/name 可分别赋值；同一属性被标题、别名或属性名重复绑定时仍返回 InvalidHeader。

#### 直接测试与结果

新增 `NpoiImportHeaderBindingTest.Import_CaseDistinctFixedProperties_ShouldBindBothValues` 共 8 项：XLS/XLSX × 同步/异步 × OrdinalIgnoreCase/Ordinal。忽略大小写模式使用 lower-case 表头 upper/lower，验证表头仍能匹配属性标题 Upper/Lower；断言完整实体值、空错误、SourceRows=[1]、调用方输入流保持可读。

| 验证 | 结果 | 证据 |
| --- | --- | --- |
| 修改生产代码前，新测试 net8.0 | 4 failed（OrdinalIgnoreCase）、4 passed（Ordinal），成功复现 | TestResults/p2-fix/before/case-distinct-before.trx |
| 修改后，整个 NpoiImportHeaderBindingTest net6.0 / net8.0 | 各 32 passed / 0 failed / 0 skipped，含新 8 项及原 24 项 | 定向命令输出；after/p2-header-binding.trx 保存最后的 net8.0 结果 |
| Unit net6.0 | 795 passed / 0 failed / 0 skipped | TestResults/p2-fix/unit/jianx_DESKTOP-8CUU71G_2026-10-24_22_43_32_net6.0.trx |
| Unit net8.0 | 795 passed / 0 failed / 0 skipped | TestResults/p2-fix/unit/jianx_DESKTOP-8CUU71G_2026-10-24_22_43_36_net8.0.trx |
| L0 | 双框架编译、git diff --check 通过 | 定向测试构建输出与最终差异检查 |

双框架单元测试合计 1590 项通过。定向 TRX 因使用固定文件名由 net8.0 覆盖 net6.0，因此最终双 TFM 的完整证据以两个独立 Unit TRX 为准；二者均确认新增 8 项全部通过。

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release --no-restore /m:1 --filter "FullyQualifiedName~NpoiImportHeaderBindingTest" --logger "trx;LogFileName=p2-header-binding.trx" --results-directory TestResults/p2-fix/after
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release --no-build --no-restore /m:1 --logger trx --results-directory TestResults/p2-fix/unit
```

#### 最终符号追溯增补

| 生产符号 | 行为 | 测试项目及方法 |
| --- | --- | --- |
| NpoiImportSheetExecutor.CreateColumns | Name/name 属于不同目标，表头比较策略仍生效 | Bing.Offices.Tests / NpoiImportHeaderBindingTest.Import_CaseDistinctFixedProperties_ShouldBindBothValues |
| NpoiImportSheetExecutor.CreateColumns | 真正重复固定目标和动态目标仍返回 InvalidHeader | Bing.Offices.Tests / NpoiImportHeaderBindingTest.Import_DuplicateFixedHeaderTarget_ShouldReturnInvalidHeader；Import_DuplicateDeclaredDynamicHeaderTarget_ShouldReturnInvalidHeader |

#### 收口与状态

- 本轮 SHOULD_FIX：1；执行完成：1；PARTIAL/BLOCKED/FAILED：0；执行侧 OPEN_ACTIONABLE：0，FIX-001 已由前后对照测试验证修复。
- Implementation：1/1；Tests：PASS；Release Gates：0/2，远端 CI 和目标生产环境证明仍为 BLOCKED_EXTERNAL。
- Validation Run：L0、L1、L2 完成；L3 全解决方案、L4 Benchmark、L5 生产资源矩阵不适用本轮改动，未重复运行。
- Reused Evidence：未变化的 CSV、覆盖率工作流、文档门禁等原有证据；不将历史全量回归数当成本轮新运行结果。
- Blocked Approval：无；Accepted Limitations / Verified Boundaries / Deferred：无新增。
- No-Progress Check：CHANGED；Next Action：STOP。本轮执行完成，正式复核结论由后续独立 Review 判定。
- 流程脚本：已尝试 task-state.mjs review-fix，因没有 plan.md 返回 1，且在写入运行状态前退出。没有为登记脚本虚构计划；本轮时间与终态直接记录于 execution.md。task-finish 未执行，因为没有可对应的本轮运行登记；未改写其他任务的运行状态、未发送外部通知。
- 未 git add / commit / push；未修改 review.md。
