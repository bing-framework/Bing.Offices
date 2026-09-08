<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: BO-RC-20260907-001
AI_EXECUTION_FINISHED_AT: 2026-09-08T11:48:01+08:00

# 实施执行报告

## 执行结论

本轮已完成可在当前环境安全执行的核心实现与本地验证，状态为 `PARTIAL / NO-GO`。双 TFM、真实 Async API、NPOI namespace/API 分层、直接测试、Integration、Docs、双 TFM PackageReference-only consumer，以及带 warmup/重复测量和结构化预算失败记录的 100K staging 矩阵已落地。正式 API baseline 审批、性能/资源预算批准、跨平台 runner 和下一轮独立 Review 仍未完成，不能把任务标记为 Go 或 COMPLETED。

## 任务信息

- Task ID：`BO-RC-20260907-001`
- 执行器：Codex
- Branch：`master`
- 基线 commit：`4ecfe3fa7acea6b9d22d00e52e102c5ed1cf31bc`
- SQL 元数据规则：不适用，本轮未修改 `Bing.Data.Sql`
- 未自动 git commit、git push、PR、tag 或 NuGet publish

## 计划执行情况

| Phase | 状态 | 结果 |
| --- | --- | --- |
| P0 基线/合同 | COMPLETED | 新基线、需求矩阵、API 分类和 Async 合同已生成 |
| P1 net6/net8 | COMPLETED locally | NPOI/Unit/Integration 双 TFM build/test；CI 双 TFM |
| P2 API/namespace | COMPLETED locally | concrete provider public；NPOI extensions 迁移到 `Bing.Offices.Npoi.Extensions` |
| P3 Async 基础设施 | COMPLETED | `CommitAsync`、`CopyAsync`、取消/清理路径已接入 |
| P4 CSV Async | COMPLETED locally | CsvHelper async Parser/Writer、文件导出和直接测试 |
| P5 Excel Async | COMPLETED locally | 输入异步复制、输出 staging、Failure Workbook 异步复制；DOM 同步边界有文档 |
| P6 测试/Consumer | COMPLETED locally | Unit/Integration/Docs、net6/net8 PackageReference-only consumer 通过 |
| P7 Benchmark/Resource | PARTIAL | workload 修复、CSV Sync/Async Dry smoke、ResourceProbe 完成；正式预算未批准 |
| P8 文档/发布门禁 | PARTIAL | README/docs/reports 完成；API 审批和独立 Review 待外部输入 |

## 已完成事项

1. `framework.props`、NPOI、Unit、Integration 恢复 `net6.0;net8.0`；Abstractions/Core 保持 `netstandard2.0`；CI 安装并测试两个 TFM。
2. `IExcelImporter`、`IExcelExporter`、`ICsvImporter`、`ICsvExporter` 增加 Async 成员；新增 Stream async extensions。
3. `IFileExportCommitter.CommitAsync` 与 `AtomicFileCommitter` 使用 async FileStream/FlushAsync，保留原子 replace/move、异常分类和临时文件清理合同。
4. `NpoiStreamCopier.CopyAsync` 使用 `ReadAsync/WriteAsync` 和 CancellationToken。
5. CSV 使用 CsvHelper `ReadAsync`、`NextRecordAsync`、`FlushAsync`；映射/校验/转换结果保持同步管线合同。
6. Excel Async 复用同步 NPOI DOM 业务管线，外围输入复制、目标复制和 Failure Workbook staging 使用真实异步 IO；未使用 Task.Run/Result/Wait。
7. `NpoiExcelImporter`/`NpoiExcelExporter` 改为 public sealed Provider User API；NPOI 扩展 namespace 收敛到 `Bing.Offices.Npoi.Extensions`，所有测试/Docs/Benchmark/Consumer 同步迁移。
8. 新增 AsyncOnly Read/Write、Failure Workbook Async、pre/mid cancellation 测试；DefaultFileExportCommitter 增加 9 个同步/异步职责级场景；更新 fake exporter/committer、API 分类测试和旧反射测试。
9. Benchmark 修复 Unique 输入预生成、Dynamic FactoryCreation 分离、移除普通 Benchmark PeakWorkingSet；增加 CSV/Excel Sync/Async benchmark 方法。
10. README、Excel README、日期/NPOI extension docs 和 `async-io.md` 同步 Runtime/Async/namespace/迁移合同。

## 部分/未完成事项

- `build/api-snapshot-baseline.json` 仍缺 `approvedBy`/`approvedAt`；candidate net6/net8 已生成，正式 compare 按设计阻塞。
- 性能/资源正式阈值、before/candidate 完整矩阵、P95/P99 和 staging 方案审批未完成；完整 ResourceProbe 已完成，但仍无维护者批准阈值。
- Linux/macOS runner 未在本环境执行；CI 已更新为安装并测试 net6/net8，但尚未取得远端运行结果。
- 独立 Review 尚未由 Reviewer 生成；本报告不能替代独立审查。

## 修改文件

主要生产文件：

- `framework.props`、`.github/workflows/ci.yml`、`build/ApiSnapshot/Program.cs`
- Abstractions 四个 importer/exporter 接口和 `IFileExportCommitter`
- Core CSV pipeline、CSV/Excel Stream extensions、AtomicFileCommitter/DefaultFileExportCommitter
- NPOI importer/exporter、NpoiStreamCopier、Failure Workbook writer、NPOI Extensions namespace

测试/工具/文档：

- 双 TFM csproj、AsyncPipelineTest、既有 API/Stream/fake 测试、DocsConsumer/Integration using
- Benchmark/ResourceProbe 相关源文件
- README、`docs/excel/*.md`、任务目录下 baseline/matrix/contracts/reports/artifacts/PackageConsumer

## API/数据/配置变化

- Public API 增加七组 Async 成员及 Stream async extensions；第三方接口实现需要增加对应成员。
- `NpoiExcelImporter`/`NpoiExcelExporter` 从 internal 变为 public sealed；NPOI extension namespace 迁移是 Breaking Change。
- NPOI 包资产从单 net8 扩展为 net6/net8；Abstractions/Core 继续 netstandard2.0。
- 未修改数据库 schema、生产数据或 `Bing.Data.Sql`。

## 测试结果

| 验证 | 结果 |
| --- | --- |
| Release solution build `/m:1` | 0 warning / 0 error |
| Unit net6 | 509 passed / 1 API approval blocked / 0 skipped / 510 total |
| Unit net8 | 509 passed / 1 API approval blocked / 0 skipped / 510 total |
| Integration net6 | 15/15 passed |
| Integration net8 | 15/15 passed |
| Docs net8 | 10/10 passed |
| Async/CommitAsync direct tests | net6/net8 通过；net8 AsyncPipeline 8/8、Committer 9/9 |
| Package Consumer net6 | `package-consumer-ok`, real nupkg, isolated final cache |
| Package Consumer net8 | `package-consumer-ok`, real nupkg, isolated final cache |
| API candidate | net6/net8 canonical members identical；formal approval blocked |
| Benchmark smoke | CSV Sync/Async 1K/10K/100K/1M Dry 结果已保存 |
| ResourceProbe | 16/16 passed；最大 PeakWorkingSet 142,835,712 B、最大 LOH 57,147,216 B |

## Build/Typecheck/Lint/Format

- `dotnet build Bing.Offices.sln -c Release --no-restore /m:1`：通过，0 warning/0 error。
- `git diff --check`：通过；Git 仅提示两个历史/生成文件的 CRLF/LF 转换，不是 whitespace error。
- 仓库未配置独立 lint/formatter；未伪造 PASS。

## 计划偏差

1. Excel ExportAsync 采用 NPOI 同步 DOM staging + 异步 CopyToAsync；这是受 NPOI API 限制的诚实边界，完整 A/B/C 资源预算留给维护者批准阶段。
2. CSV Async 使用共享校验/转换辅助但异步 record loop 与同步 loop 各有编排代码；行为由 AsyncOnly/parity 测试约束，后续可在独立重构任务进一步消除重复。
3. CI pack 去掉 `--no-build`，避免旧 bin 产物污染最终包；独立 consumer 使用任务专属 feed/cache。

## 基线问题

- 初始 lockfile 漂移与 NuGet TLS/凭证问题已通过受控网络 force-evaluate 还原；报告保留原始阻塞。
- .NET 6 EOL 由 SDK 产生 `NETSDK1138`，属于明确支持风险，不隐藏或降低测试门槛。

## 已知问题

- API baseline 旧 hash 与当前 candidate 不同：Abstractions 16 added/5 removed、Core 12 added/2 removed、NPOI 10 added；需维护者批准后更新正式 baseline。
- CSV 1M Export 约 14.5 GB allocated、Import 约 1.98 GB；不能宣称零 GC 或低内存。
- 本地未执行 Linux/macOS；未执行完整正式 P95/P99 与批准预算矩阵。

## 风险与回归关注点

- 新增接口成员会影响第三方自定义实现；迁移文档和版本说明必须随发布提供。
- Excel Async 的同步 NPOI 阶段仍占用调用线程；大文件 staging 的内存/磁盘策略必须依赖固定环境资源证据。
- namespace 迁移没有长期旧入口，调用方需更新 using。

## Reviewer 注意事项

1. 先审 `artifacts/api-compare-final/api-diff.json` 与 `api-candidate.json`，确认新增 Async、Provider concrete、namespace move 的批准范围。
2. 审 `artifacts/reports/benchmark-report.md` 和 resource report；当前 Dry/small probe 不能替代预算批准。
3. 独立 Reviewer 需检查 AsyncOnly 测试是否覆盖真实 IO、Failure Workbook staging、异常 Observer 单次/同实例和双 TFM API parity。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BO-RC-20260907-001/review.md`
- 执行器：Codex

#### FIX-001

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityImporter.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityExporter.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：CsvHelper 的异步 Parser/Writer API 没有 CancellationToken 参数，原实现只有循环边界检查。
- 修复：新增不拥有底层流的 `CsvCancellationStream`，将原始令牌绑定到 StreamReader/StreamWriter 的 ReadAsync/WriteAsync/FlushAsync；公共 Async 入口规范化派生 TaskCanceledException。
- 验证：
  - AsyncPipeline net6/net8 专项测试：通过
  - Blocking AsyncOnly Read/Write 测试记录令牌并在取消时结束：通过

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelExporter.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：Async Import 重复复制输入；Export 和 Failure Workbook 使用完整 MemoryStream staging。
- 修复：引入 `ImportBufferedCore` 避免 Async Import 二次全量复制；Excel Export/Failure Workbook Async staging 改为同机临时文件并用 CopyToAsync 输出；补充 staging 清理和取消测试。
- 验证：
  - Excel Async blocking read/write、Failure Workbook cancellation：通过
  - ResourceProbe review1：16/16 mapping/unique 场景通过
  - BenchmarkDotNet Excel Dry：因自动生成项目 NuGet TLS/凭证失败，未取得测量结果
- 未完成原因：正式 A/B/C staging、Excel/Failure 大文件、并发和 P95/P99 资源预算仍需固定 runner 与维护者阈值批准。

#### FIX-003

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改文件：
  - `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
  - `.github/workflows/ci.yml`
  - `ai_docs/tasks/BO-RC-20260907-001/api-classification.md`
- 根因：Unit API test 固定读取 net8 baseline，CI 没有显式执行 CLI compare，candidate 删除成员没有迁移台账。
- 修复：测试按 `NET6_0`/net8 当前 TFM 选择 baseline；CI 增加 ApiSnapshot compare；记录 `ExcelMappingDiagnostic` 类型/成员及两个 diagnostics overload 的待审批删除项和迁移要求。
- 验证：
  - net6/net8 API test 均能进入审批门禁：仍因 `approvedBy/approvedAt` 为空而阻塞
  - CLI compare 明确报告 net6 baseline 缺失、审批字段为空和 net8 hash diff
- 未完成原因：维护者必须提供真实 API 审批并更新正式 net6/net8 baseline，执行器不得自批。

#### FIX-004

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs`
  - `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 修复：IAtomicFileSystem 增加 FlushAsync；FileStream 先异步 flush，再执行与同步路径一致的 flush-to-disk；测试适配器同步实现该合同。
- 验证：Release Build 通过；CommitAsync 专项测试 net6/net8 通过。

#### FIX-005

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED implementation / BLOCKED measurement
- 修改文件：
  - `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
- 修复：DynamicPlanBuildCold 使用 GlobalSetup 创建的独立 factory，测量循环只构建不同 tenant 的 plan；FactoryCreation 保持独立。
- 验证：Release Build 通过；BenchmarkDotNet 正式 Dry 启动失败于 NuGet TLS/凭证，未伪造结果。

#### FIX-006

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 修复：补齐 CSV/Excel blocking cancellation、Failure Workbook cancellation、AsyncOnly token 记录，以及 mapping/validation/converter/dynamic column parity。
- 验证：
  - Unit net6：509 passed / 1 API approval blocked / 0 skipped / 510 total
  - Unit net8：509 passed / 1 API approval blocked / 0 skipped / 510 total

#### FIX-007

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/package-consumer/Program.cs`
- 修复：consumer 实际调用 `WorkbookExtensions`/`RowExtensions`，并使用 direct importer 完成导入断言后才输出 `npoiExtensions=ok`。
- 验证：net6/net8 PackageReference-only consumer 均运行成功；net6 仅有预期 EOL warning。

#### FIX-008

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：SKIPPED
- 原因：默认 `fixScope=recommended` 不处理 OPTIONAL；XML docs 仍保留在下一轮独立改进范围。

### Round 1 汇总

- MUST_FIX：FIX-001、FIX-006 已完成；FIX-002、FIX-003 代码部分完成但外部证据/审批阻塞。
- SHOULD_FIX：FIX-004、FIX-005、FIX-007 已完成；FIX-005 正式测量受 NuGet TLS/凭证阻塞。
- PARTIAL：FIX-002、FIX-003、正式 Benchmark/资源预算。
- BLOCKED：维护者 API baseline 审批、net6 baseline、正式性能/资源阈值、可用 Linux runner；BenchmarkDotNet 自动生成项目 restore 的 NuGet TLS/凭证。
- FAILED：无。
- 回归验证：Build 0 warning/error；Integration net6/net8 各 15/15；Docs 10/10；ResourceProbe review1 16/16；Unit 双 TFM 各 509 pass + 1 审批阻塞。
- 下一步：维护者完成 API/资源审批并提供可用 runner 后，重新执行正式 Benchmark/compare，再进行独立 Review。

## Review 修复记录

### Round 2

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BO-RC-20260907-001/review.md`
- 执行器：Codex

#### FIX-009

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelExporter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs`
- 根因：Excel 与 Failure Workbook staging 清理失败此前可能被静默吞掉；正式 Memory/temp/hybrid 资源对照和维护者预算审批仍缺失。
- 修复：导出/导入 staging 路径统一保存主异常；清理失败附加到主异常 `Data`，成功路径清理失败显式抛出；修正取消异常规范化后诊断挂错实例的问题。
- 验证：
  - 代码静态检查与 `git diff --check`：PASS
  - Excel Export/ExportAsync 100K smoke：PASS，原始结果 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-report.csv`
  - Failure Workbook 100K smoke：PASS，原始结果 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.FailureWorkbookBenchmarks-report.csv`
  - Memory/temp/hybrid A/B/C、Failure 双 DOM、模板/图片/样式、并发 1/4/16/64 的正式对照与批准预算：BLOCKED（当前 Benchmark 未覆盖策略对照，仍需固定 runner 与维护者门禁）

#### FIX-010

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：缺少 Excel `ExportToFileAsync` 文件级 pre/mid cancellation、目标保留和 staging 集合断言。
- 修复：新增可注入提交器驱动的 Excel 文件级 pre-cancel/mid-write-cancel 测试；增加目标内容保留和 `bing-offices-excel-async-*`、`bing-offices-failure-async-*` staging 集合无泄漏断言。
- 验证：net6/net8 AsyncPipeline 各 15/15 PASS；完整 Unit 双 TFM 均包含该路径。

#### FIX-011

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 根因：异步提交测试未验证 FlushAsync 顺序和失败边界。
- 修复：记录型文件系统新增操作序列、异步 flush 失败开关及 sync/async flush 计数；新增 write -> flush-async -> move 顺序和 flush 失败时目标不变、临时文件清理测试。
- 验证：net6/net8 ReviewFixRegression 各 34/34 PASS；完整 Unit 双 TFM 均包含该路径。

#### FIX-012

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
- 根因：Cold/CacheMiss 工厂只在 GlobalSetup 创建，跨 invocation 复用缓存。
- 修复：为 Cold、CacheMiss、CacheHit 增加目标级 `IterationSetup`；Cold/Miss 每轮重建 factory，CacheHit 每轮独立预热固定 key，setup 开销不计入 benchmark 方法。
- 验证：DynamicPlan Cold/CacheHit/CacheMiss，PlanBuildCount 100/500，ShortRun 各 3 次 warmup + 3 次 measurement 全部 PASS；原始结果 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.DynamicPlanBenchmarks-report.csv`。

#### FIX-013

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityExporter.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：StreamWriter 同步 Dispose 会通过 Csv 包装器触发调用方 stream 的同步 Flush。
- 修复：Async CSV 路径启用禁止同步 Flush 的包装器模式；最终数据仍由显式 `FlushAsync` 写出；测试流同步 Flush 抛错时断言导出成功且同步计数为零。
- 验证：net6/net8 AsyncPipeline 各 15/15 PASS；同步 Flush 抛错测试通过。

#### FIX-014

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：`SKIPPED`
- 原因：默认 `fixScope=recommended` 不处理 OPTIONAL。

### Round 2 汇总

- MUST_FIX：FIX-010 已完成；FIX-009 清理合同与 100K smoke 已完成，正式 staging 选型/资源证据与维护者批准仍 BLOCKED。
- SHOULD_FIX：FIX-011、FIX-012、FIX-013 已完成。
- PARTIAL：FIX-009 正式策略对照与资源预算。
- BLOCKED：API baseline 审批、正式 Memory/temp/hybrid 资源矩阵、并发/P95/P99 预算和跨平台 runner。
- FAILED：无。
- 回归验证：Release solution build 0 error（NuGet vulnerability source warnings）；Unit net6/net8 各 514 pass + 1 API approval blocked；Integration 双 TFM 各 15/15；Docs 10/10；Package Consumer 双 TFM PASS；DynamicPlan、Excel/Failure smoke PASS；`git diff --check` PASS。
- 下一步：维护者解除 API/资源/跨平台门禁后补齐正式 staging 对照，再进行独立 Review。

## Review 修复记录

### Round 3

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BO-RC-20260907-001/review.md`
- 执行器：Codex

#### FIX-015

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelExporter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiFailureWorkbookWriter.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：Async staging 只有硬编码 temp-file 实现，Failure Workbook override 还会额外创建内部临时文件；策略无法职责级比较。
- 修复：新增可注入的 Memory、TempFile、Hybrid staging 工厂；默认生产策略保持 TempFile；Failure Workbook 有异步 staging override 时直接序列化到唯一 staging，避免第二个内部临时文件；补充三种策略的导出和 Failure Workbook 清理测试。
- 验证：
  - net6/net8 AsyncPipeline 各 23/23 PASS
  - staging 文件目录无残留；无遗留 `bing-offices-*-async-*.tmp`
  - 正式 100K/并发/PeakWorkingSet/LOH/P95/P99 资源矩阵和维护者预算批准：BLOCKED（当前 runner/审批门禁未提供）

#### FIX-016

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：文件级测试替换了真实 `IFileExportCommitter`，没有覆盖 Atomic temp。
- 修复：改用真实 `DefaultFileExportCommitter`，仅注入阻塞 staging 工厂触发 mid-cancel；断言已有目标不变、Atomic temp 消失、Excel staging 无残留；pre-cancel 同样断言真实文件提交边界。
- 验证：net6/net8 AsyncPipeline 各 23/23 PASS。

#### FIX-017

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelExporter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：Exporter 清理诊断挂在翻译前异常；Importer Dispose 位于统一清理合同之外。
- 修复：Exporter 将翻译后的顶层异常设为 primary；Importer 统一 Dispose staging 并将清理异常写入顶层异常；staging Dispose 即使关闭失败也继续尝试删除文件；新增导出/导入 cleanup failure 的 `Exception.Data` 断言。
- 验证：net6/net8 AsyncPipeline 各 23/23 PASS，清理诊断 key 与 `IOException` 类型均通过。

#### FIX-018

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：`SKIPPED`
- 原因：默认 `fixScope=recommended` 不处理 OPTIONAL。

### Round 3 汇总

- MUST_FIX：FIX-016 已完成；FIX-015 已完成策略实现与职责测试，正式资源矩阵和维护者预算仍 BLOCKED。
- SHOULD_FIX：FIX-017 已完成。
- PARTIAL：FIX-015 正式资源/审批证据。
- BLOCKED：API baseline 审批、正式 staging 资源预算、并发/P95/P99 矩阵和跨平台 runner。
- FAILED：无。
- 回归验证：Release solution build 0 error（14 个 NU1900 vulnerability source warnings）；Unit net6/net8 各 522 pass + 1 API approval blocked；Integration 双 TFM 各 15/15；Docs 10/10；AsyncPipeline 双 TFM各 23/23；`git diff --check` PASS。
- 下一步：维护者提供固定 runner 与资源/API 批准后，补齐正式矩阵并再次独立 Review。

## Review 修复记录

### Round 4

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BO-RC-20260907-001/review.md`
- 执行器：Codex

#### FIX-019

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：Hybrid staging 只在完整 NPOI 同步序列化后才检查阈值并迁移。
- 修复：新增阈值切换写入流；Hybrid 的同步 `Write`、`Write(ReadOnlySpan<byte>)` 和 `WriteByte` 超过阈值时立即迁移到临时文件，保持后续写入位置连续；防御性异步迁移使用 `CopyToAsync`；新增超过阈值且 Flush 前已迁移的直接测试。
- 验证：net6/net8 AsyncPipeline 各 24/24 PASS；测试确认 Export 返回前已创建临时文件，内容完整且无残留。

#### FIX-020

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：缺少任务隔离的三策略资源矩阵和维护者批准证据。
- 修复：三策略实现和导出/Failure Workbook 清理专项测试已完成，测试使用独立临时目录；生产默认仍为 TempFile。
- 验证：
  - 三策略功能/清理：net6/net8 AsyncPipeline 通过
  - 正式 100K Excel、Failure 双 DOM、模板/图片/样式、并发 1/4/16/64、磁盘峰值、LOH、PeakWorkingSet、吞吐、P95/P99 以及 approvedBy/approvedAt：BLOCKED（固定 runner/维护者预算审批未提供）

#### FIX-021

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：`SKIPPED`
- 原因：默认 `fixScope=recommended` 不处理 OPTIONAL。

### Round 4 汇总

- MUST_FIX：FIX-019 已完成；FIX-020 的策略实现和功能证据已完成，正式资源矩阵/审批仍 BLOCKED。
- SHOULD_FIX：无。
- PARTIAL：FIX-020 正式资源和维护者批准。
- BLOCKED：API baseline 审批、正式 staging 资源预算、并发/P95/P99 矩阵和跨平台 runner。
- FAILED：无。
- 回归验证：Release solution build 0 error（14 个 NU1900 vulnerability source warnings）；Unit net6/net8 各 523 pass + 1 API approval blocked；AsyncPipeline 双 TFM 各 24/24；Integration 双 TFM 各 15/15；Docs 10/10；临时目录无残留；`git diff --check` PASS。
- 下一步：维护者提供 API/资源批准和固定 runner 后补齐正式矩阵，再进行独立 Review。

## Review 修复记录

### Round 5

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BO-RC-20260907-001/review.md`
- 执行器：Codex

#### FIX-022

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `tests/Bing.Offices.Tests/AsyncPipelineTest.cs`
- 根因：Hybrid staging 迁移时无条件把文件位置设置为长度；`SetLength` 超过阈值也没有立即迁移，seek/backpatch 会改变流语义。
- 修复：迁移前保存 MemoryStream.Position，复制到临时文件后恢复相同位置；Hybrid `SetLength` 超过阈值时立即迁移；补充回写内容、Position、Length 和临时文件断言。
- 验证：
  - `HybridStaging`：net6/net8 各 3/3 PASS
  - `AsyncPipelineTest`：net6/net8 各 26/26 PASS
  - `git diff --check`：PASS

#### FIX-023

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：
  - `tests/Bing.Offices.ResourceProbe/Program.cs`
  - `tests/Bing.Offices.ResourceProbe/StagingResourceMatrix.cs`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/benchmark/staging-matrix/formal-100k-v2.jsonl`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/benchmark/staging-matrix/formal-100k-v2.md`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/reports/resource-report.md`
- 根因：原有 ResourceProbe 只覆盖 mapping/unique，没有 Excel/Failure staging 的三策略、场景、并发和尾延迟原始产物；维护者预算审批也没有证据。
- 修复：新增任务隔离的父子进程 staging matrix runner，覆盖 Memory/TempFile/Hybrid、100K Excel、Failure Workbook 双 DOM、模板/图片/样式、并发 1/4/16/64；采集 PeakWorkingSet、LOH、GC/Allocated、临时磁盘峰值、吞吐、P95/P99 和清理残留；加入 2 GiB 子进程保护，超限单元结构化记录而不耗尽 runner；报告默认 TempFile、Hybrid 迁移和无自动磁盘回退语义。
- 验证：
  - 20 行 smoke：36/36 单元 PASS，原始产物 `artifacts/benchmark/staging-matrix/smoke.jsonl`
  - 100K formal matrix：36/36 单元均有记录；15 个完整指标单元通过，21 个因 2 GiB WorkingSet guard 记录为 guarded failure；完整指标单元清理残留均为 0
  - `ResourceProbe` Release build：PASS，4 个 `NU1900` 漏洞源警告
  - 维护者 `approvedBy`/`approvedAt`：仍为空，资源预算与默认策略批准为 `BLOCKED`；执行器不得自批

#### FIX-024

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：`SKIPPED`
- 原因：默认 `fixScope=recommended` 不处理 OPTIONAL。

### Round 5 汇总

- MUST_FIX：FIX-023 已补齐可重放矩阵、原始产物和报告，但正式资源预算/默认策略维护者批准仍 BLOCKED。
- SHOULD_FIX：FIX-022 已完成。
- PARTIAL：FIX-023 的批准门禁和超 2 GiB 高并发单元；不能将资源矩阵标记为发布预算通过。
- BLOCKED：API baseline `approvedBy/approvedAt`、正式资源预算批准、跨平台 runner；本轮未伪造外部证据。
- FAILED：无代码测试失败；formal matrix 的 guarded cells 按资源超限记录。
- 回归验证：Release solution build 0 error（14 个 `NU1900`）；Unit net6/net8 各 525 passed + 1 API approval blocked / 526 total；Integration net6/net8 各 15/15；Docs 10/10；AsyncPipeline net6/net8 各 26/26；Hybrid 专项各 3/3；`git diff --check` PASS。
- 下一步：维护者补充 API/资源预算审批并重新进行独立 Review。

## Review 修复记录

### Round 6

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BO-RC-20260907-001/review.md`
- 执行器：Codex

#### FIX-025

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`PARTIAL/BLOCKED`
- 修改文件：
  - `tests/Bing.Offices.ResourceProbe/Program.cs`
  - `tests/Bing.Offices.ResourceProbe/StagingResourceMatrix.cs`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/benchmark/staging-matrix/smoke-v6.jsonl`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/benchmark/staging-matrix/smoke-v6.md`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/benchmark/staging-matrix/formal-100k-v6.jsonl`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/benchmark/staging-matrix/formal-100k-v6.md`
  - `ai_docs/tasks/BO-RC-20260907-001/artifacts/reports/resource-report.md`
- 根因：父进程守卫终止子进程后只记录 `result=null`；100K workload 的输出低于生产 Hybrid 阈值，未验证真实迁移；矩阵没有 warmup/重复测量声明。
- 修复：父进程为每个子进程提供隔离 staging 目录并持续采样工作集；守卫触发后记录实际峰值、预算阈值、临时磁盘残留、文件数和清理结果，输出结构化 `budget-failed` 结果。子进程增加 1 次 warmup 和 2 次测量；Excel/template workload 加入确定性的高基数 512 字符描述，确保 100K 输出跨越 8 MiB Hybrid 阈值。新增 v6 原始 JSONL/报告并标记 v2 已被替代。
- 验证：
  - 1K smoke：36/36 单元通过；36/36 有结果、两次测量、无空结果。
  - 100K formal v6：36/36 有记录；Memory/TempFile/Hybrid 各 3 个完整指标、9 个结构化预算失败、0 个空结果；所有守卫记录均有实际峰值和父进程清理结果。
  - Hybrid 100K `excel-100k` 并发 1：临时磁盘峰值 `41,721,723 B`，清理后 `0 B`；template 并发 1：`41,704,272 B`，清理后 `0 B`。
  - ResourceProbe Release build：PASS，4 个 `NU1900` 漏洞源警告。

#### FIX-026

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：`BLOCKED`
- 修改文件：无（不能由执行器伪造维护者审批）
- 根因：正式 API baseline 的 `approvedBy`/`approvedAt` 为空，且维护者尚未批准 net6/net8 成员级差异。
- 处理：保留候选快照、成员级差异和审批门禁；未修改正式 baseline 元数据，也未弱化契约测试。需要有权维护者审查并提供真实审批字段后才能完成。
- 验证：
  - API contract net6：525 passed，1 个 `BLOCKED: API baseline approvedBy is empty`。
  - API contract net8：525 passed，1 个 `BLOCKED: API baseline approvedBy is empty`。

#### FIX-024

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：`SKIPPED`
- 原因：`fixScope=recommended` 默认跳过 OPTIONAL。

### Round 6 汇总

- MUST_FIX：FIX-025 的采集实现和矩阵证据已完成，但资源预算/默认策略审批 BLOCKED；FIX-026 因真实 API 维护者审批缺失而 BLOCKED。
- SHOULD_FIX：无。
- PARTIAL：FIX-025 的外部资源预算/默认策略审批。
- BLOCKED：API baseline `approvedBy/approvedAt`、正式资源预算批准、跨平台 runner。
- FAILED：无代码测试失败；完整 Unit 的唯一失败为预期 API 审批门禁。
- 回归验证：Release solution build 0 error（14 个 `NU1900`）；AsyncPipeline net6/net8 各 26/26；Integration net6/net8 各 15/15；Docs net8 10/10；完整 Unit net6/net8 各 525 passed + 1 API approval blocked / 526 total；`git diff --check` PASS。
- 下一步：维护者完成 API/资源预算审批后重新进行独立 Review。

## Git 状态

- 工作区包含本任务生产代码、测试、文档和 artifacts；未执行 git add/commit/push。
- 未自动创建 PR、tag 或发布 NuGet。
- `git diff --check` 已通过。
