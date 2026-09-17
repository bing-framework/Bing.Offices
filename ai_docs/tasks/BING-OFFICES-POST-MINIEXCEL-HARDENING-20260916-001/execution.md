<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001
AI_EXECUTION_FINISHED_AT: 2026-09-16T15:14:01.3707889Z

# 实施执行报告

## 执行结论

核心代码整改、测试、双 TFM Release 构建、最终包和 Consumer 验证已完成；Review Fix Executor 本轮状态为 `COMPLETED`，任务发布判定仍为 `PARTIAL`。资源矩阵虽然 36/36 场景通过，但没有批准预算；500K/1M、生产机器和外部 CI 按约束标记为 `NOT_VERIFIED`。因此本任务不宣称已经达到可发布状态。

## 任务信息

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 分支：`feat/miniexcel-provider`
- 计划：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/plan.md`
- 基线：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/baseline.md`
- Review：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/review.md`

## 计划执行情况

| 计划项 | 状态 | 说明 |
| --- | --- | --- |
| H-000 基线、版本冻结、产物扫描 | PASS | 保存任务起点 hash、TFM、IVT 和项目目录扫描。 |
| H-001/H-100 Core 兼容与预检所有权 | DEVIATED_OK | 没有复现历史声称的 netstandard2.0 编译错误；仍将预检改为 Core 共享源并由两个 Provider 链接编译。 |
| H-101 生产 IVT 清理 | PASS | Core 到 NPOI/MiniExcel 的生产友元为 0，无新增 SPI。 |
| H-200/H-201 日期与 RowIndex | PASS | Core 日期规则、1900/1904、DateTime/DateTimeOffset offset 策略和固定/动态 converter 物理行列号均有直接测试。 |
| H-300 异步/取消/资源/流/文件 | PARTIAL | 直接测试和 36 场景资源矩阵通过；资源预算审批仍 BLOCKED。 |
| H-400/H-401 热路径与关系 | NOT_VERIFIED | 已补齐 NPOI/MiniExcel 同 workload 对照；没有可重放的修改前样本，因此不得宣称优化前后收益。关系 binder 保留 `O(P×C)` 语义。 |
| H-500/H-501 专项和跨 Provider 合同 | PASS | MiniExcel 25/25 双 TFM；cross-provider 5/5 双 TFM，覆盖 ValidationMode、ValueMap、Unique、Converter、Relation、日期和结构化 Error。 |
| H-600/H-601/H-602 产物、Consumer、API 门禁 | PARTIAL | 最终包独立 cache Consumer 与 API identity/compare 通过；外部 CI 和发布审批仍未完成。 |
| H-700/H-701 中文注释与文档 | PASS | 本任务修改范围内的 XML 注释和能力边界文档已同步。 |
| H-800 最终收口 | PARTIAL | 代码和本地门禁通过，但外部证据/批准缺口阻止 `COMPLETED`。 |

## 已完成事项

- 保持 `Bing.Offices.Core` 为 `netstandard2.0`，并通过共享 `ExcelXlsxZipPreflight` 源文件让 NPOI/MiniExcel 使用同一 ZIP/XML 安全策略。
- 移除 Core 生产 `InternalsVisibleTo("Bing.Offices.Npoi")` 和 `InternalsVisibleTo("Bing.Offices.MiniExcel")`；测试友元保留。
- MiniExcel 日期导入改走 Core `DateTimeExcelValidationRule`/`ExcelDateParser`，传递 workbook `date1904`，修复跨目标 `DateTime`/`DateTimeOffset` 原始值归一化。
- 修复固定列和动态列 converter 的物理 1-based `RowIndex`，并缓存行循环中的反射绑定和动态键集合。
- 异步导入复制使用真实 `ReadAsync`/`WriteAsync`；取消、source/destination 所有权、临时文件、原子提交和旧目标保护均有直接证据。
- 增加 MiniExcel 专项、真实文件集成和 NPOI/MiniExcel 结构化合同测试；unsupported 能力保持 fail-fast 且不产生输出。
- BenchmarkDotNet 输出固定到根 `artifacts/benchmarks`；测试、包、Consumer、资源和 API 证据集中到根 `artifacts/`。
- Consumer 改为必须注入从本次包 nuspec 发现的版本；四个包同版本且 MiniExcel 无 NPOI 依赖。
- README 和 Excel 文档明确记录 XLSX-only、日期系统、异步/流语义、性能和 `NOT_VERIFIED` 边界。

## 部分/未完成事项

- 资源矩阵 `approvalStatus=BLOCKED`，缺少维护者批准人和批准时间；本机数据不能作为已批准的 2 CPU/4 GiB 发布预算。
- 本轮最终包已使用独立 cache `artifacts/consumers/review-round2-final/cache/net6|net8` 完成 restore/build/run；此前空 cache 的 TLS 失败仅保留为历史环境记录。
- 500K、1M、生产 2C/4GiB 机器和外部 CI 未执行，全部标记 `NOT_VERIFIED`。
- 资源探针产物沿用现有 ResourceProbe 的内部 task 标识 `BO-RC-20260908-002`，报告已明确其为复用证据而非当前任务审批。

## 修改文件

主要业务和测试文件：

- `src/Bing.Offices.Core/Bing.Offices.Core.csproj`
- `src/Bing.Offices.Core/AssemblyInfo.cs`
- `src/Bing.Offices.Core/Bing/Offices/Dates/ExcelDateParser.cs`
- `src/Bing.Offices.Core/Bing/Offices/IO/ExcelXlsxZipPreflight.cs`
- `src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj`
- `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiXlsxZipPreflight.cs`
- `src/Bing.Offices.MiniExcel/**`
- `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
- `tests/Bing.Offices.Tests/MiniExcelProviderContractTest.cs`
- `tests/Bing.Offices.Tests.Integration/MiniExcelProviderIntegrationTest.cs`
- `.github/workflows/ci.yml`
- `benchmarks/Bing.Offices.Benchmarks/Program.cs`
- `benchmarks/Bing.Offices.Benchmarks/MiniExcelProbe.cs`
- `benchmarks/Bing.Offices.Benchmarks/MiniExcelRealIoBenchmarks.cs`

任务证据文件：`baseline.md`、`api-diff.md`、`unit-test-report.md`、`integration-test-report.md`、`package-consumer-report.md`、`benchmark-report.md`、`resource-report.md`、`symbol-test-map.md`、`review.md`、`final-report.md`。

## API/数据/配置变化

- 公共成员签名无新增、删除或修改；双 TFM API snapshot compare 通过。
- 没有新增 Provider SPI，也没有修改版本属性或版本号。
- Consumer 保留 `PackageReference`，版本由 CI/local package feed 的唯一 nuspec 结果注入。
- `artifacts/` 为所有本任务验证输出根目录；`src`、`tests`、`benchmarks` 项目目录残留扫描计数为 0。

## 测试结果

- 全量 Unit：net6 `761/761`，net8 `761/761`。
- MiniExcel 专项：net6/net8 各 `25/25`（含真实 1900/1904 日期矩阵）。
- NPOI Preflight：net6/net8 各 `27/27`。
- Cross-provider：net6/net8 各 `5/5`。
- Integration：net6/net8 各 `39/39`。
- Docs：net8 `10/10`。
- 资源矩阵：36 个 child 场景均 `passed`，审批 BLOCKED。

各 TRX 路径和职责级映射见 `unit-test-report.md`、`integration-test-report.md`、`symbol-test-map.md`。

## Build/Typecheck/Lint/Format

- `dotnet build Bing.Offices.sln -c Release --no-restore --nologo -p:DocumentationFile= -p:GenerateDocumentationFile=false`：0 warning/0 error。
- Core `netstandard2.0`、NPOI/MiniExcel 双 TFM、Benchmark 和 ResourceProbe Release 独立构建均通过。
- Benchmark list/probe、四包 pack、API identity self-test/compare、最终包 Consumer 独立 cache restore/build/run 均通过；外部 CI 仍未执行。
- `git diff --check` 无 whitespace error；Git 只报告工作树 LF→CRLF 提示。

## 计划偏差

1. H-001 原计划要求复现具体 netstandard2.0 错误；实际 Core Release 构建没有错误，因此记录为“错误已不存在/未能复现”，没有伪造错误文本。
2. 预检实现采用 Core 源文件链接到两个 Provider，而不是引入 SPI；这是保持同一安全算法且清理生产 IVT 的最小方案。
3. 解决方案配置将 ResourceProbe 映射到 Debug；为满足 Release 门禁另行执行 ResourceProbe Release 构建和矩阵，并保留该配置偏差。

## 基线问题

工作树在任务开始时已包含前一 MiniExcel 任务变更，包括 `version.dev.props`。本任务开始和结束的版本文件 SHA-256 均保持不变：

- `version.props`：`77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1`
- `version.dev.props`：`BABFE7703A15E1C11F46B45D34BC59D7913ECB1E3DCFDA901E72F4D60A8D134D`

前一任务的项目包产物已移到 `artifacts/legacy-project-packages/`，未还原或覆盖用户已有工作树改动。

## 已知问题

- MiniExcel v1 的 `QueryAsync` 返回后仍由当前适配层同步枚举行对象；异步边界使用真实第三方 async API 和输入复制，但不能宣称端到端逐行异步。
- MiniExcel 不支持的 XLS、样式、模板、失败工作簿等能力继续结构化拒绝；未扩大 unsupported 能力。
- 关系绑定仍为 `O(P×C)`，当前没有足够稳定证据支撑更换 comparer/重复键语义。

## 风险与回归关注点

- 需要维护者批准资源预算并在目标机器复跑矩阵后，才能做正式发布判定。
- 最终包的干净独立 cache Consumer 已在本轮通过；仍需在外部 CI 复跑 PackageReference Consumer 与 workflow。
- 需要单独补跑或记录 500K/1M 容量证据，不能从 100K probe 外推。

## Reviewer 注意事项

独立审查结果写入 `review.md`，Round 2 状态为 `NEEDS_FIX`。本轮 REVIEW_FIX 已按 `recommended` 完成 FIX-004、FIX-007、FIX-008、FIX-009；`review.md` 保留为 Reviewer 独立证据，未被修改。下一步必须重新执行独立 Review，不能由本执行报告替代。

## Git 状态

- 当前分支：`feat/miniexcel-provider`。
- 未自动执行 `git add`、`git commit`、`git push`、tag、PR 或 NuGet publish。
- 工作树仍包含本任务及前一任务的未提交变更，需由维护者按项目流程审阅和提交。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/review.md`

#### FIX-001

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
  - `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
  - `docs/excel/09-providers.md`
- 根因：MiniExcel 导入路径没有读取 `ExcelImportValidationMode`，导致 Disabled 仍执行配置规则，WorkbookRules 也被静默忽略。
- 修复：按 `ConfiguredRules`/`ConfiguredAndWorkbook` 门控配置校验和 Unique；`Disabled` 跳过配置规则；MiniExcel 无法提供 Workbook 原生校验时，在预检阶段抛出结构化 `BingOfficesUnsupportedFeatureException`。
- 验证：`ValidationMode_ShouldDisableConfiguredRulesAndRejectWorkbookRules`（net6/net8）及全量 Unit 759/759：PASS。

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/Dates/ExcelDateParser.cs`
  - `tests/Bing.Offices.Tests/ExcelDateParserTest.cs`
- 根因：raw `DateTime` 和 Excel serial 转 `DateTimeOffset` 时绕过了 `ExcelDateOffsetPolicy`，默认隐式生成零 offset。
- 修复：所有无显式 offset 的 raw/serial 路径统一要求合法 `UseFixedOffset`；默认 `RequireExplicitOffset` 返回失败，并补充非法偏移和 nullable 覆盖。
- 验证：`ExcelDateParserTest`、MiniExcel 日期职责测试和全量 Unit 759/759（net6/net8）：PASS。

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
  - `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
  - `tests/Bing.Offices.Tests/MiniExcelProviderContractTest.cs`
- 根因：动态列 converter、validation 和转换错误使用固定的 1/0 列号，而不是表头中的物理位置。
- 修复：从已解析的物理表头建立 1-based 列索引，动态转换、校验和错误统一传递真实位置。
- 验证：`DynamicColumns_ShouldUsePhysicalColumnIndexForConverterValidationAndErrors` 及 cross-provider 5/5（net6/net8）：PASS。

#### FIX-004

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
  - `tests/Bing.Offices.Tests/MiniExcelProviderContractTest.cs`
  - `ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/symbol-test-map.md`
- 根因：1904 场景不是经真实 MiniExcel importer 的 workbook fixture，且职责级与跨 Provider 合同矩阵不完整。
- 修复：使用 XDocument 命名空间安全地改写真实导出 workbook 的 `date1904`，补齐 ValueMap、Validation、Unique、Converter、Relation、Error 的直接/跨 Provider 断言。
- 验证：MiniExcel 23/23、cross-provider 5/5、Integration 39/39、双 TFM：PASS。

#### FIX-005

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
- 根因：文件 API 只有调用前取消证据，没有写入暂存文件中途取消时的旧目标保护和临时文件清理证据。
- 修复：增加真实 `ExportToFileAsync` 中途取消测试，converter 在首行后取消 token，并断言旧目标字节、无 `.tmp` 残留和 converter 已执行；资源预算审批仍按现状记录为 BLOCKED。
- 验证：`AsyncFileExport_MidFlightCancellation_ShouldPreserveExistingTargetAndCleanTempFile`（net6/net8）及 Integration 39/39：PASS。

#### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `benchmarks/Bing.Offices.Benchmarks/Program.cs`
  - `benchmarks/Bing.Offices.Benchmarks/ProviderComparisonBenchmarks.cs`
  - `benchmarks/Bing.Offices.Benchmarks/ProviderComparisonProbe.cs`
  - `ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/benchmark-report.md`
- 根因：原 benchmark 只测 MiniExcel，不能和 NPOI 在相同 workload 下比较，也没有可重放的修改前样本。
- 修复：新增 NPOI/MiniExcel 相同 100K rows、同 request 的 Sync/Async Export+Import 基准和 3 次重复受控 probe；报告明确 `after` 对照，不合成 `before` 收益。
- 验证：BenchmarkDotNet list 包含四个 ProviderComparison 方法；`artifacts/review-fix/provider-comparison-100k-final.json` 含 12 个样本：PASS。历史 before、500K/1M、生产机仍 `NOT_VERIFIED`。

### Round 1 汇总

- MUST_FIX：FIX-001、FIX-002；已完成 2/2。
- SHOULD_FIX：FIX-003 至 FIX-006；已完成 4/4。
- PARTIAL：无 FIX 项；发布级资源审批、隔离 cache、500K/1M、生产机器和外部 CI 仍按计划为 `PARTIAL`/`NOT_VERIFIED`。
- BLOCKED：资源矩阵批准人/时间缺失，approvalStatus=`BLOCKED`。
- FAILED：无。
- 回归验证：双 TFM Unit `759/759`、MiniExcel `23/23`、cross-provider `5/5`、Integration `39/39`、API snapshot compare `net6/net8 PASS`、Solution Release build `0 warning/0 error`。
- 下一步：重新进行独立 Review；在资源批准和外部环境证据补齐前，不得宣称发布完成。

### Round 2

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/review.md`

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/Dates/ExcelDateParser.cs`
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelValueAdapter.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs`
  - `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
- 根因：MiniExcel 对日期样式数值返回 `DateTime` 后，原始 serial 和 `date1904` 语义未进入 Core，真实 XLSX 日期会静默偏移。
- 修复：MiniExcel 在 OLE Automation 可表示范围内把第三方 `DateTime`/`DateTimeOffset` 归一化为 serial 并传递 `date1904`；Core 使用 `FromOADate` 保留 1900 伪闰日合同并消除浮点 ticks 漂移；NPOI 日期样式数值保留原始 numeric serial，两个 Provider 统一走 Core 规则。
- 验证：真实 1900/1904 serial `0/1/59/60/61` fixture、跨 Provider 日期合同和 DateTimeOffset 精度测试在 net6/net8 均 `3/3 PASS`；`artifacts/review-round2-probe/final-date-matrix.log` 中两个 Provider 的 12 个边界结果一致；双 TFM 全量 Unit `761/761 PASS`。

#### FIX-007

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `build/ApiSnapshot/Program.cs`
  - `build/ApiSnapshot/CandidateIdentityContractTest.cs`
- 根因：API identity 工具对 `lib/*.dll` 只复用 Release 路径 identity，没有读取包内条目。
- 修复：逐个提取包内 DLL 到唯一临时目录，使用 MetadataLoadContext 重新计算程序集 identity 并与 Release 资产比较；补充损坏、缺失、有效错程序集和错 TFM 契约场景。
- 验证：identity self-test、LF/CRLF/source/assembly/nupkg/approval tamper 合同均按预期通过或失败；本轮最终 MiniExcel DLL 替换为文本后 API compare `exit=1`，证据为 `artifacts/api-round2-final3/compare-tampered.log`。

#### FIX-008

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/package-consumer-report.md`
  - `ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/final-report.md`
- 根因：旧 Consumer 日志早于最终 pack，不能证明消费了本轮包。
- 修复：重新 pack 四个 `2.0.0` 包，在 `artifacts/consumers/review-round2-final/` 使用独立 net6/net8 cache restore/build/run，并记录最终包哈希、日志时间和输出 DLL 与包条目哈希。
- 验证：net6/net8 restore/build/run 均 PASS；最终包写入约 `08:39:50Z`-`08:39:56Z`，Consumer build/run 日志约 `08:45:59Z`-`08:46:01Z`，输出均含 `package-consumer-ok`。

#### FIX-009

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
  - `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`
- 根因：固定映射按 mapping 顺序使用 `index + 1`，表头重排时 converter、validation 和结构化错误的 `ColumnIndex` 错误。
- 修复：固定列与动态列统一通过物理表头查找 1-based 列号，并在转换、校验和错误上下文复用。
- 验证：`FixedColumns_ShouldUsePhysicalColumnIndexWhenMappingOrderDiffers` 在 net6/net8 均 PASS；双 TFM 全量 Unit `761/761 PASS`。

#### FIX-010

- 严重程度：LOW
- 处理要求：OPTIONAL
- 执行状态：DEFERRED
- 原因：本轮 `fixScope=recommended` 默认不纳入 OPTIONAL；不改变 benchmark 数据或发布结论，保留 Reviewer 记录供后续处理。

### Round 2 汇总

- MUST_FIX：FIX-004、FIX-007；已完成 `2/2`。
- SHOULD_FIX：FIX-008、FIX-009；已完成 `2/2`。
- OPTIONAL：FIX-010；按默认 scope `DEFERRED`。
- PARTIAL：资源预算批准、500K/1M、生产 2 CPU/4 GiB 和外部 CI 仍未完成，保持 `PARTIAL`/`NOT_VERIFIED`。
- BLOCKED：资源矩阵 approvalStatus=`BLOCKED`，缺少批准人和批准时间。
- FAILED：无。
- 回归验证：双 TFM 全量 Unit `761/761`、Integration `39/39`、Docs `10/10`、API compare `net6/net8 PASS`、tampered package `exit=1`、最终包 Consumer net6/net8 restore/build/run PASS、Solution Release build `0 warning/0 error`、`git diff --check` 无 whitespace error。
- 下一步：重新进行独立 Review；本执行终态 `COMPLETED` 仅表示本轮 recommended FIX 已完成，不代表 Reviewer 已通过或可直接发布。

### Round 3

- 触发原因：Round 2 收口后补强 FIX-004 的真实职责证据；`review.md` 保持 Reviewer 原始 `NEEDS_FIX` 内容不变。
- 修改范围：仅增强 `tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`。真实 XLSX serial `0/1/59/60/61` fixture 现在同时写入 `DateTime` 与固定 `+08:00` `DateTimeOffset` 两列；两列使用同一日期样式、同一 `ExcelImport` request，并分别由 MiniExcel/NPOI 导入后逐行完整比较。
- 固定 offset 通过测试请求级命名 converter 显式构造 `ExcelDateAttribute` 并调用共享 Core 日期解析器，避免为 nullable 属性引入无关公共 API 变更。
- 验证：`RealXlsxDateSystems_ShouldUseCoreSerialContract` 在 `artifacts/tests/review-fix-round2/net6-real-date-offset/` 与 `artifacts/tests/review-fix-round2/net8-real-date-offset/` 均 `1/1 PASS`；此前 Round 2 的双 TFM 全量 Unit `761/761`、Integration `39/39`、API、Consumer、Solution build 和 `git diff --check` 证据仍有效。
- 结论：Round 3 未引入新的生产行为或版本变更；FIX-004 的真实日期、1900/1904、DateTime/DateTimeOffset、固定 offset 和跨 Provider 同文件合同证据已补齐。继续保持下一步独立 Review，发布级资源批准、500K/1M、生产机和外部 CI 仍为 `PARTIAL`/`NOT_VERIFIED`。

### Round 4

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/review.md`（本轮未修改）

#### FIX-004

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：`COMPLETED`。
- 根因：MiniExcel 将日期样式单元格先转换为 `DateTime`，导致负小数 serial 的原始信息丢失；Core 原有基于 OA 的 1900/负数映射也不满足线性 Excel 合同。
- 修复：新增 `MiniExcelRawDateSerialReader` 从 worksheet XML 保留数值 serial 和物理行列；导入、转换及验证共享该 `ExcelCellValue`；Core 对负数和 1904 日期系统使用线性毫秒舍入，对 1900 正数保留伪闰日合同。
- 修改范围：`src/Bing.Offices.Core/Bing/Offices/Dates/ExcelDateParser.cs`、`src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelRawDateSerialReader.cs`、`src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelValueAdapter.cs`、`src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs` 及对应日期/Provider 测试和日期文档。
- 验证：`ExcelDateParserTest` 与 `MiniExcelProviderTest` 目标测试双 TFM `37/37 PASS`；全量 Unit 双 TFM `765/765 PASS`；Integration 双 TFM `39/39 PASS`；真实 XLSX 日期探针双 TFM 均通过，`artifacts/review-fix-round4/date-probe-net6.log` 和 `date-probe-net8.log` 中 NPOI/MiniExcel 的 1900/1904、负数、0/1/59/60/61 和 DateTimeOffset 结果一致。

#### FIX-009

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：`COMPLETED`。
- 修复：固定列和动态列统一以 `ReadColumnRange.StartIndex` 为基准，从实际物理表头生成绝对 1-based `ColumnIndex`，供 converter、validation 和结构化转换错误复用。
- 验证：`ReadColumnRange_ShouldReportAbsolutePhysicalColumnIndexAcrossProviders` 及既有固定/动态列职责测试双 TFM通过；真实探针中两 Provider 均报告 `ValueConversion,row=2,column=3,property=Age`。

#### FIX-011

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：`COMPLETED`。
- 修复：`.github/workflows/ci.yml` 将 `sha256sum` 输出转为大写后再与冻结常量比较，避免 Git Bash 大小写差异造成误报；版本文件和冻结常量未变。
- 验证：`artifacts/review-fix-round4/ci-hash/verify.sh` 使用真实 Git Bash 对原文件和隔离修改副本分别完成正/负断言并退出 `0`；外部 CI 仍为 `NOT_VERIFIED`。

#### FIX-012

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：`COMPLETED`。
- 根因：默认 MiniExcel `DateTimeOffset` 导出返回 DTO 原值，第三方序列化后丢失 offset；显式 converter/formatter 优先级必须保持不变。
- 修复：固定列默认 DTO 导出使用不变文化的 ISO `O` 文本，动态列默认同样使用 `O` 文本；显式 `Formatter`、`NumberFormat` 或 converter 仍优先。
- 验证：固定/动态 `DateTimeOffset` 导出测试和真实探针双 TFM 均通过，worksheet XML 含完整 offset 文本，MiniExcel/NPOI 导入后 DTO 精确相等；Docs `net8.0 10/10 PASS`。

### Round 4 汇总

- MUST_FIX：FIX-004、FIX-012，已完成 `2/2`。
- SHOULD_FIX：FIX-009、FIX-011，已完成 `2/2`。
- OPTIONAL：FIX-010 按 `recommended` scope `DEFERRED`，未修改 benchmark 数据或发布结论。
- 回归验证：Unit 双 TFM `765/765`、Integration 双 TFM `39/39`、Docs `net8.0 10/10`、带 XML 文档的 Solution Release build `0 warning/0 error`、真实日期/DTO/物理列探针双 TFM通过、`git diff --check` 无 whitespace error。
- 产物门禁：首次使用禁用 XML 文档的 `--no-build pack` 触发 `NU5026`，未将失败误记为包通过；随后以 `dotnet pack --no-restore -p:GenerateDocumentationFile=true` 重新生成四个 `2.0.0` nupkg 到 `artifacts/review-fix-round4/packages`。独立 Consumer net6/net8 restore/build/run 均 `exit 0`，8 行包/cache/Release/output DLL 原始 SHA-256 全部一致，证据见 `artifacts/review-fix-round4/final-product-verification.md` 和 `consumers/results/hash-evidence.csv`。
- API 门禁：当前 Release 与本轮包完成 capture；同步当前候选源码清单及三份变更包的 baseline identity 哈希后，API compare 双 TFM `exit 0`，最终重建后的 `artifacts/review-fix-round4/api-compare-post-build/api-diff.json` 的 `net6.0`、`net8.0` 均为空，公共成员快照未变化。
- 发布边界：资源矩阵审批仍 `BLOCKED`；500K/1M、生产 2 CPU/4 GiB 和外部 CI 仍 `NOT_VERIFIED`；未修改版本文件，未执行 commit、push、tag、PR 或 NuGet publish。
- 下一步：重新执行独立 Review；本轮执行收口不替代 Reviewer 结论。

## Review 修复记录

### Round 5

- Review 状态：`NEEDS_FIX`。
- Fix Scope：`recommended`。
- Review 文件：`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/review.md`（本轮未修改）。

#### FIX-013

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：`COMPLETED`。
- 根因：MiniExcel `Query`/`QueryAsync` 行枚举包含表头后首条正文，`physicalRow` 却从 `DataRowStartIndex + 1` 开始；跳过正文时继续递增会使原始日期 serial、`SourceRows`、converter 和 validation 上下文行号整体偏移。
- 修复：共享同步/异步 `ImportTypedSheetRows` 路径将物理行起点改为 `HeaderRowIndex + 2`，保留跳过逻辑；补充真实 `XSSFWorkbook` fixture，覆盖表头行 `0/2`、跳过 `1/2` 行、1900/1904 日期系统、正负小数 serial、固定 `DateTime`/显式 `DateTimeOffset` DTO 和动态日期列，并断言完整日期结果、错误集合、`SourceRows`、converter/validation `RowIndex` 与 `ColumnIndex`。
- 修改文件：`src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`、`tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`、`ai_docs/tasks/BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001/symbol-test-map.md`、`build/api-snapshot-baseline.json`。
- 验证：MiniExcel 目标测试双 TFM `34/34 PASS`；Unit 双 TFM `771/771 PASS`；Integration 双 TFM `39/39 PASS`；Docs `net8.0 10/10 PASS`；带 XML 文档的 Solution Release build `exit 0`（0 error，92 个既有 `CS1591` warning）；真实 probe 双 TFM `exit 0` 且 NPOI/MiniExcel 日期与物理行均为 `rows=2|3`；四包 pack、Consumer net6/net8 restore/build/run、8 行原始哈希证据均通过；MiniExcel nuspec 无 NPOI 依赖；API capture/compare `exit 0` 且 `api-diff.json` 双 TFM 为空；identity self-test 通过；版本文件哈希未变；嵌套产物扫描 `nestedCount=0`；`git diff --check` 无 whitespace error（仅既有 CRLF 提示）。

#### FIX-010

- 严重程度：LOW；处理要求：OPTIONAL；按 `recommended` scope `DEFERRED`，未修改 benchmark 数据或发布结论。

### Round 5 汇总

- MUST_FIX：`1/1` 完成；SHOULD_FIX：`0`；OPTIONAL：`FIX-010` 延后；FAILED：无。
- 回归验证：Unit、Integration、Docs、Solution build、真实 provider probe、四包 Consumer、哈希、API snapshot 和产物目录治理均完成并留存于 `artifacts/review-fix-round5/`。
- 发布边界：资源矩阵审批仍 `BLOCKED`；500K/1M、生产 2 CPU/4 GiB 和外部 CI 仍 `NOT_VERIFIED`；任务发布状态保持 `PARTIAL`。
- 下一步：执行下一轮独立 Review；`review.md` 保持 Reviewer 原始证据；未执行 commit、push、tag、PR 或 NuGet publish。
