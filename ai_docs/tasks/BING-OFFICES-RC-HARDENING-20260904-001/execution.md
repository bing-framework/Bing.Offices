<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-RC-HARDENING-20260904-001
AI_EXECUTION_FINISHED_AT: 2026-09-05T20:56:44.2183698+08:00

# 实施执行报告

## 执行结论

`BLOCKED`。本任务已完成当前工作区内可安全执行的核心实现、测试、包消费、资源探针、文档、独立 Review 和 Review 修复；正式 API approval、性能/资源预算、完整矩阵及发布交付条件仍未满足，因此不能进入发布。

## 任务信息

- Task ID：`BING-OFFICES-RC-HARDENING-20260904-001`
- 执行器：Copilot plan-executor / Review Fix Executor
- 计划：`plan.md`
- 执行状态文件：`execution.md`
- 原始 Review：`review.md`，保留 `AI_REVIEW_STATUS: NEEDS_FIX`
- 修复后复核：`review-round2.md`，结论 `PASS_WITH_ISSUES`
- 发布结论：`No-Go`
- 工作区：`e:\Bing_Framework\Bing.Offices`

## 计划执行情况

### 已完成或已验证

- `RC0-01`/`RC0-02`/`RC0-03`：环境、项目、API、异常、日期、NPOI 扩展、弃用和基线矩阵已形成证据。
- `RC1-01`：统一异常层次、稳定 code/operation/provider/stage、inner exception、observer 和边界翻译已接入主链。
- `RC1-02`：Excel/CSV 共用确定性日期 parser，覆盖 ISO、显式格式、offset、1900/1904 和 Workbook DATE/TIME。
- `RC1-03`：API Contract 重复分类问题已修复；生产程序集之间无 IVT；formal hash 未被修改。
- `RC1-04`/`RC1-05`：XLSX ZIP preflight、资源限制、XML 安全、Failure Workbook 和 Data Validation 回归已完成；完整资源矩阵仍不完整。
- `RC2-01`/`RC2-02`/`RC2-03`：NPOI public extensions、Mapping 死字段和命名迁移已完成并经过 consumer/Unit 验证。
- `RC2-04`：已批准的旧 API/全局 CSV 状态清理完成；DataTable 显式兼容类及部分旧异常/执行细节仍待治理。
- `RC4-01` 至 `RC4-04`：Unit、Integration、Docs、package-only consumer 已重新验证。
- `RC5-03`/`RC5-04`：Excel 12 场景、mapping/unique 16 场景和 MappingValidation 10 场景证据已保存；预算未批准。
- `RC6-01`/`RC6-02`/`RC6-03`/`RC6-04`：报告、candidate API diff、独立 Review 和最终 No-Go 门禁已形成；API approval 和发布门禁阻断保留。

## 已完成事项

- 新增 `BingOfficesException` 及 configuration/import/export/resource/file commit/unsupported 子类和稳定枚举。
- 增加异常观察器接口、单次通知 dispatcher、observer failure 诊断和并发保护。
- 统一 Excel/CSV 日期解析，默认 `yyyy-MM-dd`、InvariantCulture、`DateTimeKind.Unspecified`；DateTimeOffset 使用显式或固定 offset。
- 增加 `ExcelCellValue.IsDate1904`，统一 1900/1904 serial 和 Workbook DATE/TIME 解析。
- 在 NPOI `WorkbookFactory.Create` 前增加 XLSX ZIP preflight：entry、单项/总解压大小、压缩比、sharedStrings/styles/worksheet、路径、重复 entry、DTD/entity 和取消检查。
- 将六类 NPOI 扩展恢复为 Provider User API，内部 extension helper 继续隐藏。
- 完成 `AddBingOfficesNpoi`、`MaxSerializedBytes`、`RequireExpectedHeaders`、`MaxReadColumns`、`ReportEmptyRows`、`StopAtFirstEmptyRow` 等 API 收敛。
- 删除已批准的旧 converter、旧 validation attributes、CSV 全局可变 separator/quote 和 Mapping 死字段。
- Review Round 1 修复 `ExcelResourceLimits.Validate()` 对 `MaxZipCompressionRatio = null` 的错误 `.Value` 访问。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`review.md`

#### FIX-001

- 严重程度：HIGH
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`
	- `tests/Bing.Offices.Tests/NpoiXlsxZipPreflightTest.cs`
- 根因：可选的 `MaxZipCompressionRatio` 为 null 时，校验逻辑无条件访问 `Nullable.Value`。
- 修复：仅在 `HasValue` 时校验非正数、NaN 和 Infinity；null 保持“不限制”语义。
- 验证：
	- net8 ZIP 专项：`16/16`，PASS。
	- net6 ZIP 专项：`16/16`，PASS。
	- netcoreapp3.1 ZIP 专项：`16/16`，PASS。
	- `get_errors`：相关源码/测试目录 `No errors found`。

### Round 2 独立复核

- Review 文件：`review-round2.md`
- 结论：`PASS_WITH_ISSUES`
- P0/P1 代码修复项：`0 open`。
- 发布治理/证据问题：保留，不能据此改为 Go。

## 部分/未完成事项

1. Formal API candidate 与 formal baseline hash 不一致，尚无本任务维护者批准；未修改正式 baseline 或 hash 断言。
2. 性能和资源预算没有维护者批准；当前 Benchmark/ResourceProbe 只能作为证据，不能作为发布通过。
3. Failure Workbook 双 DOM、完整 NPOI DOM、100k/1M Excel/CSV、取消延迟和 XLS/OLE 内部结构矩阵未完成。
4. 历史 StreamPipeline 100k 记录的 3 个异常未形成当前任务可采信的解释；历史 Failure Workbook 高分配记录不能代替当前测量。
5. DataTable 显式兼容类、Office exception 家族、UniqueTracker 和剩余 public execution detail 尚未逐符号完成审批或迁移。
6. 同版本本地 nupkg 和隔离缓存 consumer 成功不等于正式 feed、clean clone 或发布不可变性证明；历史深路径存在 `MSB3106` 环境问题。
7. 导出 Workbook 元数据仍使用 `DateTime.Now`；不影响当前输入日期 parser，但未满足更强的导出元数据确定性目标。

## 修改文件

本任务涉及的主要生产/测试文件包括：

- `src/Bing.Offices.Abstractions/`：异常合同、`ExcelCellValue`、导入资源限制、请求 API。
- `src/Bing.Offices.Core/`：异常分发、日期 parser、日期属性、CSV/Mapping/IO 异常边界。
- `src/Bing.Offices.Npoi/`：导入/导出异常边界、日期 façade、ZIP preflight、public extensions。
- `tests/Bing.Offices.Tests/`：异常、日期、ZIP、API、Mapping、ResourceProbe 回归。
- `tests/Bing.Offices.Tests.Integration/`：真实 Excel/CSV/文件异常合同回归。
- `tests/Bing.Offices.ResourceProbe/`：独立 Excel 资源拒绝场景。
- `ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/`：计划、进度、决策、API、测试、包、Benchmark、资源、Review 和最终报告。

完整变更以最终 `git status --short`/`git diff --stat` 为准；未覆盖或回滚共享工作区中的前序任务改动。

## API/数据/配置变化

- 新增统一异常公共合同和资源限制配置。
- `ExcelCellValue` 增加 1904 date-system 标识。
- `ExcelDateAttribute` 增加 `OffsetPolicy`、`OffsetMinutes`。
- `AddNpoi`、`MaxBytes` 和四个含糊导入命名完成 Major rename。
- NPOI 六类 extension 成为 Provider User API。
- 未更新 formal API baseline；candidate 与 formal 的差异见 `api-diff.md`。
- 未执行数据迁移、数据库操作或外部基础设施变更。

## 测试结果

- Unit net8：`414 passed / 415 total / 1 failed`；唯一失败为 formal API hash。
- Unit net6：`414 passed / 415 total / 1 failed`；唯一失败为 formal API hash。
- Unit netcoreapp3.1：`414 passed / 415 total / 1 failed`；唯一失败为 formal API hash。
- ZIP/资源专项：三目标框架均 `16/16`。
- Integration net8：`15/15`。
- Integration net6：`15/15`。
- Docs net8：`11/11`。
- Package-only consumer：net6/net8 在隔离 `.packages-final` 缓存下 restore/build/run 成功，输出 `package-consumer-ok`。
- Excel ResourceProbe：12 个 child-process 场景；新增 5 个 Preflight reject 场景均 `importedRows=0`。
- Mapping/Unique ResourceProbe：`16/16`。
- MappingValidation Benchmark：`10/10` ShortRun。
- `get_errors`：相关源码/测试目录 `No errors found`。

TRX、JSONL、BDN log 和结果文件路径见各专项报告。

## Build/Typecheck/Lint/Format

- Release solution build：PASS，0 error，22 warnings。
- Review 修复后的相关源码/测试 `get_errors`：PASS，无诊断错误。
- 独立 lint 入口：未发现，未伪造 lint 通过。
- `git diff --check`：通过；输出仅包含 CRLF/LF 转换提示，没有 whitespace error。

## 计划偏差

- 计划中的 P3 大规模职责拆分、完整 CSV writer A/B、Failure Workbook 双 DOM 资源矩阵和完整 Excel/CSV 1k/10k/100k/1M 矩阵未在当前任务继续扩大执行，避免在 API/预算未收口时引入额外行为风险。
- 计划要求的正式 API approval、性能/资源预算和发布环境复现仍等待维护者决策或环境条件。
- package consumer 为解决同版本缓存污染使用任务专属隔离缓存；该 workaround 已如实记录，不作为正式 feed 证明。

## 基线问题

- formal baseline 位于上一任务 `BING-OFFICES-PRE-RC-CLEANUP-20260901-001` 的 `api-snapshot-formal-baseline-20260903.json`，本任务 candidate 位于当前任务 `artifacts/api-snapshot-candidate-20260904/`。
- candidate 与 formal hash 不一致，属于已记录的 API breaking/change set，尚无当前任务批准。
- 缺失的指定实现分析报告和 methodology router 未伪造，已在 `baseline.md`/`decisions.md` 记录。

## 已知问题

- `ExcelHelper` 导出元数据使用本地当前时间。
- XML 预检没有独立最大深度预算。
- XLS/OLE 不具备 XLSX ZIP 等价的 DOM 前内部预检。
- public execution detail 的长期边界仍需治理。
- netcoreapp3.1 构建存在依赖包不支持警告；本地 runtime 可运行，不能推广为依赖包长期支持。

## 风险与回归关注点

- Public API breaking diff 必须由维护者逐项批准后才能更新 formal baseline。
- 资源限制只在配置值被传入并通过请求构建器校验时生效；ZIP preflight 不等价于 NPOI DOM 内存硬上限。
- File API 的原子提交和目标流所有权具有 Windows-specific 行为，跨平台仍需独立验证。
- 同版本 nupkg 重新生成会影响本地 lockfile 内容哈希；发布前必须使用不可变 feed/clean clone 复核。
- `DateOnly` 当前不公开、不支持；不得在文档或 consumer 中暗示已支持。

## Reviewer 注意事项

- `review.md` 保留 Round 1 独立 `NEEDS_FIX` 原始证据；`review-round2.md` 记录修复后独立复核 `PASS_WITH_ISSUES`。
- Round 2 未发现新的 P0/P1 代码问题，但未批准的 API、预算和资源矩阵仍足以阻断发布。
- 不要把 Unit 的唯一 formal hash failure 归类为行为回归，也不要以修改 hash 代替 API approval。
- 不要把 mapping/unique ResourceProbe 或七个/十二个 Excel 样本解释成任意 Workbook 内存硬上限。
- 不要把本地 package-only consumer 成功解释成 feed、clean clone 或 publish 证明。

## Git 状态

- 未自动执行 `git add`。
- 未自动执行 `git commit`。
- 未自动执行 `git push`。
- 未创建 tag、PR 或执行 publish。
- 未执行 `git reset`、`git clean` 或覆盖未知用户修改。

## Review 修复记录

### Round 2

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/review.md`
- 执行状态：当前 fixScope 纳入的 MUST_FIX 已完成；不代表 Reviewer 已通过。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 根因：字段处理、属性 setter、关系委托、Failure Workbook 序列化和诊断回调存在嵌套异常捕获边界；反射 setter 和 NPOI 序列化还会将用户异常包在 `TargetInvocationException` 或 NPOI 运行时异常中。
- 修复：
	- CSV/NPOI converter、validator、setter、UniqueTracker、关系绑定、Failure Workbook 和 Excel 导出转换路径统一保留 `OperationCanceledException`、`OutOfMemoryException`、`StackOverflowException` 及既有领域异常。
	- CSV 与 NPOI 属性 setter 使用 `ExceptionDispatchInfo` 解包反射目标异常，避免取消/致命异常被二次包装。
	- Failure Workbook 序列化路径解包 NPOI 包装异常，传播内部取消/致命异常；诊断 sink 仅隔离普通异常。
- 修改文件：
	- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPropertyBinding.cs`
	- `src/Bing.Offices.Npoi/ExcelColumnPlan.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiImportRowMaterializer.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiRelationBinder.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
	- `tests/Bing.Offices.Tests/CsvTest.cs`
	- `tests/Bing.Offices.Tests/StreamPipelineTest.cs`
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 直接回归覆盖：CSV converter/validator/setter，Excel import/export converter、validator/setter，五类关系委托，Failure Workbook serializer 和 diagnostic sink；取消与 OOM 均使用直接抛出异常对象验证原样传播。
- 验证：
	- 三目标框架专项异常传播测试：PASS。
	- 完整 Unit：net8 `453 passed / 454 total / 1 failed`；net6 `453 passed / 454 total / 1 failed`；netcoreapp3.1 `453 passed / 454 total / 1 failed`。唯一失败均为既有 formal API snapshot hash 未获批准，不修改 baseline。
	- Integration：net8 `15/15` PASS；net6 `15/15` PASS。
	- `get_errors`：src/tests `No errors found`。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 根因：XLSX XML 预检只有与 ZIP entry 长度绑定的文档处理，没有独立可配置的字符预算和深度预算。
- 修复：
	- `ExcelResourceLimits` 增加 `MaxXmlCharacters`（默认 `64 MiB`）和 `MaxXmlDepth`（默认 `256`），null 表示关闭，非正值在配置边界拒绝。
	- `NpoiXlsxZipPreflight.ValidateXmlSafety()` 在 `WorkbookFactory.Create()` 前使用流式 `XmlReader`，独立统计节点名、节点值、属性名和值，并检查 `reader.Depth`；超限抛 `BingOfficesResourceLimitException`，阶段为 `Preflight`。
	- ResourceProbe 增加 `xml-depth-limit` 和 `xml-character-limit`，通过真实 importer 验证预检拒绝且 `importedRows=0`。
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiXlsxZipPreflight.cs`
	- `tests/Bing.Offices.Tests/NpoiXlsxZipPreflightTest.cs`
	- `tests/Bing.Offices.ResourceProbe/Program.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 验证：
	- `NpoiXlsxZipPreflightTest`：net8/net6/netcoreapp3.1 均 PASS；覆盖 null、无效值、边界、超限、DTD/entity、取消、损坏 ZIP/XML 和 workbook/sharedStrings/styles/worksheet 四类 XML 部件。
	- Excel ResourceProbe：14 个 child-process 场景 PASS；7 个 Preflight reject 场景（含 XML 深度和字符预算）均 `importedRows=0`。
	- `dotnet build .\Bing.Offices.sln -c Release --no-restore`：PASS，0 errors，28 warnings。
	- `git diff --check`：PASS；仅有 CRLF/LF 转换提示，无 whitespace error。

### Round 2 汇总

- MUST_FIX：`FIX-002`、`FIX-003`
- 已完成：`FIX-002`、`FIX-003`
- 跳过：`FIX-004`、`FIX-005`，因用户明确要求本轮只处理 MUST_FIX；未修改其生产代码或测试。
- FAILED：无纳入范围的失败项。
- 回归验证：新增直接异常传播测试、三 TFM Unit/ZIP 专项、三 TFM完整 Unit、net6/net8 Integration、ResourceProbe、Release solution build、`get_errors` 和 `git diff --check` 已完成。
- 已知外部阻断：formal API candidate hash 仍需维护者批准；不以修改 formal baseline 或测试 hash 绕过。
- 下一步：再次进行独立 Review；Reviewer 应重新检查 `review.md` 中 `FIX-002`、`FIX-003` 的证据，不应将本 execution 终态视为 Review PASS。

### Round 3

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/review.md`
- 执行状态：当前 fixScope 没有开放的 `MUST_FIX`；本轮无需修改代码。

#### FIX-004

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：当前用户明确要求仅根据 `MUST_FIX` 继续修复；该项为 `SHOULD_FIX`，不在本轮 `fixScope=must` 范围内。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：当前用户明确要求仅根据 `MUST_FIX` 继续修复；该项为 `SHOULD_FIX`，不在本轮 `fixScope=must` 范围内。

### Round 3 汇总

- MUST_FIX：无。
- 已完成：无代码修复项；上一轮 `FIX-002`、`FIX-003` 已保持已完成状态。
- 跳过：`FIX-004`、`FIX-005`，因本轮使用 `fixScope=must` 且用户要求仅处理 `MUST_FIX`。
- 修改文件：仅更新本 `execution.md`；未修改 `review.md`、业务代码或测试代码。
- 验证：无代码变更，不重复运行无关测试；任务状态脚本校验 Review 状态和 fix scope 通过。
- 下一步：如需处理剩余 `SHOULD_FIX`，应使用 `fixScope=recommended`，然后再次进行独立 Review。

### Round 4

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RC-HARDENING-20260904-001/review.md`
- 执行状态：当前 fixScope 没有开放的 `MUST_FIX`；本轮无需修改代码。

#### FIX-004

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：本轮仅处理 `MUST_FIX`；该项为 `SHOULD_FIX`，不在本轮范围内。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：本轮仅处理 `MUST_FIX`；该项为 `SHOULD_FIX`，不在本轮范围内。

### Round 4 汇总

- MUST_FIX：无。
- 已完成：无代码修复项；`FIX-001`、`FIX-002`、`FIX-003` 保持已完成状态。
- 跳过：`FIX-004`、`FIX-005`，因本轮使用 `fixScope=must` 且用户要求仅处理 `MUST_FIX`。
- 修改文件：仅更新本 `execution.md`；未修改 `review.md`、业务代码或测试代码。
- 验证：无代码变更，不重复运行无关测试；Review Fix 状态注册已通过。
- 下一步：如需处理剩余 `SHOULD_FIX`，应使用 `fixScope=recommended`，然后再次进行独立 Review。
