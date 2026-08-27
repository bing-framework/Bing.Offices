<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260825-01
AI_EXECUTION_FINISHED_AT: 2026-08-26T14:39:47.538Z

AI_EXECUTION_STARTED_AT: 2026-08-26T08:34:08.263Z

| RH-202 API 收敛 | BLOCKED | 按计划不删除未批准的公开 API；当前 2.x 保持兼容。 |
| RH-301 内部协作边界 | PARTIAL | 已沿用既有 importer/materializer/validation/failure writer/sheet writer 拆分；未完成全量测试文件分解。 |
| RH-302 Mapping/cache 审计 | PARTIAL | 既有 patch/reset 测试保持通过；完整字段、缓存隔离和并发契约未完成。 |
| RH-401 Regex 缓存 | PARTIAL | 已改为容量 256 的有界缓存并保留 1 秒 timeout；缺少专门容量/并发测试。 |
| RH-402 导入热点 | PARTIAL | 已完成 UniqueTracker pending 计数和图片行索引；ValidationRangeIndex 优化及性能证据未完成。 |
| RH-403 导出热点/样式缓存 | PARTIAL | 已预计算动态 key 集合、移除空字典分配并修复 XSSF 边框颜色；HeaderAttribute 样式统一缓存未完成。 |
| RH-404 基准治理 | PARTIAL | 已移除伪 legacy benchmark；真实历史基线、retained capacity 和资源探针重构未完成。 |
| RH-501 测试强化 | PARTIAL | 已增加 P0、失败产物和 XSSF 边框回归；完整矩阵和生产符号追溯表未完成。 |
| RH-502 文档 | PARTIAL | 修复失效索引、major 占位和 Workbook-first 描述；完整迁移示例和样式/资源文档未完成。 |
| RH-601 build/test/pack/consumer | COMPLETED | 当前可执行矩阵全部通过，见下文。 |
| RH-602 人工互操作 | BLOCKED | 当前环境未提供 Excel、LibreOffice、WPS 的可审计验证证据。 |
| RH-603 发布结论 | COMPLETED | 已明确 `NO-RELEASE`，未执行发布。 |

## 已完成事项

- Workbook 原生整数比较操作符 `LESS_OR_EQUAL` 改为只使用 `Formula1`。
- Workbook 原生校验读取 `IDataValidation.EmptyCellAllowed`，并保持 Workbook 失败时跳过后续转换和配置校验。
- `ExcelDateAttribute.Format` 指定时基于原始规范化文本执行 `TryParseExact`。
- 失败工作簿使用临时文件和 `LimitedWriteStream`，序列化期间执行 `MaxBytes` 限制，成功后才复制到调用方目标流。
- 失败工作簿复制增加行属性、列宽/隐藏、冻结窗格、超链接、批注、富文本文本值和数据校验显示属性的保真处理。
- `UniqueTracker` 使用增量 `_pendingValueCount`，图片空行判断使用按行索引集合。
- 导出动态列合法 key 集合改为每个 Sheet 预计算，缺失 dynamic values 不再创建空字典。
- Regex 缓存改为容量上限为 256 的锁保护缓存，并保留正则表达式 timeout。
- 删除不代表历史实现的伪 legacy stream benchmark 方法。
- 修复 XSSF 自定义四边框 RGB 写入；HSSF 自定义颜色仍按明确能力边界拒绝。
- 修正文档中的失效设计索引、`MIGRATION_CURRENT_MAJOR` 占位和 Workbook 校验 Continue 描述。

## 部分/未完成事项

- `ExcelCellStyle` 仍使用非 nullable enum 默认值，无法完整区分未指定和显式 reset；未在未批准 breaking change 前改变公共 DTO。
- ErrorRowsOnly 尚未覆盖条件格式、命名区域、全部富文本格式运行、打印设置等完整对象清单；当前文档不应宣称完全保真。
- `ExcelMapping.For<T>()`、旧 `Mapping(...)` 重载、布尔语义方法、`AddNavigationSheet`、`ExcelSetting.Default` 和 `ICellValueConverter` 未删除或重命名，原因是没有批准的 next-major 策略。
- 没有真实历史 NuGet/tag/独立程序集可作为性能对照；本轮不能给出相对历史版本的性能提升结论。
- 未完成 Office、LibreOffice、WPS 的打开/保存/重开和截图/哈希证据。
- 用户指定的 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260825-220502.md` 不存在，未引用或推断其结论。

## 修改文件

- `.agents/runtime/current-task.json`
- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`
- `docs/excel/README.md`
- `docs/excel/import-validation.md`
- `docs/excel/nuget-migration.md`
- `src/Bing.Offices.Abstractions/Bing/Offices/Providers/UniqueTracker.cs`
- `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiExportSheetWriter.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiImportRowMaterializer.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiWorkbookValidationPipeline.cs`
- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
- 本执行报告及同目录既有 `plan.md`

## API/数据/配置变化

- 未删除或重命名公开 API。
- 未改变数据库、外部服务或持久化数据。
- 失败工作簿输出内部改用临时文件；调用方目标流仍由调用方拥有，超限时在复制到目标流前失败。
- XSSF 自定义边框颜色现在使用 RGB 通道；HSSF 仍仅支持可映射的 indexed palette 颜色。
- 当前 major 仍为 `2.x`，迁移文档明确记录兼容 API 和未批准的 next-major breaking 事项。

## 测试结果

| 命令 | 结果 |
|---|---|
| `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore` | 通过，224/224 |
| `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore` | 通过，224/224 |
| P0 定向测试（net8） | 通过，32/32；失败工作簿、Workbook 校验、日期和结构保真覆盖 |
| Excel 请求/样式与 P0 定向测试（net8） | 通过，66/66 |
| `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore` | 通过，12/12 |
| `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore` | 通过，12/12 |
| `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore` | 通过，8/8 |
| `dotnet test ... --filter "ReviewFixRegressionTest|ExcelP0RegressionTest"` | 通过，52/52 |

测试输出存在既有弃用警告，包括旧 attributes、`ICellValueConverter`，以及 netcoreapp3.1 对部分 System.Security 程序包的 TFM 支持警告；没有因这些警告失败。

## Build/Typecheck/Lint/Format

- `dotnet build Bing.Offices.sln -c Release --no-restore`：通过。
- `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore`：通过。
- `dotnet pack Bing.Offices.sln -c Release --no-build --no-restore`：通过，生成 `Bing.Offices.Abstractions.2.0.0.nupkg`、`Bing.Offices.Core.2.0.0.nupkg`、`Bing.Offices.Npoi.2.0.0.nupkg`。
- changed C# 文件诊断：未发现错误。
- `git diff --check`：无空白错误；Git 仅提示工作区 CRLF 在后续 Git 操作中可能转换为 LF。
- 未运行专用 lint/formatter，仓库未提供独立门禁命令。

## 计划偏差

1. 计划要求先批准 breaking table 再实施 API 删除；当前没有批准，因此保留所有兼容入口并把迁移边界写入文档。
2. 计划要求完整人工办公套件证据；当前环境未提供客户端，按 release blocker 记录。
3. 计划要求真实历史性能 baseline；未找到可审计历史包/tag/独立实现，因此只移除伪 legacy 对照，不声称性能提升。
4. 计划原列出多个 `ai_docs/excel/*.md` 设计文件，但目录实际不存在，README 已改为现有 `docs/excel` 页面索引。

## 基线问题

- 工作区在执行前已有用户/任务相关变更；本轮未回退或清理。
- `.agents/runtime/current-task.json` 是任务状态脚本的必要变更。
- `tests/Bing.Offices.ProfileFixtures/Bing.Offices.ProfileFixtures.xml` 在 Git 警告中出现，但未被本轮修改。
- 外部实现审查报告缺失，无法作为证据输入。

## 已知问题

- 2.x 仍包含旧 API 和过时属性，当前只保留兼容行为，不能视为 API 收敛完成。
- NPOI 仍是 DOM 管线，不是零 GC 或真正 streaming XLSX 引擎。
- HSSF 不能表达所有 XSSF ARGB 自定义颜色；不支持颜色会抛出明确 `NotSupportedException`。
- ErrorRowsOnly 的结构复制已增强，但尚未覆盖所有 Excel 对象，不能宣称任意复杂工作簿完整保真。
- 人工互操作结果为空，因此不能据此宣称 Excel/LibreOffice/WPS 兼容。

## 风险与回归关注点

- 失败工作簿依赖 NPOI 对 Stream 的实际写入路径；当前已用 net6/net8 回归，但其他未运行 TFM 仍需验证。
- 临时文件删除失败会被吞掉以避免覆盖主异常，部署环境应监控临时目录残留。
- Regex 缓存锁保证容量边界，但没有专门的容量观测接口和基准证据。
- 模板样式显式 reset 仍受公共 DTO 默认值限制，后续 major 设计必须补充序列化、缓存 key 和模板组合测试。

## Reviewer 注意事项

- 重点审查 `NpoiFailureWorkbookWriter` 的跨 Workbook 样式克隆、冻结窗格行映射、批注/超链接复制和 NPOI 异常解包。
- 重点审查 `NpoiStyleCache` 的 XSSF/HSSF 分支，确认 HSSF 不支持颜色时的异常属于公开且文档一致的行为。
- 重点审查 `LimitedWriteStream` 在 NPOI 其他写入方法或未来 span/async 路径中的限制覆盖。
- 发布前必须补齐 RH-104、RH-201/RH-202、RH-404、RH-602，并重新运行完整包消费者和互操作矩阵。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE

#### FIX-001

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs`
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
- 根因：样式叠加缺少显式 reset 表达，且填充前景色与背景色共用语义。
- 修复：新增可序列化的 `ExcelCellStyleReset`，将 reset 标志纳入样式缓存键；Compose 按属性保留模板值或恢复默认值；XSSF 前景/背景分别写入；HSSF 不可表示的自定义颜色明确抛出 `NotSupportedException`。
- 验证：
	- `Export_RequestStyle_Compose_ShouldPreserveAndResetSelectedProperties`：PASS
	- `Export_RequestStyle_ShouldKeepXlsxForegroundAndBackgroundIndependent`：PASS
	- `Export_RequestStyle_Xls_ShouldHonorColorCapabilityBoundary`：PASS
	- `dotnet test ... --framework net8.0` 定向集合：97/97 PASS
	- `dotnet test ... --framework net6.0` 全量：249/249 PASS
	- `dotnet test ... --framework net8.0` 全量：249/249 PASS

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
	- `docs/excel/import-validation.md`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 根因：失败工作簿对象边界和既有批注冲突语义未形成可验证合同，富文本格式运行复制及 HSSF 行属性兼容处理不完整。
- 修复：新增 `Preserve/Append/Replace/Fail` 批注冲突策略，默认 `Append` 保持兼容；已有批注在原对象上更新，避免同一单元格创建第二个批注；补充 XSSF/HSSF 富文本 formatting run 复制；对 HSSF 未实现的行属性读取按能力边界处理；文档列出支持、部分支持和不承诺复制的对象矩阵。
- 验证：
	- `Import_AnnotatedFailureWorkbook_ShouldApplyCommentConflictPolicy`：PASS
	- `Import_ErrorRowsOnly_ShouldPreserveRichTextRunsAfterReopen`（XLSX/XLS）：PASS
	- 既有 ErrorRowsOnly 结构复制测试：PASS
	- `dotnet build Bing.Offices.sln -c Release --no-restore`：PASS

#### FIX-003

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiWorkbookValidationPipeline.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 根因：比较操作符使用裸编码，且 HSSF 将数值边界保存在 `DVConstraint.Value1/Value2` 而非 Formula 字段。
- 修复：使用 NPOI 公共 `OperatorType` 语义常量；补充 HSSF 数值边界回退读取；覆盖 Between、NotBetween、Equal、NotEqual、GreaterThan、LessThan、GreaterOrEqual、LessOrEqual 的 XSSF/HSSF 真实 NPOI fixture，并保留 Workbook 失败不物化语义。
- 验证：
	- `Import_WorkbookNumericValidation_ShouldSupportAllOperators`：XSSF/HSSF 全矩阵 PASS
	- `Import_WorkbookLessOrEqual_ShouldUseFormula1Only`：PASS
	- `Import_WorkbookValidation_EmptyCellAllowed_ShouldControlEmptyValue`：PASS
	- `dotnet test ... --framework net8.0` 全量：249/249 PASS
	- `dotnet test ... --framework net6.0` 全量：249/249 PASS

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：BLOCKED
- 修改文件：
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/execution.md`
- 阻塞证据：只读检查显示 `excel`、`soffice`、`libreoffice`、`wps`、`et` 命令均不可用；环境中未发现可执行的 Excel、LibreOffice Calc 或 WPS 客户端。
- 未执行事项：没有伪造客户端版本、打开/保存/重开结果、截图或 fixture hash；因此无法完成 XLSX/XLS 的真实办公软件互操作矩阵。
- 解除条件：提供可运行的 Excel、LibreOffice Calc、WPS 客户端及版本信息；对每个客户端执行 XLSX/XLS 打开、保存、重开，记录 fixture hash、结构复解析和必要截图。
- 发布影响：FIX-004 保持 BLOCKED，`RH-602` 未关闭；整体继续保持 `PARTIAL` 和 `NO-RELEASE`。

### Round 1 汇总

- MUST_FIX：4
- 已完成：FIX-001、FIX-002、FIX-003
- PARTIAL：无（代码修复项）
- BLOCKED：FIX-004，外部办公客户端不可用
- FAILED：无
- 回归验证：net8 Unit 249/249、net6 Unit 249/249、P0 定向 97/97、Release build PASS、诊断无错误、`git diff --check` PASS
- 备注：构建仍有仓库既有弃用警告及旧 TFM 支持警告；未发现本轮新增错误。
- 下一步：交由独立 Reviewer 再次验收；在人工互操作 blocker 解除前不得发布。

## Git 状态

终态检查显示存在本任务的未提交修改和新建任务目录；未发现本轮需要处理的无关文件变更。没有执行自动 `git add`、`git commit`、`git push`、reset、clean、tag 或 PR 操作。

### Round 2

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE
- 说明：本轮未修改 `review.md`；独立 Reviewer 仍需重新验收。

#### FIX-001

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs`
- 根因：上一轮实现虽已纳入字体 reset 缓存键并将 XSSF 空颜色写为无色，但缺少调用顺序隔离、多属性 reset 和构建警告的直接证据。
- 修复/补强：补充字体 reset 前后调用顺序测试；补充填充、四边框、水平/垂直对齐、换行、缩进和数字格式 reset 测试；确认 `ExcelCellStyle.FontSize` XML 注释格式正确，XSSF reset 保持 `null` 而不是黑色 RGB。
- 验证：
	- `Export_RequestStyle_FontReset_ShouldIsolateFontCacheEntries`：PASS
	- `Export_RequestStyle_Reset_ShouldRestoreAllCoveredProperties`：PASS
	- `Export_RequestStyle_XlsxFillColorReset_ShouldClearColors`：PASS
	- net6 Unit：269/269 PASS
	- net8 Unit：269/269 PASS
	- `dotnet build Bing.Offices.sln -c Release --no-restore`：0 errors；未再出现 `CS1570`

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
- 根因：需要证明 HSSF formatting run 使用源 Workbook 的字体表，而不是按相同索引解释目标 Workbook 字体。
- 修复/补强：保留生产实现从 `sourceWorkbook` 获取 HSSF run 字体并克隆到目标 Workbook；测试 fixture 预创建多个未使用字体，使源字体索引非连续，并在 XLS/XLSX 重开后分别断言 bold/italic run 属性。
- 验证：
	- `Import_ErrorRowsOnly_ShouldPreserveRichTextRunsAfterReopen(false)`：PASS
	- `Import_ErrorRowsOnly_ShouldPreserveRichTextRunsAfterReopen(true)`：PASS
	- Round 2 定向集合（`ExcelP0RegressionTest` + `ExcelWorkbookRequestTest`）：111/111 PASS
	- net6 Unit：269/269 PASS
	- net8 Unit：269/269 PASS
	- Release build：PASS

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiWorkbookValidationPipeline.cs`
- 根因：上一轮生产比较逻辑已有语义化操作符和 HSSF 边界回退，但需要补齐各约束类型及模式的直接矩阵证据。
- 修复/补强：确认并保留 decimal、date、time、text-length 的 XSSF/HSSF 测试；覆盖单边 Formula2 缺失、区间边界、`EmptyCellAllowed`、Continue、StopOnFirstFailure，以及 Workbook 失败时不物化。
- 验证：
	- `Import_WorkbookDecimalValidation_ShouldSupportAllOperators`：XLSX/XLS PASS
	- `Import_WorkbookTypedValidation_ShouldSupportSingleBoundAndRange`：XLSX/XLS PASS
	- `Import_WorkbookValidation_EmptyCellAllowed_ShouldControlEmptyValue`：PASS
	- `Import_WorkbookValidation_ShouldRespectContinueAndStop`：PASS
	- `Import_WorkbookValidationFailure_Continue_ShouldSkipMaterialization`：PASS
	- Round 2 定向集合：111/111 PASS
	- net6 Unit：269/269 PASS
	- net8 Unit：269/269 PASS

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改/证据文件：
	- `artifacts/interop-round2/run-excel-interop.ps1`
	- `artifacts/interop-round2/excel/manifest.json`
	- `artifacts/interop-round2/excel/interop-results.txt`
	- `artifacts/interop-round2/excel/failure-fixture-xlsx.xlsx`
	- `artifacts/interop-round2/excel/failure-fixture-xlsx-roundtrip.xlsx`
	- `artifacts/interop-round2/excel/failure-fixture-xls.xls`
	- `artifacts/interop-round2/excel/failure-fixture-xls-roundtrip.xls`
	- `artifacts/interop-round2/excel/failure-fixture-xlsx.png`
	- `artifacts/interop-round2/excel/failure-fixture-xls.png`
	- `artifacts/interop-round2/Program.cs`
	- `artifacts/interop-round2/InteropCheck.csproj`
- 已完成：通过 Excel COM `16.0`、Build `17932` 对 XLSX/XLS fixture 执行打开、保存、重开；记录 source/roundtrip SHA-256、截图状态和客户端路径。独立 net8/NPOI 复解析确认两种格式均保留公式 `1+1`、1 个合并区域、批注、1 个数据校验、绘图对象和 `_ImportErrors` Sheet。
- 当前阻塞：重新检查注册表/常见安装目录和用户目录后，LibreOffice 与 WPS 均未发现；没有伪造其版本或结果。
- 发布影响：RH-602 仍未完全关闭，继续保持 `NO-RELEASE`。解除条件是提供可运行的 LibreOffice/WPS 客户端并对相同 fixture 执行打开、保存、重开、哈希、复解析和截图。

### Round 2 汇总

- MUST_FIX：4
- 已完成：FIX-001、FIX-002、FIX-003
- PARTIAL：FIX-004，Excel 已完成，LibreOffice/WPS 缺失
- BLOCKED：FIX-004 的 LibreOffice/WPS 子矩阵
- FAILED：无
- 回归验证：Round 2 定向 111/111；net6 Unit 269/269；net8 Unit 269/269；Release build 0 errors/9 existing dependency warnings；`git diff --check` 无空白错误
- review.md 完整性：SHA-256 `61CD74216229443605D15C8696AE11B6C5F8157F335D35AEA2ED4D1D92539CF7`，修复过程中未修改
- 下一步：交由独立 Reviewer 再次验收；在 FIX-004 完整关闭前不得发布。

## Review 修复记录

### Round 3

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`（本轮未修改）

#### FIX-001

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
- 修复：通过公开 `ExcelExport.Workbook(...)` 入口覆盖固定列与动态列的 Header/Body 样式；在全新 XSSFWorkbook/HSSFWorkbook 中分别执行 reset-first 与 preserve-first；XLSX/XLS 重开后断言字体、填充、对齐、数字格式和动态列样式，并验证 style/font 数量固定上界。
- 验证：
	- `Export_RequestStyle_PublicFixedAndDynamicColumns_ShouldResetAndBoundResources`：XLSX/XLS PASS
	- `Export_RequestStyle_ResetOrder_ShouldIsolateFonts`：XLSX/XLS、两种调用顺序 PASS
	- net8 定向测试：171/171 PASS
	- net6 定向测试：171/171 PASS

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 修复：补齐 date、time、text-length 的 Between、NotBetween、Equal、NotEqual、GreaterThan、LessThan、GreaterOrEqual、LessOrEqual 正反边界；每个用例均在 XLSX/XLS 执行；所有单边操作符均断言 Formula2 为空。
- 验证：
	- `Import_WorkbookTypedValidation_ShouldSupportAllOperators`：date 32、time 32、text-length 32 个理论用例；每个用例覆盖 XLSX/XLS，共 192 次格式断言
	- net8 定向测试：171/171 PASS
	- net6 定向测试：171/171 PASS
	- net8 Unit 全量：323/323 PASS
	- net6 Unit 全量：323/323 PASS

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED（产品来源与 Excel 子矩阵）；LibreOffice/WPS 子矩阵 BLOCKED
- 修改/证据文件：
	- `artifacts/interop-round2/Program.cs`
	- `artifacts/interop-round2/InteropCheck.csproj`
	- `artifacts/interop-round2/run-excel-interop.ps1`
	- `artifacts/interop-round3/excel/product-input-manifest.json`
	- `artifacts/interop-round3/excel/manifest.json`
	- `artifacts/interop-round3/excel/export-product-xlsx.xlsx`
	- `artifacts/interop-round3/excel/export-product-xls.xls`
	- `artifacts/interop-round3/excel/failure-product-xlsx.xlsx`
	- `artifacts/interop-round3/excel/failure-product-xls.xls`
- 修复：临时 net8 宿主通过 Bing.Offices 的真实 `NpoiExcelExporter` 生成 XLSX/XLS 导出工作簿，并通过 `NpoiExcelImporter` 配置 `ErrorRowsOnly` 生成 `NpoiFailureWorkbookWriter` 可观察的失败工作簿；Excel COM 脚本仅打开、保存、重开这些产品来源文件，不再使用 `Workbooks.Add()` 生成输入。
- 验证：
	- 产品来源生成命令：`dotnet run --project artifacts/interop-round2/InteropCheck.csproj --no-restore`：PASS
	- Excel `16.0`、Build `17932`：export/failure 两类、XLSX/XLS 共 4 个输入均打开、保存、重开 PASS
	- 4 个截图状态均为 `PASS`；source/roundtrip SHA-256 已写入 `artifacts/interop-round3/excel/manifest.json`
	- 失败工作簿两种格式均保留转换错误文本，并由 NPOI 二次复解析成功
	- LibreOffice/WPS：按注册表、常见安装目录和用户目录探测，当前环境未发现，子矩阵保持 BLOCKED；发布结论继续 `NO-RELEASE`

### Round 3 汇总

- MUST_FIX：FIX-001、FIX-003、FIX-004
- 已完成：FIX-001、FIX-003；FIX-004 的 Bing.Offices 产品来源和 Excel COM 子矩阵
- PARTIAL：无代码修复项
- BLOCKED：FIX-004 的 LibreOffice/WPS 子矩阵，因客户端未安装
- FAILED：无
- 回归验证：net8/net6 定向各 171/171；net8/net6 Unit 各 323/323；Release build 通过；`git diff --check` 无空白错误
- 发布结论：`NO-RELEASE`；未执行 commit、push、tag、publish 或 PR
- 下一步：交由独立 Reviewer 再次验收；`review.md` 仍保持独立的 `NEEDS_FIX` 证据。

### Round 4

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`（本轮未修改）
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE

#### FIX-001

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
- 根因：上一轮公共入口样式测试从新建工作簿默认样式开始，无法证明已有模板前景/背景、边框、对齐和数字格式被 reset。
- 修复：新增 `Export_RequestStyle_PublicReset_ShouldClearNonDefaultTemplateProperties`，用真实 XLSX/HSSF 模板预置非默认字体、前景/背景、填充、四边框、水平/垂直对齐、换行、缩进和数字格式；通过公开 `ExcelExport.Workbook(...)` 分别覆盖 `SheetStyle`、固定列 Header/Body 和动态列 Header/Body reset；序列化重开后断言全部属性恢复默认，并保留 style/font 上界断言。
- 验证：
	- net8 专项测试：XLSX/XLS 2/2 PASS
	- net6 专项测试：XLSX/XLS 2/2 PASS
	- net8 相关定向集合：173/173 PASS
	- net6 相关定向集合：173/173 PASS
	- net8 Release Unit：325/325 PASS
	- net6 Release Unit：325/325 PASS

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改/证据文件：
	- `artifacts/interop-round2/Program.cs`
	- `artifacts/interop-round2/run-excel-interop.ps1`
	- `artifacts/interop-round3/excel/product-input-manifest.json`
	- `artifacts/interop-round3/excel/interop-blocked.json`
	- `artifacts/interop-round3/excel/npoi-roundtrip-parse.json`
	- `artifacts/interop-round3/excel/export-product-xlsx.xlsx`
	- `artifacts/interop-round3/excel/export-product-xls.xls`
	- `artifacts/interop-round3/excel/failure-product-xlsx.xlsx`
	- `artifacts/interop-round3/excel/failure-product-xls.xls`
- 根因：上一轮产品 fixture 过于简单，且 NPOI roundtrip 复解析没有可运行入口；Excel COM 阶段还必须区分真实客户端阻塞与成功结果。
- 修复：
	- 通过生产 `NpoiExcelExporter`/`NpoiExcelImporter` 生成四个产品源文件；输入 manifest 只列四个 source，不再混入旧 roundtrip 文件。
	- 导出模板加入公式、批注、合并区域和数据验证；XLSX 产品加入真实图表；XLS 使用公开支持的 indexed 颜色，XLSX 保留自定义 ARGB，避免静默跨 Provider 降级。
	- 新增 `--verify-roundtrip` NPOI 宿主入口，存在 roundtrip 文件时断言公式、批注、合并、数据验证、XLSX 图表和 `_ImportErrors` 错误文本；文件缺失时生成明确 `Status=BLOCKED` 的 `npoi-roundtrip-parse.json`。
	- COM 脚本记录真实公式、批注、验证、合并、图形、前景色、字体颜色、版本和哈希；打开失败时生成 `interop-blocked.json`，不伪造 manifest 或截图结果。
- 验证：
	- `dotnet run --project artifacts/interop-round2/InteropCheck.csproj --no-restore`：PASS，四个 Bing.Offices 产品源文件生成成功。
	- 互操作宿主编译：PASS；仅有既有 nullable/旧 TFM 警告。
	- `dotnet run ... -- --verify-roundtrip`：按缺失 Excel roundtrip 文件输出 `Status=BLOCKED`，未伪造结构通过。
	- Excel `16.0` Build `17932` COM 探测成功，但 `Workbooks.Open` 在当前 PowerShell COM 宿主中所有参数形式均失败，`interop-blocked.json` 记录错误“不能取得类 Workbooks 的 Open 属性”。
	- LibreOffice/WPS 命令和安装路径仍未发现，继续 BLOCKED。
- 未完成：当前没有 Excel 保存后的 roundtrip 文件，因此不能宣称 Excel 打开/保存/重开或 NPOI 二次结构断言通过；`FIX-004` 保持 PARTIAL/BLOCKED。

### Round 4 汇总

- MUST_FIX：FIX-001、FIX-004
- 已完成：FIX-001
- PARTIAL：FIX-004，产品 fixture 和可运行 NPOI 验证入口完成，Excel COM roundtrip 被当前宿主阻塞
- BLOCKED：FIX-004 的 Excel COM roundtrip、LibreOffice/WPS 子矩阵
- FAILED：无
- 回归验证：net6/net8 相关定向各 173/173；net6/net8 Release Unit 各 325/325；Release build 0 errors/234 warnings；`git diff --check` 无空白错误
- 发布结论：`NO-RELEASE`；未执行 commit、push、tag、publish 或 PR
- 下一步：交由独立 Reviewer 再次验收；`review.md` 保持独立的 `NEEDS_FIX` 证据。

### Round 5

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE
- 说明：本轮仅处理 `MUST_FIX` 的 `FIX-001` 和 `FIX-004`；未修改 `review.md`，未处理 `FIX-005` 至 `FIX-007`。

#### FIX-001

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
- 根因：公共 reset 测试已从非默认模板开始，但没有直接断言字体名称、字号、字体颜色以及 XLSX/HSSF 前景和背景恢复为 provider 默认。
- 修复：扩展 `AssertResetStyle`，比较 reset 后字体的名称、字号、颜色、粗体、斜体和下划线与 Workbook 默认字体；XSSF 断言前景/背景颜色为空，HSSF 断言前景/背景为 `Automatic`。保留 Sheet、固定列、动态列、XLSX/XLS 重开和资源上界覆盖。
- 验证：
	- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj --no-restore -c Release --framework net8.0 --filter "FullyQualifiedName~ExcelWorkbookRequestTest"`：PASS，48/48。
	- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj --no-restore -c Release --framework net6.0 --filter "FullyQualifiedName~ExcelWorkbookRequestTest"`：PASS，48/48。
	- 修改文件诊断：PASS，无错误。

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改/证据文件：
	- `artifacts/interop-round2/Program.cs`
	- `artifacts/interop-round2/run-excel-interop.ps1`
	- `artifacts/interop-round5/excel/product-input-manifest.json`
	- `artifacts/interop-round5/excel/interop-blocked.json`
	- `artifacts/interop-round5/excel/npoi-roundtrip-parse.json`
	- `artifacts/interop-round5/excel/export-product-xlsx.xlsx`
	- `artifacts/interop-round5/excel/export-product-xls.xls`
	- `artifacts/interop-round5/excel/failure-product-xlsx.xlsx`
	- `artifacts/interop-round5/excel/failure-product-xls.xls`
- 根因：上一轮互操作目录混有旧 PASS 与新 BLOCKED 产物，fixture 未包含实际图片，verifier 在首个缺失格式处提前返回，且当前 Excel COM 宿主仍无法打开产品文件。
- 修复：
	- 使用独立 `interop-round5/excel` 目录和 `FixtureId`，避免与旧代际 manifest、截图和 roundtrip 混用。
	- 通过生产 `NpoiExcelExporter`/`NpoiExcelImporter` 生成四个产品源文件；导出模板加入公式、批注、合并区域、数据验证、XLSX 图表和 XLSX/XLS 均可读取的真实 PNG 图片；manifest 增加每个源产品的 NPOI 结构摘要。
	- `--verify-roundtrip` 在 roundtrip 不完整时汇总全部四个缺失文件，并输出带 FixtureId 的 `BLOCKED` 结果；完整时断言公式、批注、合并、数据验证、图片、XLSX 图表和 `_ImportErrors` 错误文本。
	- Excel 脚本改用 Round 5 目录，并在 BLOCKED 输出中记录 FixtureId。
- 验证：
	- Round 5 产品生成：PASS；manifest 记录生产 exporter/importer，导出 XLSX/XLS 均包含 Formula=`1+1`、Comment、1 个合并区域、1 个数据验证和 1 个图片；失败 XLSX/XLS 均包含 `_ImportErrors` 及转换错误文本。
	- `dotnet run --project artifacts/interop-round2/InteropCheck.csproj --no-restore -c Release -- --verify-roundtrip`：PASS（命令完成）；因四个 roundtrip 文件均不存在，输出 `Status=BLOCKED` 并完整列出四个缺失路径，未伪造结构通过。
	- Excel `16.0` Build `17932` COM：BLOCKED；`Workbooks.Open` 仍报“不能取得类 Workbooks 的 Open 属性”，Round 5 `interop-blocked.json` 未生成 PASS、截图或哈希。
	- LibreOffice/WPS：BLOCKED；当前环境仍未发现可运行客户端。
- 未完成：当前没有 Excel 保存后的四组 roundtrip 文件，不能宣称 Excel 打开/保存/重开或 NPOI 二次结构断言通过；`FIX-004` 保持 PARTIAL/BLOCKED。

### Round 5 汇总

- MUST_FIX：FIX-001、FIX-004
- 已完成：FIX-001
- PARTIAL：FIX-004；独立代际 fixture、图片覆盖、完整缺失清单和可运行 verifier 已完成，Excel COM 与 LibreOffice/WPS 仍阻塞
- BLOCKED：FIX-004 的 Excel COM roundtrip、LibreOffice/WPS 子矩阵
- FAILED：无
- 回归验证：net8/net6 全量 Unit 各 325/325；net8/net6 Integration 各 12/12；Release build 0 errors/144 warnings；`git diff --check` 无空白错误，仅有行尾转换提示
- 发布结论：`NO-RELEASE`；未执行 commit、push、tag、publish 或 PR
- 下一步：交由独立 Reviewer 再次验收；`review.md` 保持 Round 4 的 `NEEDS_FIX` 独立证据。

### Round 6

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE
- 说明：本轮仅处理 `MUST_FIX` 的 `FIX-004`；未修改 `review.md`，未处理 `FIX-005` 至 `FIX-007`。

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改/证据文件：
	- `artifacts/interop-round2/Program.cs`
	- `artifacts/interop-round2/run-excel-interop.ps1`
	- `artifacts/interop-round5/excel/product-input-manifest.json`
	- `artifacts/interop-round5/excel/interop-blocked.json`
	- `artifacts/interop-round5/excel/npoi-roundtrip-parse.json`
	- `artifacts/interop-round5/excel/export-product-xlsx.xlsx`
	- `artifacts/interop-round5/excel/export-product-xls.xls`
	- `artifacts/interop-round5/excel/failure-product-xlsx.xlsx`
	- `artifacts/interop-round5/excel/failure-product-xls.xls`
- 根因：Round 5 已有 source SHA，但客户端 BLOCKED 记录与最终 source 产品生成不具备强制的代际和哈希绑定；NPOI verifier 也未拒绝 source 被替换的情况。
- 修复：
	- 生成器为每批四个 source 创建唯一 `GenerationId`，并在 `product-input-manifest.json` 为每个 source 记录 SHA-256。
	- Excel 脚本在打开前读取并校验同一 manifest 和全部 source SHA；每次实际尝试打开的 source 记录 `Kind`、`Format`、路径和 SHA-256；成功/阻塞输出记录 `GenerationId`、source manifest 生成时间和 manifest SHA-256。
	- NPOI `--verify-roundtrip` 读取 manifest，逐个核对 source SHA；缺失 roundtrip 时输出同一代际和完整 source 清单，成功时输出 source/roundtrip SHA-256。
- 验证：
	- `dotnet run --project artifacts/interop-round2/InteropCheck.csproj --configuration Release --no-restore`：PASS；生成 `GenerationId=56c55ad5f903479ba8bedd2d706ed2ad`，四个 source 均写入 SHA-256。
	- `dotnet run --project artifacts/interop-round2/InteropCheck.csproj --configuration Release --no-restore -- --verify-roundtrip`：PASS；因四个客户端 roundtrip 文件不存在，准确输出同一 GenerationId 的 `Status=BLOCKED` 和完整缺失清单。
	- `powershell -NoProfile -ExecutionPolicy Bypass -File artifacts/interop-round2/run-excel-interop.ps1`：BLOCKED；Excel `16.0` Build `17932` 的 `Workbooks.Open` 仍失败，但 `interop-blocked.json` 已绑定 `GenerationId`、`SourceManifestSha256`、manifest source 清单和实际 attempted source SHA。
	- `dotnet build artifacts/interop-round2/InteropCheck.csproj --configuration Release --no-restore`：PASS，0 errors / 3 warnings。
	- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj --configuration Release --no-restore --framework net8.0 --filter "FullyQualifiedName~ExcelP0RegressionTest|FullyQualifiedName~ExcelWorkbookRequestTest"`：PASS，167/167。
	- `git diff --check`：无空白错误，仅有既有 CRLF/LF 转换提示。
- 未完成：Excel 仍不能完成四个当前 source 的打开、保存、重开、截图和 roundtrip SHA；LibreOffice/WPS 仍未安装；NPOI roundtrip 结构 PASS 仍无法执行。
- 发布影响：`RH-602` 仍为 BLOCKED，`FIX-004` 保持 PARTIAL/BLOCKED，整体继续 `PARTIAL` 和 `NO-RELEASE`。

### Round 6 汇总

- MUST_FIX：FIX-004
- 已完成：source fixture 的 GenerationId/SHA 绑定、Excel/NPOI 证据代际校验和 BLOCKED 结果追踪
- PARTIAL：FIX-004，真实客户端 roundtrip 仍未完成
- BLOCKED：FIX-004 的 Excel COM roundtrip、LibreOffice/WPS 子矩阵
- FAILED：无
- 回归验证：互操作宿主 Release build PASS；相关 net8 回归测试 167/167 PASS；NPOI verifier 准确输出完整缺失清单
- 发布结论：`NO-RELEASE`；未执行 commit、push、tag、publish 或 PR
- 下一步：交由独立 Reviewer 再次验收；`review.md` 保持独立的 `NEEDS_FIX` 证据。

<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260825-01
AI_EXECUTION_FINISHED_AT: 2026-08-26T09:42:30.000Z

### Round 7 汇总

- MUST_FIX：FIX-004
- 已完成：直接 .NET COM 验证路径、只读 source 保护、GenerationId/source SHA/manifest SHA 绑定和完整缺失 roundtrip 清单
- PARTIAL：FIX-004，当前环境的 Excel COM `Workbooks.Open` 仍阻塞真实客户端 roundtrip
- BLOCKED：FIX-004 的 Excel COM roundtrip、LibreOffice/WPS 子矩阵和 NPOI 二次结构 PASS
- FAILED：无
- 回归验证：互操作宿主 Release build PASS；只读 COM 运行完成并输出结构化 BLOCKED 证据；NPOI verifier 输出四个缺失 roundtrip 文件
- 发布结论：`NO-RELEASE`；未执行 commit、push、tag、publish 或 PR

### Round 8

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`（本轮未修改）
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE
- 说明：本轮仅处理 `MUST_FIX` 的 `FIX-004`；未修改 `review.md`，未处理 `FIX-005` 至 `FIX-007`。

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改/证据文件：
	- `artifacts/interop-round2/Program.cs`
	- `artifacts/interop-round5/excel/interop-dotnet-com.json`
	- `artifacts/interop-round5/excel/npoi-roundtrip-parse.json`
- 根因：显式参数和反射式 COM 分派仍无法调用 Excel `Workbooks.Open`；为排除 MTA 线程和 late-binding 差异，本轮将验证过程放入专用 STA 线程，并改用 C# `dynamic` COM 调用，同时保留 `ReadOnly=true` 和 15 参数调用。
- 修复/验证：
	- `Program.cs` 的 `--verify-excel-com` 改为通过 `RunOnStaThread` 在 STA 线程执行，`OpenComWorkbook` 改为 `dynamic` 调用；未修改生产业务路径。
	- 继续使用当前 source generation：`GenerationId=3eb6f7c1ae9f4afb8e245d7518b66b6f`，`SourceManifestSha256=7E747ED5737D0E9FDF56717756940E3B7E0D9D299BAE34B3FECDBC4B685C4605`；四个 source SHA 与 manifest 保持一致。
	- Excel `16.0` Build `17932` 仍无法打开四个当前 source：XLSX export 返回“不能取得类 Workbooks 的 Open 属性”，XLSX failure 与两个 XLS 返回 `DISP_E_MEMBERNOTFOUND (0x80020003)`；`interop-dotnet-com.json` 为 `Status=BLOCKED`，包含四条独立结果。
	- `--verify-roundtrip` 已运行；`npoi-roundtrip-parse.json` 为 `Status=BLOCKED`，明确列出四个缺失 roundtrip 文件，未执行或伪造结构 PASS。
	- 未生成截图、roundtrip SHA 或客户端 PASS；`soffice`、`libreoffice`、`wps`、`et` 仍不可用。
- 未完成：Excel 四个 source 的打开、保存、重开、截图和 NPOI 二次解析；LibreOffice/WPS 子矩阵仍阻塞。
- 发布影响：`RH-602` 仍为 BLOCKED，`FIX-004` 保持 PARTIAL/BLOCKED，整体继续 `PARTIAL` 和 `NO-RELEASE`。

### Round 8 汇总

- MUST_FIX：FIX-004
- 已完成：STA + dynamic Excel COM 替代验证路径、ReadOnly source 保护、GenerationId/source SHA/manifest SHA 绑定和四文件独立尝试记录
- PARTIAL：FIX-004，当前 Excel COM `Workbooks.Open` 仍无法分派
- BLOCKED：FIX-004 的 Excel COM roundtrip、LibreOffice/WPS 子矩阵和 NPOI 二次结构 PASS
- FAILED：无
- 回归验证：interop host Release build 0 errors；STA/dynamic COM 运行输出四条结构化 BLOCKED 结果；NPOI verifier 输出四个缺失 roundtrip 文件；`git diff --check` 无空白错误
- 发布结论：`NO-RELEASE`；未执行 commit、push、tag、publish 或 PR
- 下一步：交由独立 Reviewer 再次验收；`review.md` 保持独立的 `NEEDS_FIX` 证据。

### Round 9

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`（本轮未修改）
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改文件：无新增互操作生产修复；复核现有 `artifacts/interop-round5/excel` 证据
- 复核：重新执行 `--verify-excel-com` 和 `--verify-roundtrip`。当前代际四个 source 的 Excel COM 结果仍全部为 `BLOCKED`，Excel `16.0` Build `17932` 的 `Workbooks.Open` 仍无法分派；NPOI verifier 仍列出四个缺失 roundtrip 文件。LibreOffice/WPS 仍不可用。
- 结论：没有生成 roundtrip、截图或 roundtrip SHA，不宣称办公软件互操作通过；`RH-602` 继续为 release blocker。

#### FIX-005

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`
	- `docs/excel/import-validation.md`
	- `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 修复：保留容量为 256 的锁保护 FIFO Regex 缓存和 1 秒 timeout；将容量设为 `internal` 测试契约；新增重复命中、容量上限、FIFO 淘汰、灾难性回溯 timeout、512 个并发模式和缓存上限测试；文档明确进程级缓存、容量、淘汰和并发语义。
- 验证：
	- `RegexCache_ShouldEnforceCapacityAndFifoEviction`：PASS
	- `RegexCache_ShouldTimeoutAndRemainBoundedUnderConcurrency`：PASS
	- net8 Unit 全量：328/328 PASS
	- net6 Unit 全量：328/328 PASS

#### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExportSheetWriter.cs`
	- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
- 修复：HeaderAttribute 字体和派生样式接入 Workbook 级缓存；保持基样式的填充、数字格式等未指定属性；导出 benchmark 每次创建新的 `MemoryStream`，不再通过复用容量隐藏 retained capacity；新增宽表头导出重开、字体/样式复用、属性保留和资源上界测试。
- 验证：
	- `Export_HeaderAttribute_ShouldReuseFontsAndStylesForWideHeaders`：PASS
	- `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --configuration Release --no-restore`：PASS
- 未完成：当前仓库仍没有可审计的 BenchmarkDotNet JSON/Markdown 结果，也没有按 failure/style/regex/unique/validation 分组的绝对性能、P95、allocated bytes 和 retained capacity 产物；不宣称 RH-404 完成。

#### FIX-007

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：DEFERRED/BLOCKED
- 修改文件：
	- `docs/excel/nuget-migration.md`
	- `tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`
- 处理：迁移文档补充当前 2.x 兼容入口、候选 next-major before/after 表、替代路径、版本策略和批准状态；包消费者测试覆盖 `ExcelMapping.For<T>()` 与已 obsolete 的 `RegexAttribute` 编译/运行路径。
- 未执行：未删除 `ExcelMapping.For<T>()`、旧 `Mapping(...)`、`ExcelSetting.Default` 或 `ICellValueConverter`，未建立删除项 negative baseline，也未做全局默认设置去除；原因是 breaking table 尚未获产品/维护者批准，直接执行会产生未经授权的 breaking change。
- 解除条件：批准 next-major breaking table、迁移版本和 `ExcelSetting.Default` 配置范围设计后，再同步 obsolete shim、negative baseline、并发隔离测试和实际 API 收敛。
- 验证：
	- `LegacyCompatibility_ExternalConsumer_ShouldKeepCurrentMajorEntrypoints`：PASS（含预期 `RegexAttribute` obsolete warning）
	- `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj --configuration Release --no-restore --framework net8.0`：9/9 PASS

### Round 9 汇总

- MUST_FIX：FIX-004，仍为 PARTIAL/BLOCKED，因 Excel COM `Workbooks.Open`、roundtrip、截图和 LibreOffice/WPS 子矩阵不可用。
- 已完成：FIX-005；FIX-006 的 HeaderAttribute 缓存、目标流生命周期和直接资源测试；FIX-007 的兼容迁移文档与包消费者当前 major 验证。
- PARTIAL：FIX-006，缺少真实 BenchmarkDotNet 结果和计划要求的分组性能/资源证据。
- DEFERRED/BLOCKED：FIX-007，等待批准的 breaking table 与配置去全局化设计。
- FAILED：无。
- 回归验证：
	- net8 Unit：328/328 PASS
	- net6 Unit：328/328 PASS
	- net8 Integration：12/12 PASS
	- net6 Integration：12/12 PASS
	- Docs consumer net8：9/9 PASS
	- `dotnet build Bing.Offices.sln --configuration Release --no-restore`：PASS
	- `dotnet pack Bing.Offices.sln --configuration Release --no-build --no-restore`：PASS
	- Benchmark project build：PASS
	- Interop host build：PASS；Excel COM 与 NPOI roundtrip verifier 保持 BLOCKED
	- `git diff --check`：无空白错误，仅有既有 CRLF/LF 转换提示
- 安全/兼容边界：未删除或重命名公开 API；未修改 `review.md`；未执行 commit、push、tag、publish、reset、clean 或 PR。
- 最终结论：`PARTIAL` / `NO-RELEASE`。
- 下一步：交由独立 Reviewer 进行 Round 9 再次验收；不得在本 Review Fix 轮次内替代 Reviewer 修改 `review.md`。

### Round 10

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review Round：10
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`（本轮未修改）
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改文件：无；复核现有 `artifacts/interop-round5/excel` 证据。
- 复核：重新执行 `--verify-excel-com` 和 `--verify-roundtrip`。当前 GenerationId 的四个 source SHA 与 manifest 一致；Excel `16.0` Build `17932` 的四次 `Workbooks.Open` 仍全部无法分派，`interop-dotnet-com.json` 保持 `Status=BLOCKED` 且包含四条结果；NPOI verifier 保持 `Status=BLOCKED` 并列出四个缺失 roundtrip 文件。
- 未完成：没有生成办公客户端 roundtrip、截图或 roundtrip SHA；LibreOffice/WPS 仍未安装，无法完成对应子矩阵。
- 结论：不伪造互操作 PASS；`RH-602` 继续为 release blocker，整体保持 `PARTIAL`/`NO-RELEASE`。

#### FIX-005

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：RESOLVED（沿用上一轮 COMPLETED 结果）
- 本轮处理：未重复修改；确认 Regex cache 容量、FIFO、timeout、并发上界和文档合同保持有效。
- 验证：net8/net6 Unit 全量在本轮分别为 329/329 PASS；既有 Regex 专项测试继续通过。

#### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED（本轮可执行范围）
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`
	- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`
	- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
	- `benchmarks/Bing.Offices.Benchmarks/Program.cs`
- 修复：
	- HeaderAttribute 宽表头测试扩展为 XLSX/XLS 两种格式；XSSF 断言 ARGB，HSSF 按 indexed palette 断言，并分别使用 Provider 固定 style/font 上限。
	- 增加 `FailureWorkbookExport`、`HeaderStyle`、`ValidationRange` 和 `RegexCacheHit` benchmark 分组；保留既有 `Import`、`Export`、`UniqueJournal` 和映射缓存分组。
	- 每次 Export benchmark 创建新的 `MemoryStream`；删除 benchmark 文件中的孤立字段注释。
	- 资源探针移除与生产路径无关的人工 LOH payload，改为记录 mapping-plan/UniqueTracker workload，保留 16 组场景、LOH 和工作集上限检查。
- 验证：
	- `Export_HeaderAttribute_ShouldReuseFontsAndStylesForWideHeaders`：net8 XLSX/XLS 2/2 PASS。
	- 同一专项：net6 XLSX/XLS 2/2 PASS。
	- `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --configuration Release --no-restore`：PASS，0 errors。
	- Stream benchmark：15 个参数组合完成，生成 `artifacts/benchmark-round10/results/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-report-full-compressed.json` 和 Markdown/CSV/HTML；包含 Mean、Error、StdDev、GC 和 Allocated。
	- Mapping benchmark：`RegexCacheHit`、`UniqueJournal` 与 `DynamicPlanBuildCacheHit` 分别完成，结果分别保存在 `artifacts/benchmark-round10-regex/results`、`artifacts/benchmark-round10-unique/results` 和 `artifacts/benchmark-round10-plan-cache/results`，每个目录均包含 `Bing.Offices.Benchmarks.MappingValidationBenchmarks-report-full-compressed.json` 及 Markdown/CSV/HTML。
	- `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --configuration Release --no-build -- --resource-probe artifacts/resource-round10.json`：PASS，16/16 场景；全部子进程 exit code 为 0，全部 status 为 passed，`payloadBytes` 字段不存在，最大工作集为 136863744 bytes。

#### FIX-007

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED（文档与当前 major 兼容范围）/DEFERRED（未批准 breaking change）
- 修改文件：
	- `docs/excel/nuget-migration.md`
	- `tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`
- 修复：
	- 修正迁移示例，保存根 `ExcelMapping` builder 后再调用 `Build()`，避免在列 builder 上调用不存在的 API。
	- 将 `nuget-migration.md` 加入 Markdown C# fence 编译执行清单，并补充该示例所需的消费者上下文。
	- 保持当前 2.x public API、obsolete 兼容层和未批准的 next-major breaking table，不擅自删除、重命名或改变全局配置语义。
- 未执行：breaking table、删除项 negative baseline、`ExcelSetting.Default` 去全局化及其并发隔离测试仍需产品/维护者批准和设计输入；按 Review 要求记录为 deferred，不伪造完成。
- 验证：
	- `DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually`：PASS，10 个 fence。
	- Docs consumer 全量 net8：9/9 PASS，保留预期 `RegexAttribute` obsolete warning。

### Round 10 汇总

- MUST_FIX：FIX-004，PARTIAL/BLOCKED，真实 Excel/LibreOffice/WPS roundtrip 仍不可执行。
- 已完成：FIX-006 的 XLSX/XLS 样式回归、代表性 benchmark 分组、JSON/Markdown/CSV/HTML 结果和真实 workload 资源探针；FIX-007 的迁移示例修正、文档 fence 覆盖和当前 2.x consumer 验证；FIX-005 保持 RESOLVED。
- PARTIAL：FIX-004 的证据代际绑定与四 source 独立尝试成立，但客户端 roundtrip、截图和二次结构 PASS 缺失。
- BLOCKED：FIX-004 的 Excel COM roundtrip、LibreOffice/WPS 子矩阵和 NPOI roundtrip verifier。
- DEFERRED：FIX-007 中未经批准的 next-major breaking change、`ExcelSetting.Default` 去全局化及并发隔离设计。
- FAILED：无。
- 回归验证：
	- net8 Unit：329/329 PASS。
	- net6 Unit：329/329 PASS。
	- net8 Integration：12/12 PASS。
	- net6 Integration：12/12 PASS。
	- Docs consumer net8：9/9 PASS。
	- Release build：PASS，0 errors；保留既有弃用和旧 TFM 支持警告。
	- Release pack：PASS。
	- benchmark project build：PASS。
	- resource probe：16/16 PASS。
	- `git diff --check`：无空白错误，仅有既有 CRLF/LF 转换提示。
- 安全/兼容边界：未删除或重命名公开 API；未修改 `review.md` 或 `plan.md`；未执行 commit、push、tag、publish、reset、clean 或 PR。
- 发布结论：`PARTIAL` / `NO-RELEASE`。
- 下一步：交由独立 Reviewer 进行 Round 10 再次验收；不得在本 Review Fix 轮次内替代 Reviewer 修改 `review.md`。

### Round 11

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review Round：11
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260825-01/review.md`（本轮未修改）
- 本轮终态：PARTIAL
- 发布结论：NO-RELEASE

#### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL/BLOCKED
- 修改文件：无；保留并复核现有 `artifacts/interop-round5/excel`、`artifacts/interop-dotnet-com.json` 和 `artifacts/npoi-roundtrip-parse.json` 证据。
- 复核：当前 GenerationId、manifest SHA 和四个 source SHA 继续一致；Excel COM `16.0` Build `17932` 的四次 `Workbooks.Open` 仍失败；四个 roundtrip 文件、截图、roundtrip SHA 和 NPOI 二次结构 PASS 仍不存在；LibreOffice/WPS 仍未发现可执行客户端。
- 未完成：无法在当前环境完成真实 Excel/LibreOffice/WPS 打开、保存、重开矩阵；未伪造客户端版本、截图、哈希或兼容 PASS。
- 结论：`RH-602` 继续为 release blocker，FIX-004 保持 BLOCKED，整体保持 `PARTIAL`/`NO-RELEASE`。

#### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED（本轮可执行范围）
- 修改文件：
	- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`
	- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
	- `benchmarks/Bing.Offices.Benchmarks/Program.cs`
- 修复：
	- `FailureWorkbookBenchmarks.FailureWorkbookExport` 使用 `FailureRowCount=1000/10000/100000` 生成逐行 `BAD-*` 输入，并通过真实 `ExcelRegex` 规则触发失败工作簿写出；不再使用固定单行 fixture。
	- `ValidationRangeBenchmarks.ValidationRange` 使用随 `ValidationRowCount=1000/10000/100000` 增长的真实 NPOI DataValidation 行区间，并构造重叠、多规则区间，实际经过 importer 的 `ValidationRangeIndex`。
	- `HeaderStyleBenchmarks` 使用 12 列 `[Header(...)]` 模型和请求级 HeaderStyle，`RowCount` 直接影响真实导出行数；保留 XLSX/XLS HeaderAttribute 功能测试证据。
	- 映射 benchmark 按职责拆为 `DynamicPlanBenchmarks`、`TenantPlanCacheBenchmarks`、`RegexCacheBenchmarks` 和 `UniqueJournalBenchmarks`，消除 Regex/Unique 携带无关 Plan/Tenant 参数；修正 10K 命名规则计划 benchmark 的显式 Import fixture。
	- `ResourceProbe` 在强制 GC 前保持 `plans`、`tracker` 和 `values` 存活；资源场景继续使用真实 mapping-plan/UniqueTracker workload。导出另有 fresh destination 的 `ExportDestinationCapacity` 基准，记录目标流容量。
- 验证：
	- `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --configuration Release --no-restore`：PASS，0 errors。
	- Failure benchmark：`artifacts/benchmark-round11-failure/results`，1000/10000/100000 三组均有 Mean、Error、StdDev、GC、Allocated，Allocated 约 28.36 MB/259.03 MB/2555.21 MB，随失败行数增长。
	- HeaderStyle benchmark：`artifacts/benchmark-round11-header-style/results`，1000/10000/100000 三组均有 Mean、Error、StdDev、GC、Allocated，Allocated 约 29.23 MB/282.24 MB/2778.86 MB，实际使用 HeaderAttribute 宽模型。
	- ValidationRange benchmark：`artifacts/benchmark-round11-validation-range-final/results`，1000/10000/100000 三组均有 Mean、Error、StdDev、GC、Allocated，Allocated 约 5.21 MB/48.04 MB/527.34 MB，实际包含多规则重叠 DataValidation 区间。
	- Mapping benchmark：`artifacts/benchmark-round11-mapping-final/results`，所有方法均有数值结果，无 `NA` 或 benchmark issue；`MultiRulePlanBuild` 约 225.8 us，10K 命名规则已进入 workload。
	- Dynamic/Tenant/Regex/Unique benchmark：分别位于 `artifacts/benchmark-round11-dynamic-plan/results`、`artifacts/benchmark-round11-tenant-cache/results`、`artifacts/benchmark-round11-regex/results`、`artifacts/benchmark-round11-unique/results`，参数只作用于对应 benchmark 类型。
	- Stream benchmark：`artifacts/benchmark-round11-stream/results`，Import/Export/ExportDestinationCapacity 覆盖 1000/10000/100000 行，结果含 Mean、Error、StdDev、GC、Allocated 和 runtime/SDK 环境字段。
	- `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --configuration Release --no-build -- --resource-probe artifacts/resource-round11.json`：PASS，16/16 场景，全部子进程 exit code 为 0；最大 `lohSampledPeakBytes`/`lohRetainedBytes` 为 56800952，最大工作集为 137113600，均低于 512 MiB/1 GiB 门限。

### Round 11 汇总

- MUST_FIX：FIX-004，PARTIAL/BLOCKED；当前环境仍无法完成真实办公客户端互操作。
- 已完成：FIX-006；Failure、HeaderAttribute style、ValidationRangeIndex、映射参数作用域、资源探针存活边界和独立 benchmark artifact 已修复并验证。
- PARTIAL：FIX-004 的代际与 SHA 绑定证据保持可靠，但 Excel/LibreOffice/WPS roundtrip、截图、roundtrip SHA 和 NPOI 二次结构 PASS 缺失。
- BLOCKED：FIX-004 的真实 Excel COM 打开保存重开、LibreOffice/WPS 子矩阵及 NPOI roundtrip verifier。
- FAILED：无。
- 回归验证：
	- net8 Unit：329/329 PASS。
	- net6 Unit：329/329 PASS。
	- net8 Integration：12/12 PASS。
	- net6 Integration：12/12 PASS。
	- Docs consumer net8：9/9 PASS。
	- `dotnet build Bing.Offices.sln -c Release --no-restore`：PASS，0 errors。
	- `dotnet pack Bing.Offices.sln -c Release --no-build --no-restore`：PASS。
	- benchmark project build：PASS，0 errors。
	- resource probe：16/16 PASS。
	- `git diff --check`：无空白错误，仅有既有 CRLF/LF 转换提示。
- 安全/兼容边界：未删除或重命名公开 API；未修改 `review.md` 或 `plan.md`；未执行 commit、push、tag、publish、reset、clean 或 PR。
- 发布结论：`PARTIAL` / `NO-RELEASE`。
- 下一步：交由独立 Reviewer 进行再次验收；不得在本 Review Fix 轮次内替代 Reviewer 修改 `review.md`。
