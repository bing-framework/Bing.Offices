<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260827-001
AI_EXECUTION_FINISHED_AT: 2026-08-27T22:18:34.5564014+08:00

# 实施执行报告

## 执行结论

Round 2 按用户指定的 `fixScope=must` 处理 FIX-001、FIX-002、FIX-005、FIX-006；FIX-003、FIX-004、FIX-007 为 SHOULD_FIX，按本轮范围延期。相关生产代码、测试、文档、API 快照、2.0.0 包和 Docs package-only consumer 已更新并完成定向/回归验证。本报告仅表示 Executor 完成当前范围，不代表 Reviewer 已通过；未执行发布。

## 任务信息

- Task ID：`BING-OFFICES-RELEASE-HARDENING-20260827-001`
- 执行模式：`review-fix`
- Review Round：`2`
- Fix Scope：`must`
- 未执行自动 `git add`、`git commit`、`git push`、Tag、NuGet publish 或 PR 创建。

## 计划执行情况

| 计划项 | 状态 | 证据 |
| --- | --- | --- |
| RH27-101 重复物理 Sheet selector | DEFERRED | FIX-004 为 SHOULD_FIX，本轮按用户要求不处理 |
| RH27-102 请求级 Workbook metadata | DONE/VERIFIED | `ExcelWorkbookMetadataOptions`、Build 快照、模板 XLS/XLSX preserve/override、API 和回归测试通过 |
| RH27-103 Excel/CSV 文件原子写入 | DONE | 统一 `AtomicFileCommitter`、flush 后取消检查和取消/失败测试通过；真实权限/磁盘故障证据待补 |
| RH27-104 Failure Workbook 诊断 | DONE | 请求级临时目录、稳定分类、可替换 IO 边界和故障注入 Unit 通过；Windows 权限目录证据待补 |
| RH27-105 资源限制真实语义 | PARTIAL | CSV 输入/行/错误/字段/列限制已实现；Excel DOM 边界已准确文档化，但 ZIP/OLE/DOM 独立进程证据待补 |
| RH27-106 配置/异常/释放 | DONE | CSV 错误分类、资源截断元数据和枚举器释放路径已覆盖；完整反射异常矩阵仍属未纳入本轮范围的剩余证据 |
| RH27-107 async/取消/公式 ADR | DONE | 取消检查、公式缓存和 NPOI DOM/不可取消边界已记录于 decisions.md；未新增伪异步 API |
| Phase 2 P0 API/发布门禁 | DONE | API diff、2.0.0 版本、三包和 Docs package-only consumer 已完成；未执行发布 |

## 已完成事项

- 删除 `ExcelSetting.Default` 静态可变实例和生产读取点；metadata 改为请求级、构建时复制的不可变配置。
- 为 Workbook 导出增加 `Metadata(...)` 入口，并验证 XLSX 核心属性写入。
- Excel/CSV `ExportToFile` 使用随机 `CreateNew` 同目录临时文件；导出失败、取消或序列化失败不会截断已有目标。
- CSV 导入记录枚举器使用 `using`，覆盖正常、异常和取消退出路径。
- Failure Workbook 临时文件清理失败通过 `ExcelImportFailureDiagnostic` 和请求级 `DiagnosticSink` 报告；无 sink 时保持可观察，诊断回调异常不会覆盖主异常。
- CSV 增加输入字节、数据行、错误数、字段长度和列数限制，并以 `CsvImportErrorCode.ResourceLimit` 与结果截断元数据表达超限。
- 模板 metadata 默认 preserve，显式 `Metadata(...)` 对 XLS/XLSX override；NPOI DOM 边界、公式缓存和取消语义已在代码与文档中统一。
- Excel 与 CSV File API 共用 `AtomicFileCommitter`，在 flush 后和不可逆提交前检查取消。
- 增加 selector、metadata、Excel/CSV 原子文件失败合同测试，并更新公共 API 类型/成员快照。
- 更新 Excel 使用与 NuGet 迁移文档，明确 metadata 和 breaking 变化。

## 部分/未完成事项

- 未完成 `RH27-000` 要求的完整 baseline、符号到测试追溯矩阵、benchmark 原始产物和全量发布清单文档。
- 已补充受控 XLSX ZIP/DOM/shared strings/styles/drawings 与 XLS OLE 独立进程探针，并记录子进程峰值工作集字段；该证据不宣称阻止所有压缩放大或 DOM 峰值。
- 未完成 API 约 180 个顶层类型的治理、唯一入口决策和职责拆分。
- 未建立独立于 Docs 项目的额外 PackageConsumer；Docs 项目已作为 package-only consumer 使用本轮 2.0.0 包通过。
- 未重跑 Benchmark、办公客户端互操作矩阵和 Windows 权限目录 Integration；net5/net7/netcoreapp3.1 全目标测试仍未完成。

## 修改文件

- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelWorkbookMetadataOptions.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelExport.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelWorkbookExportRequest.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Settings/ExcelSetting.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelStreamExtensions.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/CsvStreamExtensions.cs`
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
- `src/Bing.Offices.Npoi/ExcelHelper.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
- `tests/Bing.Offices.Tests/CsvTest.cs`
- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
- `docs/excel/README.md`
- `docs/excel/nuget-migration.md`

## API/数据/配置变化

- 新增 `ExcelWorkbookMetadataOptions`、`ExcelWorkbookExportBuilder.Metadata(...)`、`ExcelImportFailureDiagnostic` 和 `ExcelImportFailureOptions.DiagnosticSink`。
- 移除 `ExcelSetting.Default` 静态属性，属于已记录的 breaking change；实例 `ExcelSetting` 类型仍保留。
- 未修改数据库、外部服务、生产配置或持久化数据。

## 测试结果

- net8.0 Unit：`349` 通过，`0` 失败，`0` 跳过。
- net6.0 Unit：`349` 通过，`0` 失败，`0` 跳过。
- net8.0 Integration：`12` 通过，`0` 失败，`0` 跳过。
- net6.0 Integration：`12` 通过，`0` 失败，`0` 跳过。
- Docs package-only consumer：`9` 通过，`0` 失败，`0` 跳过。
- FIX 专项故障/取消、CSV、独立资源探针测试：通过。
- 公共 API 定向门禁：`6` 通过；Abstractions/Core/NPOI 快照已记录于 `api-diff.md`。
- 三个 2.0.0 nupkg 与 snupkg 已生成并检查 DLL、XML docs、nuspec；未执行发布。

## Build/Typecheck/Lint/Format

- Abstractions/Core netstandard2.0 和 NPOI net8.0 Release build：通过；存在仓库既有弃用警告。
- Unit Test net8.0 Release build：通过；存在仓库既有弃用和 xUnit analyzer 警告。
- `git diff --check`：通过。
- 未执行独立 lint/formatter；`get_errors` 无错误，`git diff --check` 通过。

## 计划偏差

- 为绕过本机 NuGet fallback 配置中不存在的 `I:\Data\VisualStudio\Shared\NuGetPackages`，验证时使用了当前会话临时盘符映射和本地 NuGet 包目录；未修改仓库或用户 NuGet 配置。
- 本轮只按用户指定范围处理 MUST_FIX；剩余独立资源/权限/consumer 证据按事实保留为 PARTIAL，未将未执行项目标为 VERIFIED。

## 基线问题

- 机器级 NuGet 配置仍引用不存在的 I 盘 fallback；正常 restore/build 可能产生 `NU1301` 或 `ResolvePackageAssets` 错误，需要环境维护者修复。
- 工作区在执行前已有其它 dirty 改动；本次未 reset、clean 或覆盖陌生改动。

## 已知问题

- Failure Workbook 已有受控文件系统故障注入 Unit，但尚无 Windows 权限目录 Integration 证据。
- `File.Replace` 的目标替换依赖操作系统文件系统合同；目标目录权限和跨卷路径仍由调用方负责。
- NPOI 仍是 DOM 导入/导出管线；本次未宣称可阻断所有压缩放大或 DOM 峰值。

## 风险与回归关注点

- `ExcelSetting.Default` 移除会影响直接访问该静态属性的消费者，迁移到每个 Workbook 请求的 `Metadata(...)`。
- 临时文件提交前的失败合同已覆盖单元测试，但多进程竞争、权限和磁盘故障仍需 Integration 取证。
- 其它目标框架、独立资源放大进程、外部数据库/缓存集成和办公客户端互操作尚未作为本次通过依据；Docs package-only consumer 已作为本轮包消费依据。

## Reviewer 注意事项

- 重点审查 `ExcelWorkbookMetadataOptions` 的 breaking API 是否符合 next-major 版本策略。
- 重点审查 `File.Replace`/`File.Move` 的 Windows 目标替换合同和临时文件清理诊断是否需要统一 Output Committer。
- 不要将本报告的 `333/333` net8.0 Unit Test 结果解释为所有目标框架、Integration 或发布门禁均已通过。

## Git 状态

- 工作区保持 dirty；包含本任务修改及执行前已有修改。
- 未自动执行 `git add`、`git commit`、`git push`、reset、clean、Tag、PR 或发布操作。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/review.md`

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelStreamExtensions.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Extensions/CsvStreamExtensions.cs`
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `tests/Bing.Offices.Tests/CsvTest.cs`
- 根因：Excel 与 CSV 文件提交各自维护 staging 逻辑，flush 后和不可逆替换前缺少取消检查。
- 修复：抽取统一 `AtomicFileCommitter`；随机同目录 staging 写入后执行 `Flush(true)`，在 flush 后及 `File.Replace/File.Move` 前检查取消，并尽力清理 staging。
- 验证：
	- Excel/CSV 预取消和写后取消测试：PASS。
	- net6/net8 Unit、Integration：PASS。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 根因：清理失败在无 sink 时被吞掉，临时目录固定且 IO 依赖不可注入。
- 修复：增加请求级 `TemporaryDirectory`、随机临时文件名、创建/序列化/复制/取消/删除分类和 `IFailureWorkbookFileSystem` 故障边界；主异常优先，清理异常通过诊断或 `Exception.Data` 保留。
- 验证：
	- 清理失败、主异常伴随清理失败、sink 抛异常测试：PASS。
	- 请求级临时目录清理测试：PASS。
	- Windows 权限目录 Integration：未完成，保留为剩余证据缺口。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelWorkbookExportRequest.cs`
	- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelExport.cs`
	- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelWorkbookMetadataOptions.cs`
	- `src/Bing.Offices.Npoi/ExcelHelper.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `docs/excel/README.md`
- 根因：模板加载路径完全忽略请求 metadata，未区分默认 preserve 与显式配置。
- 修复：Build 时记录 `MetadataSpecified`；模板默认保留原 metadata，显式 `Metadata(...)` 对 XLS/XLSX 的六个字段统一 override；请求构建时 clone 快照。
- 验证：
	- XLS/XLSX 模板默认 preserve 与显式 override：PASS。
	- net6/net8 Unit：PASS；Docs package consumer：PASS。

#### FIX-005

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修改文件：
	- `src/Bing.Offices.Abstractions/Bing/Offices/Csv/CsvImportOptions.cs`
	- `src/Bing.Offices.Abstractions/Bing/Offices/Csv/CsvImportError.cs`
	- `src/Bing.Offices.Abstractions/Bing/Offices/Csv/CsvImportResult.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvHeaderBinder.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `docs/excel/README.md`
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/decisions.md`
- 根因：CSV 资源上限和错误分类未实现，且 NPOI DOM/公式/取消边界表述不准确。
- 修复：增加 CSV 输入字节、行、错误、字段长度、列数限制及 `ResourceLimit`/截断元数据；保留 Excel 既有限制；将 NPOI 描述改为内存 Workbook DOM，并记录公式缓存和取消边界决策。
- 验证：
	- CSV 字段/行/列/输入字节限制和截断结果：PASS。
	- net6/net8 Unit、Integration：PASS。
	- ZIP/OLE/DOM 独立进程和峰值内存证据：未完成，不能宣称完整资源放大保护。

#### FIX-006

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修改文件：
	- `version.props`
	- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
	- `tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/api-diff.md`
	- `ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/package-consumer.md`
	- `docs/excel/nuget-migration.md`
- 根因：breaking API 没有同步 major 版本、before/after diff、包产物和本地 consumer 验证。
- 修复：按用户版本决策保持 `2.0.0`；记录 Abstractions/Core/NPOI 快照；生成三包及符号包；Docs 项目精确消费本轮本地 2.0.0 包。
- 验证：
	- Public API gate：`6/6` PASS。
	- 三个 nupkg/snupkg 内容检查：PASS。
	- Docs package-only consumer：`9/9` PASS。
	- 独立于 Docs 的额外临时 PackageConsumer：未建立，保留为证据缺口。

#### FIX-004 / FIX-007

- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：用户明确要求本轮只处理 `MUST_FIX`；未修改 `review.md`，下一轮可独立安排 selector 复用、Phase 3～6、Benchmark、互操作和发布材料。

### Round 1 汇总

- MUST_FIX：FIX-001、FIX-002、FIX-003、FIX-005、FIX-006。
- 已完成代码修复：全部 MUST_FIX 均已完成对应实现和最小回归。
- PARTIAL：当前范围无未完成实现；独立于 Docs 的额外 consumer 和 Windows 权限目录仍是增强性证据缺口。
- BLOCKED：无。
- FAILED：无。
- 回归验证：Unit net8/net6 各 `349/349`；Integration net8/net6 各 `12/12`；Docs package consumer `9/9`；API gate 通过。
- 下一步：由独立 Reviewer 再次验收；本报告不代表 `review.md` 已通过。

### Round 2

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/review.md`
- 版本决策：保持用户指定的 `2.0.0`，未修改 `version.props`。

#### FIX-001

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修复：保留 `AtomicFileCommitter` 的同目录独占 staging、flush、提交前取消检查，补充 Excel 成功新建、成功替换和提交失败测试。
- 验证：net8/net6 Unit 与 net8/net6 Integration 全部通过。

#### FIX-002

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修复：Failure Workbook 每个复制写入块后再次检查取消，补充主异常伴随删除失败、复制后取消和临时残留断言。
- 验证：Failure Workbook 专项测试通过；主异常保留 `Exception.Data` 清理异常，取消路径无临时残留。

#### FIX-005

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修复：CSV seekable/non-seekable 输入字节、header 绑定和后续 `MoveNext()` 资源超限统一返回 `ResourceLimit` 与 `IsTruncated`；补 `MaxErrors`/`InvalidHeader` 测试；新增 net8 独立资源探针覆盖 XLSX ZIP/DOM/shared strings/styles/drawings 与 XLS OLE 样本，并输出子进程 `PeakWorkingSet64`。
- 验证：net8 独立探针测试通过 6 个子进程；net8/net6 Unit 和 net8/net6 Integration 全部通过。探针证据限定为受控样本，不宣称阻止所有压缩放大或 DOM 峰值。

#### FIX-006

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修复：API 成员快照校准为当前源码实际值；生成三个 `2.0.0` nupkg/snupkg；Docs consumer 精确引用 `2.0.0`，隔离恢复资产确认无 `NU1601` 且解析为 `2.0.0`；同步 API diff、package consumer、decisions 和迁移文档。
- 验证：API gate 通过；Docs package-only consumer `9/9`；包及符号包内容检查通过；未执行 NuGet publish。

#### FIX-003 / FIX-004 / FIX-007

- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：用户明确要求本轮只处理 MUST_FIX；未修改 `review.md`。

### Round 2 汇总

- MUST_FIX：FIX-001、FIX-002、FIX-005、FIX-006。
- 已完成：FIX-001、FIX-002、FIX-005、FIX-006。
- PARTIAL：当前范围无未完成实现；独立于 Docs 的额外 consumer 和 Windows 权限目录仍属于增强性证据缺口。
- BLOCKED：无。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 `349/349`；net8/net6 Integration 各 `12/12`；Docs `2.0.0` package-only consumer `9/9`；API gate 通过。
- 下一步：交回 `code-reviewer` 进行再次独立 Review。

### Round 3

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/review.md`
- 版本决策：保持用户指定的 `2.0.0`，未修改版本号。

#### FIX-001

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 根因：原子提交实现直接依赖静态文件 API，无法确定性构造 staging 删除失败及主异常伴随清理失败的故障证据；缺少真实 Windows 目标文件故障覆盖。
- 修复：增加 `IAtomicFileSystem` 内部故障边界，覆盖写入、替换/移动、删除和 flush 路径；保留主异常并通过 `Bing.Offices.{format}.TemporaryCleanupException` 记录清理异常；补充成功提交、提交失败、主异常+清理失败和 Windows 目标锁定场景。
- 修改文件：
	- `src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs`
	- `src/Bing.Offices.Core/AssemblyInfo.cs`
	- `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
	- `tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs`
- 验证：
	- 原子提交专项测试：PASS。
	- net8/net6 Unit：PASS。
	- Windows 目标锁定 Integration：PASS。

#### FIX-002

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 根因：Failure Workbook 的目录创建、临时文件创建和 destination 复制故障缺少独立可重复的故障证据，真实文件路径冲突场景此前被测试夹具自身的未注册校验特性提前拦截。
- 修复：补充目录创建失败、临时文件创建失败和 destination 复制失败的受控文件系统测试；将真实目录冲突 Integration 改为无自定义校验的独立数值转换错误夹具，使其实际进入 Failure Workbook 创建路径，并验证稳定错误分类和清理结果。
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
	- `tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs`
- 验证：
	- Failure Workbook 创建、复制、主异常+清理失败和取消专项测试：PASS。
	- net8/net6 Integration：均为 `14/14` 通过，包含真实临时目录路径冲突。

#### FIX-005

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 根因：资源探针不同 mode 复用了同一小型输入，未能证明独立进程 workload、指标和资源超限语义；异常退出路径也无法向父进程传递失败。
- 修复：为 ZIP、DOM、DOM limit、shared strings、styles、drawings 和 OLE 分别生成受控 workload；探针输出输入字节、Sheet/行/列/单元格、shared strings、styles、pictures、耗时、峰值工作集和错误字段；`dom-limit` 使用实际 `MaxRows` 限制，未处理异常返回非零退出码，资源限制返回结构化状态。
- 修改文件：
	- `tests/Bing.Offices.ResourceProbe/Program.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 验证：
	- net8 独立资源探针 6 个 workload/进程：PASS，指标和退出状态断言通过。
	- net8/net6 Unit：均为 `354/354` 通过。
	- net8/net6 Integration：均为 `14/14` 通过。

#### FIX-003 / FIX-004 / FIX-006 / FIX-007

- 处理要求：`SHOULD_FIX`
- 执行状态：`DEFERRED`
- 原因：用户明确要求本轮只处理 `MUST_FIX`；未修改 `review.md`。

### Round 3 汇总

- MUST_FIX：FIX-001、FIX-002、FIX-005。
- 已完成：FIX-001、FIX-002、FIX-005。
- PARTIAL：无本轮纳入范围内的未完成修复。
- BLOCKED：无。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 `354/354`；net8/net6 Integration 各 `14/14`；`get_errors` 未发现错误；`git diff --check` 待最终收口命令复核。
- 下一步：交回 `code-reviewer` 进行再次独立 Review；本报告不代表 `review.md` 已通过。

### Round 4 Review Fix（当前执行轮次）

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-RELEASE-HARDENING-20260827-001/review.md`
- 版本决策：保持用户指定的 `2.0.0`，未修改版本号。
- 本轮仅处理 `FIX-001`、`FIX-002`；未修改 `review.md`。

#### FIX-001

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 修复：补充 `Move` 和 `Replace` 提交失败同时发生 staging 删除失败的 Excel/CSV 双格式测试；断言主提交异常、格式专属 `TemporaryCleanupException`、删除调用、旧目标未改变及 staging 残留可诊断。
- 验证：
	- 原子提交双故障定向 Unit：Excel/CSV 的 `Move`、`Replace` 组合通过。
	- Windows 目标文件占用 Integration：通过，旧目标内容保持且 staging 清理完成。
	- net8.0 Unit：`358/358` 通过。
	- net6.0 Unit：`358/358` 通过。

#### FIX-002

- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
	- `tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs`
- 修复：补充目录创建、临时文件创建和 destination 复制失败的 destination 未污染、删除调用及临时路径状态断言；新增真实 Windows 锁定目标流场景，验证 Failure Workbook 复制失败分类、原目标内容保持和临时目录清理。
- 验证：
	- Failure Workbook 故障注入定向 Unit：通过。
	- Windows 文件占用及目录冲突 Integration：`3/3` 通过。
	- net8.0 Integration：`15/15` 通过。
	- net6.0 Integration：`15/15` 通过。
	- 编辑器错误检查：相关 C# 文件无错误。

#### 延期项

- `FIX-003`、`FIX-004`、`FIX-006`、`FIX-007`：`SHOULD_FIX`，按用户明确要求的 `fixScope=must` 延期；未修改 `review.md`。

### Round 4 Review Fix 汇总

- MUST_FIX：`FIX-001`、`FIX-002`。
- 已完成：`FIX-001`、`FIX-002`。
- PARTIAL：无本轮纳入范围内的未完成修复。
- BLOCKED：无。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 `358/358`；net8/net6 Integration 各 `15/15`；`get_errors` 无错误；`git diff --check` 通过。
- 工作区安全：未执行 `git add`、`git commit`、`git push`、Tag、NuGet publish 或 PR；版本仍为 `2.0.0`。
- 下一步：交回 `code-reviewer` 进行再次独立 Review；本报告不代表 `review.md` 已通过。
