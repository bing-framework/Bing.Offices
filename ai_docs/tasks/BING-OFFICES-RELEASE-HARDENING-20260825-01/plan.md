<!-- AI_PLAN_STATUS: READY -->
# Bing.Offices 发布前整改实施计划

- **Task ID**：`BING-OFFICES-RELEASE-HARDENING-20260825-01`
- **计划状态**：READY（仅计划，尚未实施）
- **范围**：Bing.Offices 的发布前正确性、API 收敛、内部结构、性能证据、测试、文档和包发布准备。
- **执行入口**：`/execute-plan BING-OFFICES-RELEASE-HARDENING-20260825-01`、`/run-plan` 或 `$execute-plan`。
- **非目标**：本任务计划阶段不修改任何业务代码、测试、配置、数据库或现有任务文档；后续实施也不得自动 `git add`、commit、push、tag 或发布 NuGet 包。

## 1. 输入、约束和证据

### 1.1 已读取的输入

1. 当前用户的“Bing.Offices 发布前整改计划与持续执行提示词”。
2. 根目录 [AGENTS.md](../../../AGENTS.md) 的 UTF-8、安全、非破坏性 Git 和 SQL 专项约束（本任务不涉及 `Bing.Data.Sql`）。
3. [create-plan.prompt.md](../../../.github/prompts/create-plan.prompt.md) 的计划格式和仅写入 `plan.md` 约束。
4. [README.md](../../../README.md)、[docs/excel/README.md](../../docs/excel/README.md)、[docs/excel/import-validation.md](../../docs/excel/import-validation.md)、[docs/excel/nuget-migration.md](../../docs/excel/nuget-migration.md)。
5. 前序稳定化任务的 [review.md](../TASK-BING-OFFICES-STABILIZE-20260825-001/review.md) 与 `execution.md` 作为已完成整改的历史证据。
6. 当前 `src/`、`tests/`、`benchmarks/` 中与本任务有关的实现、测试和项目配置。

### 1.2 输入缺口和事实冲突

- 用户指定的 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260825-220502.md` 不存在，无法读取，不能引用或推断其结论。实施开始前应再次确认是否由用户补充；若仍缺失，以当前源码、测试与前序稳定化审查为准。
- [docs/excel/import-validation.md](../../docs/excel/import-validation.md) 声称 `ValidateMode.Continue` 会收集两个来源的错误；当前 [NpoiExcelImporter.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs) 在 Workbook 原生校验失败后立即跳过该行物化，因此不会继续执行该行配置校验。这是文档与实际契约不一致，必须修正文档或在 P0 设计评审中明确调整行为，不能保留模糊表述。
- 当前模式未提供可运行终端的命令执行能力，只能查询最近终端状态（无活动终端）；因此本计划不能宣称已完成 `git status --short`、`git diff --stat`、`git diff --name-only` 或 `git diff --check`。执行 Agent 的第一个步骤必须运行这些只读检查，记录与本任务重叠的未提交变更，并且绝不回退或清理陌生改动。

### 1.3 技术基线

- 解决方案：[Bing.Offices.sln](../../../Bing.Offices.sln)。核心链路是 `Bing.Offices.Abstractions <- Bing.Offices.Core <- Bing.Offices.Npoi`。
- NPOI 实现为 DOM 模型；当前文档已明确不承诺 streaming 或零 GC。整改目标是控制可避免的峰值复制和热路径分配，不承诺将 NPOI DOM 变为真正流式 XLSX 引擎。
- 单元测试使用 xUnit；集成测试项目为 `tests/Bing.Offices.Tests.Integration`；性能项目为 `benchmarks/Bing.Offices.Benchmarks`（BenchmarkDotNet）。
- 已知可运行运行时为 .NET 6/.NET 8；历史验证显示 net5/net7/netcoreapp3.1 测试宿主可能缺少运行时。多目标项目必须继续编译；无法执行的目标必须显式记录环境阻塞，不能伪报通过。

## 2. 当前完成度评估

### 2.1 已完成并有源码/测试证据的基础

| 领域 | 状态 | 当前事实 |
|---|---|---|
| Workbook Request 主路径与 provider-neutral 边界 | 已完成基础 | Public API 基线及“不泄漏 NPOI 类型”测试已存在；NPOI 类型主要保持 internal。 |
| 映射方向和优先级 | 已完成基础 | 当前设计为 `Attribute < Profile < JSON/XML Document < Request Fluent`，最终合并在 Mapping Plan 编译阶段。 |
| Workbook 原生校验接入 | 部分完成 | 已支持 LIST、TEXT_LENGTH、INTEGER/DECIMAL、DATE/TIME、直接范围列表和不支持特性策略。 |
| Workbook 失败后跳过物化 | 已完成 | 前序 FIX-005 让原生规则失败的行不再进入转换/配置校验/实体物化。 |
| 失败工作簿 | 部分完成 | 已支持 AnnotatedOriginal、ErrorRowsOnly、错误摘要、部分样式/合并/数据验证/图片复制。 |
| 请求级样式和 XSSF 自定义填充色 | 部分完成 | `NpoiStyleCache` 有按 Workbook 缓存；测试覆盖同样式复用和 XSSF 前景色。 |
| 模板批注冲突策略 | 部分完成 | 导出自定义表头支持 Preserve/Append/Replace/Fail；失败批注没有对应冲突策略。 |
| 基础资源限制 | 部分完成 | 已有输入、行、错误、图片、唯一值限制及测试。 |
| Benchmark/资源探针 | 部分完成 | 有 BenchmarkDotNet 和资源探针，但对照基线和保留容量方法不能证明用户要求的结论。 |

### 2.2 未完成或需要重构的关键项

| 优先级 | 问题 | 当前证据与结论 |
|---|---|---|
| P0 | `LESS_OR_EQUAL` 原生校验错误 | [NpoiWorkbookValidationPipeline.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiWorkbookValidationPipeline.cs) 的操作符 `7` 错误依赖 `Formula2`，应只用 `Formula1`。 |
| P0 | `EmptyCellAllowed` 未实现 | 原生 Data Validation 流程不读取 `IDataValidation.EmptyCellAllowed`，空单元格会被 LIST/数值等规则误判。 |
| P0 | 日期精确格式可被绕过 | [ExcelValidationRules.cs](../../../src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs) 在检查 `ExcelDateAttribute.Format` 前，若 `ConvertedValue` 或单元格值是 `DateTime`/`DateTimeOffset` 即返回 true。 |
| P0 | 失败工作簿多份峰值复制 | [NpoiFailureWorkbookWriter.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs) 先写 `MemoryStream`，再 `ToArray()`，之后才检查 `MaxBytes`。 |
| P0 | 样式显式清除语义缺失 | `ExcelCellStyle` 的枚举默认值与“未指定”混用；Compose 不能可靠表达“覆盖为默认/清除模板属性”。 |
| P0 | XLSX/HSSF 颜色和边框语义不完整 | XSSF 字体/填充支持 ARGB；边框仍走 HSSF indexed color，背景色与前景色折叠，HSSF 自定义颜色只支持少量颜色。 |
| P0 | 失败输出保真与批注冲突契约不完整 | ErrorRowsOnly 没有明确复制超链接、富文本、行级样式、列宽、批注等；AnnotatedOriginal 直接追加批注文本，没有公开冲突策略。 |
| P1 | 公开 API 仍含冗余/歧义 | `ExcelMapping.For<T>()`、双重 `Mapping(...)`、`HeaderMatch`、`MaxColumnCount`、`EnabledEmptyLine`、`IgnoreEmptyLineAfterData`、`AddNavigationSheet`、`ExcelSetting.Default`、`ICellValueConverter` 均需治理。 |
| P1 | 热路径分配与复杂度 | Regex 无界静态缓存；ValidationRangeIndex 每次查询分配 `List`/`HashSet`；UniqueTracker 每次 reserve 遍历 pending；导出逐行构造 dynamic key 集合/空字典；空行判断扫描所有图片键。 |
| P1 | 文档/API 示例不完整 | 文档索引引用不存在的 `ai_docs/excel/00-overview.md` 至 `08-validation.md`，迁移文档仍是 `MIGRATION_CURRENT_MAJOR` 占位。 |

**整体判断**：不能以单一百分比描述完成度。基础架构、前序 P0 回归修复、主路径测试已具备，但用户本次定义的发布门槛包含多项尚未完成或未证实的 P0 行为、性能和互操作性证据。因此当前状态为**可继续开发，尚不具备发布结论**。

## 3. 总体实施顺序和质量门

1. `RH-00` 建立工作区基线与可追溯矩阵。
2. Phase 1 完成 P0 正确性闭环。所有后续 API/性能工作必须建立在明确行为契约上。
3. Phase 2 先设计并批准 major-breaking API 表，再实施收敛和迁移路径。
4. Phase 3 按内部边界拆分，保持公共程序集边界和行为不变。
5. Phase 4 以可测量的热点为单位优化；不以“理论优化”替代基准数据。
6. Phase 5 同步文档、样例和 API 契约测试。
7. Phase 6 仅在全部门禁、包消费者和人工互操作性证据完备后给出 Release/No-release 建议。发布动作必须取得用户明确批准。

每个任务完成时，执行 Agent 在 `progress.md` 记录状态与命令摘要，在 `decisions.md` 记录非显然取舍，在 `test-results.md` 记录原始产物路径和通过/阻塞结论；这些文件只能由后续执行阶段创建。本计划阶段不创建它们。

## 4. Phase 0：基线与契约冻结

### RH-00 工作区、依赖和公共面基线

- **优先级**：P0
- **依赖**：无
- **目标**：在任何改动前建立可复现的 dirty-worktree、API、测试、包和性能证据基线。
- **已确认文件**：[Bing.Offices.sln](../../../Bing.Offices.sln)、[PublicApiContractTest.cs](../../../tests/Bing.Offices.Tests/PublicApiContractTest.cs)、[framework.props](../../../framework.props)、各 `.csproj`。
- **实施步骤**：
  1. 运行 `git status --short`、`git diff --stat`、`git diff --name-only`、`git diff --check`；将与本任务相关项写入执行记录，不处理无关变更。
  2. 检查目标计划、历史任务、代码库分析目录；再次确认缺失外部报告。
  3. 建立“生产符号 -> 测试方法”可追溯矩阵，至少覆盖本计划 P0 变更和所有 major-breaking API。
  4. 记录 SDK、已安装运行时、NPOI/CsvHelper/BenchmarkDotNet 版本及可执行 TFM。
  5. 对当前 Release 构建、net6/net8 Unit/Integration、Docs consumer、pack 和基准做基线运行；失败必须保留原始日志和环境原因。
- **测试**：本任务是验证基线，不改代码；所有命令见第 10 节。
- **风险**：工作区已有未提交改动可能影响测量；必须隔离记录，而非回退。
- **验收**：有可审计的环境、Git、API、测试和性能基线；没有误将环境阻塞标为代码失败或通过。

## 5. Phase 1：P0 正确性与数据安全

### RH-101 Workbook Data Validation 语义修复

- **优先级**：P0
- **依赖**：RH-00
- **目标**：按 Excel/NPOI 原生约束正确处理比较符、空单元格和 `ValidateMode`，且不破坏前序“Workbook 失败不物化”语义。
- **已确认文件**：[NpoiWorkbookValidationPipeline.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiWorkbookValidationPipeline.cs)、[NpoiExcelImporter.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs)、[ExcelP0RegressionTest.cs](../../../tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs)。
- **候选文件**：导入错误/策略契约、NPOI Data Validation 适配辅助类型。
- **实施步骤**：
  1. 将操作符编码抽取为语义明确的比较逻辑或经验证的 NPOI 常量映射，消除魔法数依赖。
  2. 修正 `LESS_OR_EQUAL`：仅要求解析 `Formula1`，比较 `value <= Formula1`，不得要求 `Formula2`。
  3. 在所有受支持 validation 类型的规则计算前处理 `IDataValidation.EmptyCellAllowed`：空单元格允许时通过；不允许时依当前规则类型给出稳定的 WorkbookValidation 错误。
  4. 明确空白归一化与“空单元格/空字符串/公式结果为空”的定义；保证 LIST、数值、日期、文本长度采用同一契约。
  5. 保持 `ExcelUnsupportedFeaturePolicy.Report` 仅记录而不使行无效、`Fail` 使行无效的现有设计；更新文档中的 Continue 行为为真实实现。
- **用例矩阵**：

| Given | When | Then |
|---|---|---|
| INTEGER `LESS_OR_EQUAL`，仅 Formula1=10 | 值=10/11 | 10 通过，11 报 WorkbookValidation；Formula2 缺失不影响合法值。 |
| DATE/TIME `LESS_OR_EQUAL`，仅 Formula1 | 边界内/外日期时间 | 只按 Formula1 判定，保留 Excel serial 与文本日期路径。 |
| LIST、INTEGER、DATE 的 `EmptyCellAllowed=true` | 目标为空 | 不产生 WorkbookValidation，行可继续进入后续合法流程。 |
| 同规则 `EmptyCellAllowed=false` | 目标为空 | 报稳定的 WorkbookValidation 错误并按 ValidateMode 处理。 |
| Workbook 失败且 `ConfiguredAndWorkbook`、Continue | 同一行另有转换/配置规则 | 不物化、不额外运行该行配置校验，Unique pending 回滚。 |
| 不支持公式 | Report / Fail | Report 保留行并记录错误；Fail 不物化。 |

- **Mock 边界**：使用内存 XSSFWorkbook/HSSFWorkbook 创建真实 NPOI 规则；不 mock NPOI validation 对象。
- **风险**：NPOI 不同格式的空值表征不同；必须为 XLSX 与 XLS 覆盖共同契约，无法统一的行为须明确文档限制。
- **验收**：新增直接单元测试覆盖所有比较操作符至少一个、`LESS_OR_EQUAL`、空允许/不允许、继续/停止和 XLSX/XLS；现有 validation 回归全部通过。

### RH-102 ExcelDate 精确格式和转换顺序修复

- **优先级**：P0
- **依赖**：RH-00
- **目标**：当 `ExcelDateAttribute.Format` 指定时，验证原始输入文本必须符合精确格式，不得因先转换为日期而绕过；未指定 Format 时保留 typed date、DateTimeOffset 和 Excel serial 合法性。
- **已确认文件**：[ExcelValidationRules.cs](../../../src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs)、[NpoiImportRowMaterializer.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiImportRowMaterializer.cs)、[CsvTest.cs](../../../tests/Bing.Offices.Tests/CsvTest.cs)、[ExcelWorkbookRequestTest.cs](../../../tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs)。
- **候选文件**：`ExcelValidationContext`、属性定义、CSV pipeline 测试。
- **实施步骤**：
  1. 先定义 raw text、provider typed value 和 converted CLR value 的验证职责。
  2. `Format` 非空时始终用原始规范化文本执行 `TryParseExact`，应用 `CultureName` 优先于请求文化的规则；不允许转换后的值短路。
  3. `Format` 为空时明确 DateTime、DateTimeOffset、数值 serial 和文本解析的接受矩阵；处理 nullable 目标类型和空值与 Required 的分工。
  4. 不把 `DateTime.TryParse` 的宽松行为当作精确格式的替代；确保 CSV 和 XLSX 复用相同规则。
- **用例矩阵**：

| Given | When | Then |
|---|---|---|
| `Format="yyyy-MM-dd"` 且输入 `2026/08/25` 可被转换器读为 DateTime | 校验 | 失败，不能被 ConvertedValue 短路。 |
| 同格式且输入 `2026-08-25` | 校验 | 通过。 |
| 指定 `CultureName` 的本地化格式 | 校验 | 使用属性文化，非请求默认文化。 |
| 无 Format 的 DateTime/DateTimeOffset/Excel serial | 校验 | 通过。 |
| 无效 serial、无效文本、空 nullable 值 | 校验 | 分别失败或按空值策略通过，错误码稳定。 |
| CSV 与 XLSX 相同输入 | 校验 | 一致结果。 |

- **风险**：不能在修复中改变 raw validation 与 converted validation 的既有顺序，除非由 API 变更任务明确批准。
- **验收**：CSV/XLSX 都有直接正例、反例、边界例；精确格式绕过复现测试先失败后通过。

### RH-103 失败工作簿受限写出、保真合同与批注冲突策略

- **优先级**：P0
- **依赖**：RH-00
- **目标**：避免不必要的整块复制，在写入期间强制 `MaxBytes`，定义并测试 ErrorRowsOnly/AnnotatedOriginal 的保真与批注冲突契约。
- **已确认文件**：[NpoiFailureWorkbookWriter.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs)、[ExcelImportPolicies.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImportPolicies.cs)、[ExcelP0RegressionTest.cs](../../../tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs)。
- **候选文件**：流计数包装器、失败输出选项/API、NPOI picture/hyperlink/comment 辅助扩展、导入文档。
- **设计决策（实施前必须记录）**：
  - 对可写 destination 采用计数限制流或等价机制，让 NPOI 直接写入受限目标，超限尽早失败；不得再使用 `MemoryStream.ToArray()` 作为正常输出路径。
  - 当目标流不可回滚时，超限后可能已有部分字节是 API 可观察事实。默认策略应优先通过 staging 文件/可回滚流或明确新增原子性选项；不能声称所有 Stream 都可原子写入。
  - ErrorRowsOnly 必须明确“支持复制”和“明确不支持复制”的对象清单。至少评估单元格值/公式、样式、行高、列宽、合并区域、数据验证、图片、超链接、批注、富文本、行/列隐藏、冻结窗格、条件格式和命名区域。未可靠支持的项目必须在文档/错误策略中披露，不能静默宣称完整保真。
  - 为导入失败注释引入显式 conflict policy，复用导出 `ExcelCommentConflictPolicy` 或在 Imports 定义方向专用且不歧义的策略；默认行为必须兼容当前“追加已有文本”的实际效果，若要改变则列为 major-breaking change。
- **用例矩阵**：

| Given | When | Then |
|---|---|---|
| 输出小于 MaxBytes | 写失败工作簿 | 目标是可打开的工作簿，长度不超过上限。 |
| 输出超过 MaxBytes | 写失败工作簿 | 期间抛出稳定异常；不创建 `ToArray` 完整副本；按公开原子性契约验证目标状态。 |
| AnnotatedOriginal 有既有批注 | Preserve/Append/Replace/Fail | 分别保留、追加、替换、失败，作者/可见性规则明确。 |
| ErrorRowsOnly 含公式、样式、合并、验证、图片 | 输出 | 已承诺部件完整可读、行映射正确。 |
| ErrorRowsOnly 含未支持部件 | 输出 | 按已声明策略保留、显式拒绝或记录；不静默损坏。 |
| 多个 Sheet、非零表头行、取消令牌 | 输出 | Sheet/表头映射正确，取消及时停止。 |

- **风险**：NPOI `IWorkbook.Write` 对 Stream 的异常/关闭行为需要真实回归验证；不能以 mock 代替。
- **验收**：新增受限流测试及大于阈值的失败测试；不再存在“先 `ToArray()` 再 MaxBytes 检查”的生产路径；能力边界写入文档。

### RH-104 样式、颜色与显式重置模型

- **优先级**：P0
- **依赖**：RH-00
- **目标**：区分样式属性“未指定”“设置值”“显式清除/恢复默认”，统一 Header/Body/固定列/动态列/模板叠加，正确处理 XSSF ARGB 和 HSSF 能力边界。
- **已确认文件**：[ExcelCellStyle.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs)、[NpoiStyleCache.cs](../../../src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs)、[NpoiExportSheetWriter.cs](../../../src/Bing.Offices.Npoi/Exports/NpoiExportSheetWriter.cs)、[ColorResolver.cs](../../../src/Bing.Offices.Npoi/Resolvers/ColorResolver.cs)、[ExcelSheetExportBuilder.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelSheetExportBuilder.cs)。
- **候选文件**：Mapping style configuration、ExcelColumnPlan、JSON/XML serialization、样式单元测试与文档。
- **实施步骤**：
  1. 先审计所有 style 的公共入口（Sheet/Header/Body、column HeaderStyle/BodyStyle、dynamic definitions、mapping document/profile、模板覆盖），形成能力矩阵。
  2. 设计 provider-neutral 三态表达。优先采用可序列化、向后兼容的 patch/reset 字段或专用 value object；不能以 enum 默认值猜测用户是否指定。
  3. 明确 FillPattern、ForegroundColor、BackgroundColor 的组合规则；Background 不得再悄然折叠为 Foreground。
  4. 为 XSSF 字体、填充、各边框颜色使用 XSSF color API；为 HSSF 仅允许可表示 indexed palette 颜色或提供明确可配置的失败/降级策略。
  5. 将 `NpoiStyleCache` key 扩展为全部语义字段（包括 reset/颜色通道/边框颜色），保持按 Workbook 隔离、线程安全和不超量创建。
  6. 缓存 HeaderAttribute 生成的样式/字体，避免每个表头单元格 `CreateCellStyle/CreateFont`；不得破坏请求样式优先级。
- **用例矩阵**：

| Given | When | Then |
|---|---|---|
| 模板有数字格式/填充/边框 | 仅叠加 Bold | 未指定属性保留。 |
| 模板有样式 | 请求显式 reset 属性 | 对应属性恢复约定默认，其他属性保留。 |
| Header、Body、固定列、动态列均配置不同样式 | 导出 | 每个区域按既定优先级生效。 |
| XSSF 指定 ARGB 字体/填充/背景/四边颜色 | 重开文件 | 颜色通道和值正确。 |
| HSSF 支持/不支持颜色 | 导出 | 支持颜色正确；不支持按公开契约拒绝或降级。 |
| 重复样式定义 | 大量单元格 | style/font 数量受缓存控制，不按单元格线性增加。 |

- **风险**：改变样式 DTO 的序列化形状或默认行为可能是 breaking change，必须交给 RH-201 决定版本策略。
- **验收**：XLSX/XLS 正反例、模板 compose/reset、固定/动态列、样式缓存全部有直接测试；颜色处理不再混用 XSSF 与 HSSF API。

### RH-105 模板覆盖、批注与自定义表头补强

- **优先级**：P0
- **依赖**：RH-104
- **目标**：明确模板单元格覆盖范围，确保多级自定义表头、注释、合并和样式与请求级样式组合可预测。
- **已确认文件**：[ExcelExportPolicies.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelExportPolicies.cs)、[NpoiExportSheetWriter.cs](../../../src/Bing.Offices.Npoi/Exports/NpoiExportSheetWriter.cs)、[ExcelWorkbookRequestTest.cs](../../../tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs)。
- **实施步骤**：
  1. 定义 `PreserveTemplate` 与 `ReplaceTemplate` 对值、公式、样式、批注、超链接、rich text 的确切含义；当前代码无论策略都会由 `SetCellValue` 覆盖公式，文档必须一致。
  2. 对已有 merged regions 与新 HeaderRows 合并冲突建立确定性策略并增加输入验证。
  3. 确保 custom header cell style 的入口明确且能与 HeaderStyle/column HeaderStyle 合成；若当前模型没有 per-cell style，不在 P0 擅自新增公开 API，先记录为 RH-201 API 设计项。
- **验收**：模板保留/替换、公式覆盖、批注四策略、多级表头合并冲突和样式优先级都有回归测试与文档。

## 6. Phase 2：API 收敛与 major-version 迁移

### RH-201 公开 API 审计、破坏性变更表与兼容策略

- **优先级**：P1（先于所有 public API 删除）
- **依赖**：RH-101 至 RH-105 的契约决定
- **目标**：形成可批准的 next-major API 表，不在未批准前直接删除 public 成员。
- **已确认文件**：[ExcelMapping.cs](../../../src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs)、[ExcelImport.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs)、[ExcelExport.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelExport.cs)、[ExcelSetting.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Settings/ExcelSetting.cs)、[ICellValueConverter.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Conversions/ICellValueConverter.cs)、[PublicApiContractTest.cs](../../../tests/Bing.Offices.Tests/PublicApiContractTest.cs)。
- **实施步骤**：
  1. 用源码、测试、Docs consumer、README 和 repository-wide search 列出全部公开 API、使用点、程序集归属、可替代 API、二进制/源码兼容影响。
  2. 为每项生成下表格式的批准清单；没有替代路径和迁移示例的 API 不得删除。

| 候选项 | 当前问题 | 建议 | 迁移目标 | 版本策略 |
|---|---|---|---|---|
| `ExcelMapping.For<T>()` | 不区分方向 | 新增/迁移为 `ExcelMapping.Import<T>()`、`ExcelMapping.Export<T>()`，方向 builder 分离 | import/export 专用构建器 | next-major 移除或当前 major Obsolete 转发，取决于包策略 |
| `Mapping(configuration)` / `Mapping(document)` | 过载按语义而非名称区分 | `MappingConfiguration(...)`、`MappingDocument(...)` 或方向化名称 | 显式方法 | next-major |
| `HeaderMatch` | 布尔含义不自说明 | `RequireExpectedHeaders` | 新名称 | next-major，保留 obsolete shim 的可行性待确认 |
| `MaxColumnCount` | 实际字段是读取安全上限 | `MaxReadColumns` | 新名称 | next-major |
| `EnabledEmptyLine` | 语义不清 | `ReportEmptyRows` | 新名称 | next-major |
| `IgnoreEmptyLineAfterData` | 语义不清 | `StopAtFirstEmptyRow` | 新名称 | next-major |
| `AddNavigationSheet` | 仅 `SelectMany + AddSheet` 的便捷包装 | 移除或 Obsolete | `AddSheet(name, parents.SelectMany(...))` | next-major |
| `ExcelSetting.Default` | 进程级可变全局状态 | DI/request metadata/options | 显式 request 或 DI 配置 | next-major |
| `ICellValueConverter` | provider object 文本提取兼容层 | `IExcelValueConverter` | 双向 provider-neutral converter | 当前 major obsolete，next-major 删除 |

  3. 审查扩展方法、类型可见性和返回类型，避免公开面暴露 NPOI、可变内部集合或仅测试用途类型。
  4. 更新 PublicApiContractTest，使 approved baseline 与版本策略一致；为删除项建立负向断言，为保留 obsolete shim 建立编译级 consumer 测试。
- **风险**：改变 `ExcelMapping` 所在 Core 程序集或移动类型会造成 package/reference 迁移成本；必须在 NuGet consumer 验证中覆盖。
- **验收**：用户/维护者批准 breaking table；每项都有迁移示例、弃用/删除版本、测试和文档位置；没有“隐式删除”。

### RH-202 实施已批准 API 收敛与配置去全局化

- **优先级**：P1
- **依赖**：RH-201
- **目标**：落实批准的 API 表，清除不必要 public surface，同时保持当前 major 或 next-major 的明确兼容界限。
- **已确认文件**：RH-201 列出的 API 文件、DI 扩展、导入/导出 builders、package consumer tests。
- **实施步骤**：
  1. 引入方向明确的 Mapping 入口和 builder，确保 Import 与 Export 配置不再混用。
  2. 使用具名 Mapping 方法替换语义重载，并用 obsolete forwarding 或文档迁移完成过渡（仅当 RH-201 允许）。
  3. 重命名布尔方法，处理参数默认值和链式调用的 source compatibility。
  4. 将 `ExcelSetting.Default` 读取点迁移到注入/请求元数据；设计线程安全、范围明确的默认值来源。
  5. 消除 `AddNavigationSheet` 和 legacy `ICellValueConverter` 的生产主路径；不能删除前必须给出兼容包装和最终移除版本。
  6. 每次 public surface 更改同步 XML 文档、迁移文档、PublicApiContractTest 和 Docs consumer。
- **测试**：新 API 正例；旧 API（若存在兼容窗）编译/运行且产生弃用信息；移除后 negative API baseline；并发请求不串扰 metadata。
- **验收**：`PublicApiContractTest`、Docs package consumer、pack 均通过；文档没有继续推荐废弃入口。

## 7. Phase 3：内部架构与可维护性

### RH-301 导入/导出职责拆分和协作对象边界

- **优先级**：P1
- **依赖**：RH-101 至 RH-105
- **目标**：保持 public API 不变，降低 `NpoiExcelImporter`、`NpoiExcelExporter`、请求 builder 和大型测试文件的认知负担。
- **当前事实**：前序任务已提取 `NpoiImportRowMaterializer`、`NpoiWorkbookValidationPipeline`、`NpoiFailureWorkbookWriter`、`NpoiExportSheetWriter`、计划构建器等；本轮不得重复拆分或仅为文件数拆分。
- **候选职责**：原生 validation formula/value parser、failure artifact feature copier、style normalization/cache、template conflict resolver、dynamic column schema、import row emptiness/image row index。
- **实施步骤**：
  1. 以依赖方向和单一变化原因审计已有协作对象，识别仍有多个独立职责的类。
  2. 只提取能被直接单测、没有 public visibility 必要的 internal collaborator；通过构造参数或纯函数传递显式依赖，避免静态隐式全局状态。
  3. 将巨大测试按领域拆分（validation、failure output、styles/templates、mapping/API、resource/performance），保持 test method 名称和历史断言可追溯。
  4. 避免生产 `InternalsVisibleTo` 扩张；优先黑盒测试，确有算法必要时仅允许测试 friend，并记录原因。
- **验收**：依赖方向仍为 Abstractions -> Core -> Npoi；没有新 public provider 类型；每个新协作对象有直接职责测试；没有无行为变化的大规模格式化。

### RH-302 Mapping Plan、配置克隆与缓存语义审计

- **优先级**：P1
- **依赖**：RH-201
- **目标**：确保新 API、style reset 和 document/profile patch 语义不会在克隆、合并或缓存键中丢失。
- **已确认文件**：[ExcelMapping.cs](../../../src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs)、`MappingConfigurationCloner`、`MappingConfigurationMerger`、`ExcelMappingPlanFactory`、`ExcelTypeMapFactory`、mapping/profile tests。
- **实施步骤**：
  1. 枚举 `ExcelColumnConfiguration` 的全部字段，逐一对照 fluent builder、document JSON/XML、cloner、merger、plan factory、cache key 与 public snapshot。
  2. 对 style/reset、validation tombstone、dynamic column、layout、converter、value mapping 等字段建立“输入 -> merge -> plan -> export/import”的直接测试。
  3. 评估 `Build()` 的深拷贝完整性，禁止 fluent builder 后续变更污染已构建 request。
  4. 对 mapping cache 命中/未命中、tenant/config version 隔离和并发构建设置行为与资源门槛。
- **验收**：没有新增字段遗漏克隆/合并/缓存键；优先级和 tombstone 语义有端到端测试。

## 8. Phase 4：性能与资源治理

### RH-401 有界 Regex 缓存与输入防护

- **优先级**：P1
- **依赖**：RH-102
- **目标**：让 Regex validation 既保留 timeout，也不接受无限不同 pattern 导致的进程级缓存增长。
- **已确认文件**：[ExcelValidationRules.cs](../../../src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs)。
- **实施步骤**：选择有容量上限、明确淘汰/不缓存策略且跨目标框架可用的实现；缓存 key 必须包括 pattern、options、timeout（如果可配置）；保留 Regex timeout 和异常处理契约。
- **测试**：正常命中、达到容量、超容量后行为、恶意大量 pattern、超时 pattern；并发调用不抛异常且不超上限。
- **验收**：不再使用无界 `ConcurrentDictionary<string, Regex>`；容量/淘汰行为可测且文档化。

### RH-402 导入热路径分配与查询复杂度优化

- **优先级**：P1
- **依赖**：RH-101、RH-301
- **目标**：消除已确认的重复扫描/短生命周期集合分配，不改变导入结果和错误顺序。
- **已确认文件**：[ValidationRangeIndex.cs](../../../src/Bing.Offices.Npoi/Imports/ValidationRangeIndex.cs)、[UniqueTracker.cs](../../../src/Bing.Offices.Abstractions/Bing/Offices/Providers/UniqueTracker.cs)、[NpoiImportRowMaterializer.cs](../../../src/Bing.Offices.Npoi/Imports/NpoiImportRowMaterializer.cs)。
- **实施步骤**：
  1. 将 ValidationRangeIndex 的 query result 去重和收集改为低分配策略，例如预计算唯一 validation 条目/调用方可复用 buffer/单规则快路径；不得共享可变缓冲导致并发错误。
  2. 为图片索引增加按行存在的 `HashSet<int>` 或等价索引，使空行判断不再 `Any(imageIndex.Keys)` 扫描所有图片。
  3. 用 `_pendingValueCount` 替代 `UniqueTracker.PendingCount()` 全表扫描；Begin/Commit/Rollback 必须精确重置，限额边界仍正确。
  4. 仅在 profiling/benchmark 证实后优化其他短生命周期字典；保留 dynamic values 的惰性分配。
- **测试**：ValidationRangeIndex 大矩形、不相交/重叠、边界行列、重复 validation、多线程并发读；Unique 精确上限/回滚/多 key；图片/空行正反例。
- **验收**：功能测试全绿，BenchmarkDotNet 或 allocation test 显示目标热点分配降低，并无并发共享状态。

### RH-403 导出热路径与样式缓存优化

- **优先级**：P1
- **依赖**：RH-104
- **目标**：避免每行构造动态列 key 集合、无意义空 dynamic 字典以及 HeaderAttribute 重复 style/font 创建。
- **已确认文件**：[NpoiExportSheetWriter.cs](../../../src/Bing.Offices.Npoi/Exports/NpoiExportSheetWriter.cs)、[NpoiStyleCache.cs](../../../src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs)。
- **实施步骤**：
  1. 在 Sheet write 前一次性预计算动态 schema/合法 key 集合；每行仅检查值字典。
  2. `WriteCell` 在无 dynamic values 且 getter 非字典时直接得到 null 值，不创建空 `Dictionary<string,object>`。
  3. 把 HeaderAttribute style/font 合并到 Workbook style cache，确保与 request/template style compose 的先后关系受测试保护。
  4. 不将数据完整物化为集合以换取微优化；继续支持一次性 `IEnumerable<T>`。
- **验收**：动态未知 key 行为不变；多行导出的分配/样式数下降；输出重开验证样式、值、批注和宽度仍正确。

### RH-404 Benchmark、资源探针和基线治理

- **优先级**：P1
- **依赖**：RH-401 至 RH-403
- **目标**：提供可信的性能结论，不使用人为增加中间 buffer 的“legacy”方法冒充历史产品基线。
- **已确认文件**：[StreamPipelineBenchmarks.cs](../../../benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs)、[MappingValidationBenchmarks.cs](../../../benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs)、[Program.cs](../../../benchmarks/Bing.Offices.Benchmarks/Program.cs)。
- **实施步骤**：
  1. 移除或重命名 `ImportWithLegacySourceBuffer`/`ExportWithLegacyDestinationBuffer`：它们是当前实现外包一层 MemoryStream，不是可审计历史实现。
  2. 若能获得基线 NuGet/历史 tag/独立 legacy assembly，使用真实调用链作为 baseline，记录包版本、commit、数据集和限制；若不可获得，明确标记“无历史可比基线”，只报告当前版本绝对指标及改动前后同代码库 baseline。
  3. 每次迭代创建新的 destination stream 或统计 Capacity/retained capacity；禁止仅 `SetLength(0)` 而隐藏复用容量对峰值的影响。
  4. 为 stream import/export、failure workbook 上限、validation index、unique tracker、regex cache、style cache 建立分组基准；固定数据规模、预热、迭代、GC/CPU/SDK 记录。
  5. 资源探针不得人为分配与生产路径无关的 LOH payload 来替代真实场景；拆分为真实 mapping/unique workload 与独立 GC 健康检查，并清楚标明。
- **验收**：每个结论可链接到 JSON/Markdown 原始产物；包含均值、P95（必要时）、allocated bytes、GC、峰值/保留容量、输入规模、环境和对照来源；无伪造 legacy 对比。

## 9. Phase 5：测试、文档和开发体验

### RH-501 测试强化与测试资产治理

- **优先级**：P1
- **依赖**：Phase 1-4 的各实现任务
- **目标**：测试关键行为而非内部调用次数，并覆盖 API、异常、边界、并发、格式与消费者场景。
- **已确认文件**：[ExcelP0RegressionTest.cs](../../../tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs)、[ExcelWorkbookRequestTest.cs](../../../tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs)、[StreamPipelineTest.cs](../../../tests/Bing.Offices.Tests/StreamPipelineTest.cs)、[CsvTest.cs](../../../tests/Bing.Offices.Tests/CsvTest.cs)、[PublicApiContractTest.cs](../../../tests/Bing.Offices.Tests/PublicApiContractTest.cs)。
- **实施步骤**：
  1. 将 Phase 1-4 的用例矩阵落为独立、命名清楚的 xUnit 测试；每个测试保留中文 XML 测试目的和 AAA。
  2. 采用真实 in-memory NPOI workbook、真实 stream、真实 DI 组合验证 provider 行为；仅 mock 时间、IO 边界或不可控外部依赖。
  3. 增加 public API baseline 与 Docs consumer test，保证包引用场景而非项目引用场景可用。
  4. 增加随机/属性测试仅用于稳定规则（日期格式、比较操作、范围索引），种子固定并保存失败最小样本。
  5. 明确 XLS/XLSX、CSV、net6/net8 覆盖分工；Office/LibreOffice/WPS 不属于自动单测，交由 RH-602 人工互操作矩阵。
- **验收**：每个最终生产符号可追溯到测试；新增 bug 都有先失败后通过的回归；测试不依赖公网、sleep 或生产数据库。

### RH-502 文档、迁移和示例同步

- **优先级**：P1
- **依赖**：RH-201、RH-202，以及所有会改变用户契约的任务
- **目标**：将实际行为、限制和迁移路径写成可执行文档，消除占位符和失效链接。
- **已确认文件**：[README.md](../../../README.md)、[docs/excel/README.md](../../docs/excel/README.md)、[docs/excel/import-validation.md](../../docs/excel/import-validation.md)、[docs/excel/mapping-profile.md](../../docs/excel/mapping-profile.md)、[docs/excel/mapping-json-xml.md](../../docs/excel/mapping-json-xml.md)、[docs/excel/dynamic-columns.md](../../docs/excel/dynamic-columns.md)、[docs/excel/nuget-migration.md](../../docs/excel/nuget-migration.md)。
- **实施步骤**：
  1. 修复现有 README 指向不存在 `ai_docs/excel/00-overview.md` 至 `08-validation.md` 的链接，改为现有文档或实际新增文档。
  2. 更新 import validation：Workbook-first 失败短路、EmptyCellAllowed、支持/不支持的 native validation、日期精确格式、错误收集与 Unique 回滚。
  3. 更新 template/style 文档：Preserve/Replace 覆盖矩阵、批注冲突、样式优先级、显式 reset、XLSX/HSSF 颜色限制。
  4. 更新 stream/resource 文档：调用方 stream ownership、DOM 限制、MaxBytes 原子性/部分写入契约、资源上限。
  5. 以真实 major 版本和 approved breaking table 替换 `MIGRATION_CURRENT_MAJOR`；给每项废弃 API 提供 before/after 示例。
  6. 仅在实现和测试完成后新增以下用户要求的文档主题：`workbook-request.md`、`mapping.md`、`profiles.md`、`dynamic-columns.md`（扩充）、`templates-and-comments.md`、`styles.md`、`validation.md`、`csv.md`、`streaming-and-limits.md`、`release-migration.md`。若沿用现有文件，使用明确目录索引，不重复制造互相矛盾页面。
  7. 将文档示例纳入 `Bing.Offices.Docs.Tests` 或等价 consumer 编译测试。
- **验收**：没有占位 major、失效索引或与代码矛盾的 Continue 描述；每个公开示例能以本地 pack 的 NuGet 消费者编译/运行。

## 10. Phase 6：构建、包、互操作和发布门禁

### RH-601 自动化构建、测试、pack 与包消费者验证

- **优先级**：P0
- **依赖**：全部实现任务
- **目标**：在干净或已记录的工作区中验证可构建、可测试、可打包、可消费。
- **真实命令矩阵**：

```powershell
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore
dotnet pack Bing.Offices.sln -c Release --no-build --no-restore
```

- **补充验证**：根据 Docs Tests 的实际项目配置，使用本地 `artifacts`/输出包执行其 consumer test；运行 API contract tests；运行 P0 定向测试；执行 `git diff --check`。
- **环境策略**：缺少 net5/net7/netcoreapp3.1 runtime 时，记录为 testhost 环境阻塞并至少完成编译；不得伪称 `dotnet test` 全 TFM 通过。
- **验收**：所有可运行的 net6/net8 Unit 和 Integration 通过；Release build、pack、Docs consumer、API contract 通过；阻塞项有命令输出和环境说明。

### RH-602 Office / LibreOffice / WPS 人工互操作验证

- **优先级**：P0
- **依赖**：RH-601
- **目标**：补齐 NPOI 自动测试无法证明的主流办公软件互操作证据。
- **测试矩阵**：

| 客户端 | 格式 | 场景 |
|---|---|---|
| Microsoft Excel（当前支持版本） | XLSX/XLS | 打开、保存、重开；样式/ARGB、批注、合并、公式、验证、图片、失败工作簿。 |
| LibreOffice Calc（记录版本） | XLSX/XLS | 打开、保存、重开；日期、列表验证、样式/模板、失败输出。 |
| WPS（记录版本） | XLSX/XLS | 打开、保存、重开；批注、图片、模板、多级表头、数据验证。 |

- **实施步骤**：准备可重复生成的 fixture，并记录软件版本、操作系统、文件哈希、截图/录屏或二次解析结果；对不可用客户端记录阻塞，不以猜测代替结果。
- **验收**：所有 P0 格式/特性组合有通过、已知限制或阻塞证据；已知限制进入 release notes/migration 文档。

### RH-603 发布结论和人工审批包

- **优先级**：P0
- **依赖**：RH-601、RH-602
- **目标**：输出诚实的 Release/No-release 建议和需人工审批的清单，不执行发布。
- **输出内容**：
  1. P0/P1/P2 完成矩阵与未完成项。
  2. 测试、benchmark、resource probe、互操作证据索引。
  3. API breaking table、SemVer 结论、迁移文档清单。
  4. NuGet 元数据、版本、依赖、license/readme/repository URL、包内容、签名策略（如仓库已有）检查结果。
  5. 已知限制：NPOI DOM 内存特征、未支持 Workbook validation 公式、HSSF 颜色边界、缺失运行时或客户端。
  6. 人工批准点：版本号、release notes、NuGet 源、签名、tag、publish 命令。
- **验收**：没有 P0 未解决项、无未解释的失败/阻塞、无伪造性能/互操作数据时才建议 Release；否则明确 `NO-RELEASE`，并列出阻断项。无论结论如何，Agent 不执行 publish。

## 11. 跨阶段质量与安全约束

- 所有源码、Markdown、XML、JSON、YAML、脚本以 UTF-8 读写；PowerShell 写文件必须显式 `-Encoding utf8`。实施优先使用项目编辑工具，避免终端内联大段中文写入。
- 不引入生产 `InternalsVisibleTo`，不把 NPOI 类型泄漏到 Abstractions/Core public API。
- 所有 stream 由调用方拥有；不得关闭输入、输出、模板或 failure destination，除非 API 明确约定并有测试。
- 资源限制属于安全边界：MaxBytes、MaxInputBytes、错误数、图片、Regex 和缓存均必须在实际边界处生效，避免无界内存增长、算法放大或部分失败后的不确定状态。
- 对日期/数值/公式解析使用明确文化和 invariant format；不得接受用户输入作为任意 CLR 类型、反射目标、文件路径或外部命令。
- 只在确有必要时添加 abstraction；避免顺手重命名、格式化或与发布整改无关的重构。

## 12. 完成定义

本任务可标记 `COMPLETED` 的必要条件：

1. RH-101 至 RH-105 的 P0 行为已实现并有直接回归测试。
2. RH-201 的 breaking table 已批准，RH-202 的实际 public surface、迁移文档和包消费者测试一致。
3. Mapping/style/template/failure workbook 的边界没有未文档化的静默降级。
4. Regex、ValidationRangeIndex、UniqueTracker、导出动态/样式热点的优化均有功能和性能证据。
5. Release build、可运行 net6/net8 Unit/Integration、pack、Docs consumer、API contract 均通过；不可运行 TFM 有明确环境记录。
6. Benchmark 和资源探针不再把人为 buffer 包装或 retained capacity 隐藏为历史性能优势。
7. Excel、LibreOffice、WPS 的人工互操作矩阵有可审计结论，或明确列为 release blocker。
8. README、Excel 文档、迁移文档、release notes 与最终 API/实现一致。
9. 最终报告明确 Release/No-release，且没有自动 commit、tag、push 或 publish。

## 13. 计划阶段完成说明

本 `plan.md` 是当前阶段唯一允许写入的产物。后续由执行 Agent 从 RH-00 开始按依赖顺序实施；若在实施时发现指定外部审查报告、历史 baseline 包或办公软件环境缺失，必须将其登记为证据缺口或 release blocker，而不是编造完成结果。
