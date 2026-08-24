<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: TASK-BING-OFFICES-20260821-MAPVAL-V2
AI_EXECUTION_FINISHED_AT: 2026-08-24T12:58:42.0603582+08:00

# 实施执行报告

> 本次执行按批准的 `plan.md` 实施；未修改 `plan.md`，未执行 commit/push/PR/发布操作。

## 执行结论

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- 当前状态：`PARTIAL`
- Compatibility Mode：`MIGRATION_CURRENT_MAJOR`
- 当前主阶段：Phase 6，`PARTIAL`
- 执行器：Copilot
- 未执行 git commit、git push、PR、Tag 或 NuGet 发布。

## 任务信息

- 计划文件：[plan.md](plan.md)
- 执行开始时间：2026-08-21T13:18:23.829Z
- 当前 HEAD：`73883f709f8eb9e58cd948db0bf90e82ca44a661`
- 当前分支：`master`

## 计划执行情况

| 阶段 | 状态 | 说明 |
| --- | --- | --- |
| Phase 0 基线与设计冻结 | `DONE` | Git、restore、build、Unit、Integration、Docs Consumer 基线已记录 |
| Phase 1 Normalized Mapping 与不可变 Plan | `PARTIAL` | 已完成方向配置、快照、覆盖合并和 CSV/NPOI 接入；完整 Workbook Plan 尚未形成 |
| Phase 2 Validation V2 与 Unique | `PARTIAL` | 新规则和 pending journal 已进入主链；FirstRowNumber/comparer/统一 Descriptor 尚未闭环 |
| Phase 3 Profile V2、Registry 与 DI | `DONE` | 双模型 Profile、方向 Builder、Registry、显式注册和程序集扫描已验证 |
| Phase 4 JSON/XML v2 | `PARTIAL` | v1/v2 normalized loader 和安全限制已完成；diagnostic/round-trip/alias registry 未完成 |
| Phase 5 Provider SPI、IVT 与 API | `PARTIAL` | 生产 IVT 已清理，Public API allowlist 已验证；窄只读 SPI 尚未完全迁移 |
| Phase 6 文档、性能与完整验证 | `PARTIAL` | 文档、隔离 PackageReference consumer、完整测试和 Mapping benchmark 矩阵已有证据；生产 tenant cache、LOH 和完整文档 fence 编译仍未完成 |

## 基线问题

- `dotnet restore Bing.Offices.sln --locked-mode` 于 2026-08-21 失败，`NU1004` 报告多个项目的 `packages.lock.json` 依赖表达式与当前项目声明不一致；详见 `00-baseline.md`。
- 工作区在本 Task 开始前已有 13 个 Metadata/Excels 文件删除，不能回滚或覆盖；本 Task 只记录并避开这些文件。
- 指定方案文档和评审文档仍未在仓库中找到；以当前源码、项目文件和测试为准。

## 基线验证结果

- `dotnet restore Bing.Offices.sln --force-evaluate`：通过；未产生受跟踪 lock 文件差异。
- `dotnet build Bing.Offices.sln -c Release --no-restore`：通过，73 个既有警告，0 错误。
- Unit net6：171 通过，0 失败。
- Unit net8：171 通过，0 失败。
- Integration net6：10 通过，0 失败。
- Integration net8：10 通过，0 失败。
- Docs Consumer net8：6 通过，0 失败。

## 已完成事项

### Mapping/Profile/Registry

- 新增方向化双模型 Profile、Import/Export Builder、快照隔离、Profile Registry、显式注册和程序集扫描。
- `ExcelTypeMapFactory` 已接入方向配置、Alias、Validation Remove/Clear/Append/Replace 和 ValueMap Append。
- CSV/NPOI 已使用方向化 Document/Profile 和列级空白策略。

### Validation/Unique

- 新增 `ExcelRequired`、`ExcelRegex`、`ExcelDate`、`ExcelRange`、`ExcelMaxValue`、`ExcelUnique`、`ExcelMaxLength`。
- Regex 使用 timeout 缓存；MaxLength/MaxValue 使用独立错误码。
- CSV/NPOI Unique 使用 committed state + pending row journal，并支持 `MaxTrackedUniqueValues`；失败行不提交唯一值。

### JSON/XML/Security

- 新增 `ExcelMappingDocument` 和 Loader document API；v1 平铺 JSON/XML 通过 adapter 归一化为 v2，并继续支持旧 Loader 返回 Import 配置。
- JSON 具有最大深度、最大文档大小、字符串/列/别名/规则限制、严格语法和未知字段拒绝。
- XML 使用 `DtdProcessing.Prohibit`、`XmlResolver = null`、文档字符限制和未知节点/属性拒绝。
- JSON/XML 文件入口改为受限流读取，避免 `ReadAllText` 绕过大小上限。

### API/NuGet/Docs

- 生产程序集 IVT 已删除，仅保留测试友元；Public API allowlist/hash 与 provider-neutral consumer 已验证。
- 三包已本地 pack；隔离 consumer 仅从 `artifacts/packages` 恢复、编译并运行成功，输出 `pack-consumer-ok`。
- Docs Consumer 覆盖 AddNpoi、Workbook 请求、v2 Document、双模型 Profile 和 Registry，`3/3` 通过。
- 保留用户已有 13 个 `Metadata/Excels` 删除项。

## 测试结果

- Unit net6：`171/171` 通过。
- Unit net8：`171/171` 通过。
- Integration net6：`10/10` 通过。
- Integration net8：`10/10` 通过。
- Docs Consumer net8：`6/6` 通过。
- Loader 安全专项：`8/8` 通过。
- `get_errors`：本轮修改的 Loader、StreamPipelineTest、DocsConsumerTest 均无错误。

## Build/Typecheck/Lint/Format

- `dotnet restore Bing.Offices.sln --force-evaluate`：通过；未产生受跟踪 lock-file 差异。
- `dotnet restore Bing.Offices.sln --locked-mode`：失败，历史 `NU1004`，详见 `00-baseline.md`。
- `dotnet build Bing.Offices.sln -c Release --no-restore`：通过，0 错误，180 个警告。
- `git diff --check`：通过；仅有 Git 对既有 CRLF 文件的换行提示。
- 未配置独立 lint/formatter 命令；C# 编译器、xUnit 和 API contract test 已执行。

## Benchmark/GC

- 命令：`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --filter "*MappingValidationBenchmarks*" --job short --warmupCount 0 --iterationCount 1 --launchCount 1`；另执行 `MultiRulePlanBuild` 16 组补测。
- 1K Import：平均 `15.90 ms`，分配 `16.36 MB`，Gen0 `984.3750`、Gen1 `921.8750`、Gen2 `93.7500`。
- 1K Export：平均 `15.68 ms`，分配 `8.26 MB`，Gen0 `500.0000`、Gen1 `343.7500`、Gen2 `218.7500`。
- 未宣称零 GC；LOH 峰值和 Peak Working Set 未由当前 Benchmark 报告测量。
- MappingValidation 完整初次矩阵执行 192 个组合，其中 176 个初次有结果；修复 benchmark 输入后，MultiRule 16 个组合全部通过。代表结果和原始日志记录在 `06-performance-gc.md`。

## 部分/未完成事项

- 完整 immutable `WorkbookPlan/SheetPlan/ColumnPlan` provider-neutral compiler 尚未形成，当前仍以 `ExcelTypeMap` 和 NPOI `ExcelColumnPlan` 适配为主。
- 尚无正式 diagnostic/round-trip loader API、model alias registry 和完整精确 JSONPath/XML path 错误模型。
- Unique FirstRowNumber 错误元数据和公开 comparer 选项尚未闭环；`ExcelValidationContext.DuplicateValues` 兼容字段仍存在。
- Provider SPI 尚不够窄；生产 IVT 已清理，但内部类型映射调用链尚未完全迁移到只读 Plan view。
- 已执行 10K/100K 行、1/5 Unique 列、Plan Build、租户缓存、JSON/XML v1/v2 parse、10K 多规则和注册扫描性能矩阵；生产 tenant cache、LOH 峰值和独立资源上限仍未完成验收。
- 未执行外部 Office/LibreOffice 重开。

## 计划偏差

- 计划要求的中心 immutable Plan、诊断/round-trip、窄 Provider SPI 和扩展性能矩阵超出本次已完成实现范围，因此任务终态为 `PARTIAL`，没有通过降低断言或跳过测试伪造完成。
- 文件入口限额修复属于 MAPVAL-401 的直接安全收口，已补测试并验证。

## 基线问题

- locked restore 的 `NU1004` 来自既有 lock 文件依赖表达式与当前项目声明不一致；执行使用 `--force-evaluate`，未改写受跟踪 lock 文件。
- 工作区在 Task 开始前已有 13 个 `Metadata/Excels` 删除项，本 Task 未恢复、覆盖或清理。
- 指定方案文档和评审文档未找到，以仓库规则、源码、测试和 `plan.md` 为准。

## Git 状态

- 当前分支：`master`。
- Task 开始时 HEAD：`73883f709f8eb9e58cd948db0bf90e82ca44a661`。
- 未执行 `git add`、`git commit`、`git push`、PR、Tag、reset、clean 或 NuGet 发布。
- 用户已有删除项保持不变；本 Task 的代码、测试、文档和证据目录均保留在工作区供后续 review。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或发布。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：完整 Workbook/Sheet/Column provider-neutral Plan 和五层来源编译器仍未形成。
- 本轮处理：复验 NPOI 列计划消费路径；未扩大架构范围。另修复该路径中验证上下文属性元数据缺失的回归。
- 验证：Unit net6/net8 各 `168/168`；Integration net6/net8 各 `10/10`。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：NPOI 与 Core 映射实现的完全解耦、最小只读 Provider SPI 和公共 API 收敛仍未完成。
- 本轮处理：未伪造 SPI 完成；仅保持 NPOI 内部 `ExcelColumnPlan` 的适配修复，不新增 Provider 公共成员。
- 验证：Public API 契约 `6/6`；Release build 通过。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：modelAlias/profile registry 消费、完整 round-trip schema 和精确 XML path diagnostics 仍未闭环。
- 本轮处理：保留已有 v1 migration diagnostics 和安全限制实现，未将不完整能力标记为完成。
- 验证：Loader/Review Fix 定向测试保持通过；Release build 通过。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`DONE`
- 根因：此前 locked restore 使用的 lock 状态与项目依赖声明不一致。
- 修复：执行 `dotnet restore Bing.Offices.sln --force-evaluate` 后重新执行 locked restore；未删除 locked mode，也未产生受跟踪 lock 文件修改。
- 验证：当前工作区 `dotnet restore Bing.Offices.sln --locked-mode` 退出码 `0`；Release build 退出码 `0`。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：统一 Validation Descriptor、CSV/NPOI 单一规则执行链和完整规则类型矩阵仍未完成。
- 修复：在 `ExcelColumnPlan` 保存 NPOI 内部目标 `PropertyInfo`；`ValidateRawValues` 和 `ValidateColumnValue` 将该元数据传入 `ExcelValidationContext`，避免命名规则读取 `context.Property` 时得到 null。
- 文件：`src/Bing.Offices.Npoi/ExcelColumnPlan.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`。
- 验证：命名校验上下文回归测试 net6/net8 各通过；Unit 全量 net6/net8 各 `168/168`。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：Registry 的多程序集顺序无关、缓存隔离和启动期完整诊断仍缺少直接证据。
- 本轮处理：保留已有显式注册、重复注册、扫描过滤和并发读取修复；未将覆盖不足标记为完成。
- 验证：Review Fix 定向测试 `15/15`；Release build 通过。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：Docs Consumer 仍不是完整本地包隔离 consumer，文档示例和 ASP.NET Core 覆盖不完整。
- 本轮处理：未修改既有文档/Consumer 范围。
- 验证：Docs Consumer net8 `3/3`。

#### FIX-008

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：批准的 10K/100K、Plan/cache、配置解析、注册扫描、LOH/Peak Working Set 性能矩阵仍未完整测量。
- 本轮处理：保留已有 Unique ShortRun 结果及未测量项说明；未宣称零 GC 或完整矩阵完成。
- 验证：已有 Unique benchmark 结果保留；本轮代码变更后的 Release build 通过。

### 本轮回归修复专项

- `Import_NamedValidationRule_ShouldExposeFullValidationContext`：net6/net8 定向通过；`ExcelValidationContext.Property` 现在为目标属性元数据。
- `Import_ReadOnlyMappedProperty_ShouldThrowConfigurationException`：net6/net8 定向通过；导入列绑定阶段检测只读属性并抛出 `InvalidOperationException`，不再延迟为行级转换错误。
- `get_errors`：本轮修改的 `ExcelColumnPlan.cs`、`NpoiExcelImporter.cs` 和相关测试无错误。
- 最终验证：Unit net6/net8 各 `168/168`；Integration net6/net8 各 `10/10`；Docs net8 `3/3`；Public API `6/6`；locked restore、Release build、`git diff --check` 均通过。

### Round 1 终态

- 当前执行状态：`PARTIAL`
- 未完成的 MUST_FIX：FIX-001、FIX-002、FIX-003、FIX-005、FIX-006、FIX-007、FIX-008。
- FIX-004 的当前工作区 locked restore 门禁已通过，但仍需下一轮独立 Reviewer 验证其可重复性和依赖状态。

### Round 2

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或发布。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：完整 Workbook/Sheet/Column immutable Plan、五层来源编译、预绑定规则执行链和生产 tenant cache 仍未形成。
- 实施文件/符号：`ExcelMappingPlanFactory`、`IExcelCompiledMappingColumn`、`CsvPropertyBinding`、`CsvEntityPipeline`、`MappingProfileRegistryTest`。
- 本轮处理：Plan 列集合、aliases、validation names 和 value map 改为只读快照；Core 内部提供 compiled PropertyInfo/getter/setter/Attribute 视图，CSV 绑定优先消费该视图并使用 compiled setter，避免同一 CSV 计划再次反射绑定。CSV 已经通过 `IExcelMappingPlanFactory` 进入统一入口。
- 实际验证：Unit net6/net8 各 `171/171`，Integration net6/net8 各 `10/10`；仍缺正式 Workbook/Sheet Plan、五层矩阵、CSV/XLSX 完全等价和生产有界缓存，因此保留 `PARTIAL`。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：Provider public view 已收窄，但 `IExcelMappingPlanFactory` 仍承担编译入口并暴露 `object`/外部配置；NPOI 默认 resolver 仍通过 Core concrete factory 组合。
- 实施文件/符号：`IExcelMappingColumn`、`IExcelMappingPlanFactory`、`IExcelCompiledMappingColumn`、`NpoiMappingPlanFactoryResolver`、`ReviewFixRegressionTest.ProviderPlanContract_ShouldExposeOnlyProviderNeutralMembers`。
- 本轮处理：公开列契约移除 Type、Attribute 集合、GetValue/SetValue 和 Delegate/反射成员；内部 compiled view 不进入公共 SPI；shape 测试扩展到属性、方法、字段、事件、泛型参数、返回值和参数类型，并禁止反射/Attribute/Delegate/Expression 泄露。
- 实际验证：Provider shape 定向测试、Public API approval、PackageReference consumer 和 Release build 通过；factory 外部配置边界及 NPOI/Core 默认组合尚未彻底拆分，保留 `PARTIAL`。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：normalized document 尚未完整进入 parser-validator-resolver-compiler 单链，v2 schema 和 XML 嵌套路径模型仍不完整。
- 实施文件/符号：`ExcelModelAliasRegistry`、`ExcelMappingPlanFactory.Create<T>(ExcelMappingDocument, ...)`、`ExcelMappingPlan.ProfileName`/`ModelAlias`、`ExcelMappingConfigurationLoader`、`StreamPipelineTest`。
- 本轮处理：alias 支持模型类型/Profile 注册和解析；Document Plan 校验 alias 模型类型与 Profile 一致性并写入 Plan identity；XML 未知节点/属性保留 `/ExcelMappingDocument/...` 路径，即使被 XmlSerializer 包装；新增 alias identity 和 XML path 回归断言。
- 实际验证：alias/Plan identity、XML/JSON loader、DTD/XXE、安全限制和 Unit 全量通过；dynamic/validation/style/layout schema、完整 nested XML path 和 registry/compiler fallback 策略仍未闭环，保留 `PARTIAL`。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：CSV/NPOI 仍各自解析 converter/attribute/named rule，`ExcelValidationContext` 仍保留兼容性 DuplicateValues 修改表面，真实 XLSX 七规则等价矩阵不足。
- 实施文件/符号：`ExcelValidationContext`、`CsvPropertyBinding`、`CsvEntityPipeline`、`ExcelColumnPlan`、`CsvImportError`、`ExcelImportError`、`UniqueTracker`。
- 本轮处理：CSV 使用 compiled Attribute 快照和 getter/setter；UniqueComparison 贯通 CSV/Excel 选项；FirstRowNumber 进入 CSV/Excel error contract；重复值路径读取 Tracker 首行号；新增 CSV 首行号回归测试并保持 rollback/comparer/上限测试。
- 实际验证：Unit net6/net8 `171/171`、Integration net6/net8 `10/10`，UniqueJournal 10K/100K、1/5 列矩阵均有结果；统一 Validation Descriptor、可变 DuplicateValues 迁移和完整 XLSX 规则矩阵仍未完成，保留 `PARTIAL`。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：Registry 基础确定性已覆盖，但没有生产级 tenant Plan cache，方向 Builder 也未覆盖完整动态/style/layout/comment/merge 能力。
- 实施文件/符号：`MappingProfileRegistryTest.AssemblyScan_TwoAssemblies_ShouldBeOrderIndependent`、`MappingProfileRegistry`、Profile DI extensions、Benchmark `TenantPlanCache`/`TenantPlanCacheEviction`。
- 本轮处理：新增两个真实程序集输入顺序相反的扫描测试；断言同一 Profile 的模型类型和方向结果一致；保留重复注册、抽象/开放泛型过滤和并发读取测试。
- 实际验证：Registry 测试、Unit 全量、Integration 全量通过；benchmark tenant cache/eviction 明确是内部 Dictionary/Queue，不宣称生产缓存，容量/淘汰和租户隔离实现仍未闭环，保留 `PARTIAL`。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：外部 PackageReference consumer 已可重复验证，但 Markdown 全部 code fence 尚未自动编译，完整文档示例验收仍缺。
- 实施文件/符号：`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`、`DocsConsumerTest`、`docs/excel/mapping-json-xml.md`、`import-validation.md`、`dynamic-columns.md`。
- 本轮处理：保持三个本地 PackageReference 和 ASP.NET Core FrameworkReference；consumer 覆盖 v1/v2 JSON/XML stream ownership、v2 validation/dynamic CSV、IFormFile upload stream、Profile/Registry、NPOI registration 和 basic workbook request。
- 实际验证：最终重新 pack 三包，在 `artifacts/consumer-packages-r2` 隔离缓存中 restore/build/test，Docs Consumer `6/6` 通过；未逐段编译全部 Markdown fence，保留 `PARTIAL`。

#### FIX-008

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：完整生产资源矩阵仍不等价于当前 benchmark 内部场景，BenchmarkDotNet 也未提供独立 LOH 峰值测量。
- 实施文件/符号：`MappingValidationBenchmarks`、`06-performance-gc.md`。
- 本轮处理：完成 192 组合初次 Mapping 矩阵；修复 null Profile 后补跑 `MultiRulePlanBuild` 全部 16 组合；记录 DynamicPlanBuild、TenantPlanCache、TenantPlanCacheEviction、JSON/XML v1/v2、10K 多规则、UniqueJournal、显式注册、程序集扫描和 PeakWorkingSetBytes 的 Mean、Ops/s、Allocated、Gen0/1/2。
- 实际验证：完整原始日志为 `BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.MappingValidationBenchmarks-20260822-135406.log`，多规则补测为 `...-20260822-140827.log`；所有补测组合成功。tenant cache 仍为 benchmark-only Dictionary/Queue，LOH 和独立资源上限未测量，保留 `PARTIAL`。

### Round 2 终态

- 当前执行状态：`PARTIAL`
- 已完成并验证的增量：Provider 列契约收窄、CSV compiled binding、alias/Plan identity、Unique 首行号/comparer、两个程序集顺序测试、隔离 PackageReference consumer、完整 Mapping benchmark 方法矩阵。
- 未完成的 MUST_FIX：FIX-001、FIX-002、FIX-003、FIX-005、FIX-006、FIX-007、FIX-008 的正式架构闭环仍需下一轮 Reviewer 验收。
- 最终 `dotnet restore Bing.Offices.sln --locked-mode` 复核仍受既有 lock-file 依赖表达式不一致（`NU1004`）和同版本本地包哈希变更（`NU1403`）阻断；`--force-evaluate`、Release build、Unit、Integration 和隔离 PackageReference consumer 均已通过。

### Round 3

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或发布。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：保留 Plan 构建期的 getter/setter、converter、validation binding 和 Unique 元数据快照；补充 converter 能力缓存、动态容器跳过静态探测，固定列、动态列和导出列均只在绑定阶段探测一次。
- 验证：`Import_FixedColumnConverter_ShouldBindOnceBeforeCellConversion`、`Import_DynamicColumnConverter_ShouldBindOnceBeforeCellConversion`、`Export_FixedColumnConverter_ShouldBindOnceBeforeCellConversion` 通过；Unit net6/net8 各 `171/171`。正式 Workbook/Sheet 五层 Plan、生产 tenant cache 和完整 CSV/XLSX 单链仍未完成。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：NPOI 默认 resolver 改为调用 `ExcelMappingPlanFactoryProvider`，不再直接构造 Core concrete factory；保留 provider-neutral binding contract，并更新完整 API allowlist/hash。
- 验证：Public API contract `6/6`、Integration net6/net8 各 `10/10`、Release build 通过。`IExcelMappingPlanFactory` 仍保留兼容性 `object`/configuration overload，窄 factory SPI 尚未完成。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：保留 Profile/modelAlias 解析诊断、normalized document 入口和 v1 migration diagnostics；未知 Profile、方向和模型不再静默回退。
- 验证：Loader、alias identity、JSON/XML 安全专项及 Docs Consumer 通过。完整 v2 dynamic/validation/style/layout schema 和精确嵌套 XML path 仍未完成。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`COMPLETED`
- 修复：停止同版本本地包覆盖；版本推进到 `1.0.2`，Docs Consumer 使用三个固定 `1.0.2` PackageReference；取消 `packages.lock.json` 的全局忽略并刷新全部项目锁文件。
- 验证：`dotnet restore Bing.Offices.sln --locked-mode` 通过；`dotnet build Bing.Offices.sln -c Release --no-restore` 通过；三包 `1.0.2` 已 pack；Docs Consumer `6/6` 通过。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：固定列和动态列改用 Plan 绑定结果；统一 binding kind、raw/converted 顺序、Unique first-row metadata 和 comparer 相关路径；converter 绑定重复探测回归已修复。
- 验证：Unit net6/net8 各 `171/171`，Integration net6/net8 各 `10/10`。完整七规则 XLSX 矩阵和 `DuplicateValues/TryAddDuplicate` 兼容面迁移仍未完成。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：保留 registry 扫描顺序、重复注册、并发读取和 Plan cache 容量 256 的实现；DI 使用统一 factory provider。
- 验证：Registry 相关 Unit、Integration 和 Release build 通过。生产 tenant/model/direction/version 隔离 cache、完整方向 API 和多贡献程序集证据仍不足。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：Docs Consumer 保持仅三个 Bing `PackageReference`，固定到稳定 `1.0.2` 包版本；覆盖 NPOI 注册、Workbook、v1/v2 JSON/XML、Profile、CSV validation/dynamic 和 ASP.NET Core upload stream。
- 验证：方案 locked restore 通过；Docs Consumer `6/6` 通过。全部 Markdown C# fence 和完整 ASP.NET Core 失败响应链路仍未逐段编译执行。

#### FIX-008

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 修复：本轮未扩大 benchmark 范围，保留 Round 2 Mapping/MultiRule/Unique 结果和已明确的 benchmark-local cache 限制。
- 验证：Release build、Unit、Integration、Docs 和 pack consumer 通过；生产 cache 淘汰路径、LOH 和独立资源上限仍未测量。

### Round 3 汇总

- MUST_FIX：`8`
- 已完成：`FIX-004`
- PARTIAL：`FIX-001`、`FIX-002`、`FIX-003`、`FIX-005`、`FIX-006`、`FIX-007`、`FIX-008`
- BLOCKED：无
- FAILED：无
- 回归验证：Release build 通过；Unit net6/net8 各 `171/171`；Integration net6/net8 各 `10/10`；Docs Consumer `6/6`；Public API contract `6/6`；locked restore 通过；`git diff --check` 待最终执行。
- 下一步：交回独立 Reviewer 进行再次验收；未执行 commit、push、PR、Tag 或 NuGet 发布。

### Round 4

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或发布。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：正式 Workbook/Sheet Plan、完整五层来源矩阵和生产 tenant/model/direction/version cache 仍未闭环。
- 修复：补齐 `IExcelMappingWorkbookPlan`/`IExcelMappingSheetPlan` 契约骨架；normalized document 的动态列、样式、布局纳入 configuration clone、来源合并和 Plan cache key；保留当前列 Plan 适配范围。
- 验证：v2 extended JSON/XML round-trip 和 nested XML path 定向测试通过；Unit/Integration 全量通过。正式生产 Workbook/Sheet 执行计划和生产 cache 仍未完成。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：factory 兼容 overload 和 Core public helper 仍在公共面，NPOI 仍通过 Core provider 创建默认实现。
- 修复：保持 NPOI 使用 `ExcelMappingPlanFactoryProvider`；更新 Provider/API 反射 allowlist；未伪造窄 factory SPI 已完成。
- 验证：Public API contract `6/6`、Release build、Integration net6/net8 通过；factory 外部配置边界和完全独立 provider package 仍未完成。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：v2 schema 的 dynamic/style/layout 子字段和完整 nested XML diagnostic 以前未进入安全 parser allowlist。
- 修复：`ExcelMappingConfigurationLoader` 增加 JSON/XML 的 dynamic columns、style、layout allowlist、动态列数量和字符串限制；新增 v2 extended schema round-trip 与未知嵌套 XML 路径测试。
- 验证：JSON/XML 定向测试、Unit net6/net8、Integration net6/net8 和 Docs Consumer 通过；完整 validation schema 和所有 nested error 诊断仍未完成。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`COMPLETED`
- 根因：源码公共面、approval、旧 `1.0.2` 包此前不同步。
- 修复：`version.props` 推进到 `1.0.3`；Docs Consumer 固定三个 `1.0.3` PackageReference；重新 pack `Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi`；更新 Public API allowlist/hash。
- 验证：Release solution build 通过；Public API `6/6`；Unit net6/net8 各 `174/174`；Integration net6/net8 各 `10/10`；隔离本地包源 restore 后 Docs Consumer `6/6`。`artifacts/packages` 中存在三包 `1.0.3`。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：忽略属性仍参与 converter capability 探测，且完整统一 Validation Descriptor/XLSX 七规则等价矩阵尚未完成。
- 修复：`ExcelMappingPlanFactory` 对 ignored property 跳过 converter binding；修复固定列 converter `CanConvert` 被无关只读集合属性二次触发的回归。
- 验证：固定列绑定一次、动态列绑定一次和导出绑定一次测试通过；Unit net6/net8 各 `174/174`；旧 Duplicate 可变兼容路径和完整 XLSX 规则矩阵仍未完成。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：生产 tenant cache、两个真实贡献程序集和完整方向能力矩阵仍不足。
- 修复：本轮仅将 extended configuration 纳入 cache key，未把 benchmark-local cache 宣称为生产实现，也未伪造多程序集验收。
- 验证：Registry 相关 Unit、全量 Unit/Integration 和 API approval 通过；生产隔离 cache、启动期诊断和第二真实贡献程序集仍未完成。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：旧 Docs Consumer 使用陈旧 `1.0.2` 包，文档 fence 尚未全部逐段编译，ASP.NET Core 链路覆盖不完整。
- 修复：新增 `DocsExamples.cs`，以仅三个 Bing `1.0.3` PackageReference 编译覆盖 README、Profile、JSON/XML、validation、dynamic、CSV 和 ASP.NET Core upload 示例；上传示例执行 importer 并返回结果；刷新 package assets。
- 验证：隔离包源 restore/build/test，Docs Consumer `6/6`；仍未将 Markdown 9 个 fence 逐段自动生成并验证，也未完成失败文件响应的完整端到端断言。

#### FIX-008

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：生产 cache、LOH 和独立资源 ceiling benchmark 仍未实现。
- 修复：本轮未扩大 benchmark 范围；仅验证 cache key 包含新 normalized configuration 字段，避免把 benchmark-only Dictionary/Queue 当作生产 cache。
- 验证：Release build、Unit、Integration、Docs Consumer 和 package pack 均通过；LOH、Peak Working Set 和生产资源 ceiling 仍未完成。

### Round 4 汇总

- MUST_FIX：`8`
- 已完成：`FIX-004`
- PARTIAL：`FIX-001`、`FIX-002`、`FIX-003`、`FIX-005`、`FIX-006`、`FIX-007`、`FIX-008`
- BLOCKED：无
- FAILED：无
- 回归验证：Release solution build 通过；Unit net6/net8 各 `174/174`；Integration net6/net8 各 `10/10`；Public API `6/6`；Docs Consumer `6/6`；`get_errors` 对本轮修改文件无错误。
- 包验证：三个 `1.0.3` 包已生成并由 Docs Consumer 使用；未执行发布。
- 下一步：交回独立 Reviewer 进行再次验收；未执行 commit、push、PR、Tag 或 NuGet 发布。

### Round 4 终态

<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: TASK-BING-OFFICES-20260821-MAPVAL-V2
AI_EXECUTION_FINISHED_AT: 2026-08-22T23:23:59.8108364+08:00

### Round 5

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或发布。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`COMPLETED`
- 根因：Workbook Plan 及 v2 dynamic/style/layout 仅参与结构和缓存键，CSV/NPOI 仍读取请求期动态定义。
- 修复：请求动态定义归一化到 Mapping Document；Core Plan 编译只读动态列、样式、布局；NPOI 导入导出改经 `CreateWorkbook<T>()` 获取 Sheet Plan；CSV 动态导入导出消费同一 Plan；NPOI 应用动态列物理位置、数字格式和 header/body style key。
- 验证：`Export_DocumentDynamicPlan_ShouldApplyLayoutFormatAndStyle`、`Import_DocumentDynamicValidator_ShouldMatchCsvAndXlsx` 通过；Unit net8 `177/177`、net6 `177/177`。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：公开 compatibility normalizer/factory overload 及 NPOI 内部默认构造兼容路径仍依赖 Core factory provider。
- 修复：本轮将 Provider 执行输入收敛为 immutable Workbook/Sheet/Column Plan，并为新增只读 dynamic/style/layout SPI 增加 public API reflection approval；未删除兼容入口，避免在本轮无迁移入口的情况下破坏现有构造链。
- 验证：Public API contract `6/6`，无 NPOI 类型泄漏或 production IVT；`ExcelMappingPlanFactoryProvider`/`NpoiMappingPlanFactoryResolver` 创建耦合仍待下一轮架构修复。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：动态 converter/validator 曾由 NPOI 在请求执行期重复解析，CSV/XLSX 动态规则输入不一致。
- 修复：Core Plan 构建期解析动态类型、命名 converter 和 validator；NPOI 删除本地动态 binding lookup；CSV/XLSX 均消费 Plan 预绑定 converter/validation；增加同一 Document 命名 validator 的 CSV/XLSX 行坐标和回滚等价测试。
- 验证：动态 Document 专项 `2/2`；既有固定列七规则和全量回归通过。完整七规则 × 固定/动态 × CSV/XLSX 显式矩阵仍未补齐。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 状态：`COMPLETED`
- 根因：锁文件与 PackageReference 不一致，旧 `1.0.3` 包与当前源码内容不一致。
- 修复：`--force-evaluate` 更新全方案 locks；版本推进至唯一 `1.0.4`；重新 pack 三个发布包；Docs Consumer 固定使用本地 `1.0.4` 附加源并在独立 `NUGET_PACKAGES` 中恢复测试。
- 验证：solution `--locked-mode` PASS；隔离 consumer force/locked restore PASS、Docs `7/7`；Abstractions/Core/Npoi 包内 DLL 与同次 Release 输出 SHA-256 均 `Match=True`。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`COMPLETED`
- 根因：原双程序集测试的第二个程序集没有 Profile 贡献。
- 修复：新增最小 `Bing.Offices.ProfileFixtures` 程序集并提供独立 Profile；Registry 测试覆盖两个真实贡献程序集的正反顺序、共存解析和跨程序集重复 key 诊断。
- 验证：`MappingProfileRegistryTest` `7/7`，双目标全量 Unit 通过。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：文档示例方法未被调用，上传测试只打开流且不执行 importer/HTTP 结果。
- 修复：Docs 示例 Profile、JSON/XML、validation、dynamic 和 export 路径由 consumer 测试真实调用；上传使用包内 exporter 生成 XLSX，执行 importer，并断言成功 `200` 与非法文件 `400`。
- 验证：独立 1.0.4 包消费者 `7/7`。当前为职责方法映射执行，尚未建立从 9 个 Markdown fence 自动提取、编译并逐 fence 对账的生成式检查。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 状态：`PARTIAL`
- 根因：benchmark 使用私有 Dictionary/Queue 模拟缓存，未采集生产缓存淘汰及 LOH 指标。
- 修复：`TenantPlanCache`/`TenantPlanCacheEviction` 改为直接驱动生产 `ExcelMappingPlanFactory` 的 tenant cache/capacity；新增 `LohSizeBytes` 并保留 Peak Working Set。
- 验证：BenchmarkDotNet Dry smoke 完成，覆盖生产 cache、eviction 和 LOH；尚未定义并断言可发布的硬资源 ceiling，因此不宣称资源门槛已完成。

### Round 5 汇总

- MUST_FIX：`7`
- 已完成：`FIX-001`、`FIX-004`、`FIX-005`
- PARTIAL：`FIX-002`、`FIX-003`、`FIX-006`、`FIX-007`
- BLOCKED：无
- FAILED：无
- 回归验证：Release solution build PASS；Unit net8/net6 各 `177/177`；Integration net8/net6 各 `10/10`；Public API `6/6`；Docs Consumer `7/7`；solution 与 consumer locked restore PASS；benchmark Dry smoke PASS；`git diff --check` 无 whitespace error，仅现有 CRLF/LF 提示；`get_errors` 无错误。
- 包验证：唯一 `1.0.4` 三包已生成，隔离 consumer 使用成功，三个包内 DLL 与 Release 源构建 SHA-256 一致；未发布。
- 下一步：交回独立 Reviewer 再次验收，重点复核 `FIX-002`、完整七规则矩阵、fence 自动执行和资源 ceiling；未执行 commit、push、PR、Tag 或发布。

### Round 5 终态

<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: TASK-BING-OFFICES-20260821-MAPVAL-V2
AI_EXECUTION_FINISHED_AT: 2026-08-23T10:28:47.0852227+08:00

### Round 6

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或发布。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`src/Bing.Offices.Npoi/Imports/ExcelImportExecutionOptions.cs`、`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`、`tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`。
- 修复：导入导出按规范化映射和模型类型分组，在 workbook scope 一次构建 `CreateWorkbook<T>()`，执行层消费对应 Sheet Plan；增加真实多 Sheet XLSX 执行、旧式 `before-` placement、动态配置快照和并发复用测试；CSV 继续消费同一 Column Plan。
- 验证：Workbook/Request 定向测试 `27/27`；Review Fix 回归 `18/18`；Unit net8/net6 各 `182/182`；Integration net8/net6 各 `10/10`。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修改文件：`src/Bing.Offices.Abstractions/Bing/Offices/Providers/IExcelMappingPlan.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Providers/IExcelMappingWorkbookPlan.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocumentFactory.cs`、`src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`、`src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelTypeMapFactory.cs`、`tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`。
- 修复：所有 Provider Plan/Sheet/Workbook、Column、Style、Layout 和 Dynamic SPI 应用 `EditorBrowsable(EditorBrowsableState.Never)`；保留的 `object profile`/configuration 兼容入口增加 `Obsolete` forwarding 迁移标记，并添加反射断言和 API approval 更新。
- 未完成：NPOI 默认 resolver/DI 仍通过 Core `ExcelMappingPlanFactoryProvider` 创建默认实现，尚未形成独立的 Core owning registration 边界，因此保留 `PARTIAL`。
- 验证：SPI/兼容 metadata 回归通过；Public API contract `6/6`；Release compile 通过；无 production IVT 和 NPOI 类型泄漏。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`PARTIAL`
- 修改文件：`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDynamicColumnConfiguration.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelDynamicColumnDefinition.cs`、动态列 clone/request 文件、`src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`、`src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`、NPOI importer/exporter 和相关测试。
- 修复：动态列新增 `ValidationRuleNames` 集合，保留旧 `ValidatorName` 兼容；集合进入 JSON/XML allowlist、clone、cache key、请求快照和 Core build-time binding；Plan 暴露有效规则顺序，CSV/XLSX 均消费预绑定 `ValidationBindings`，不在 Provider 重新 lookup。
- 验证：动态规则 CSV/XLSX 等价测试和两个命名规则顺序断言通过；JSON/XML v2 dynamic rule round-trip 通过；Unit/Integration 双目标全绿。
- 未完成：本轮尚未补齐 Required、Regex、Date、MaxValue、Range、MaxLength、Unique 七规则的固定/动态 × CSV/XLSX 完整矩阵，因此保留 `PARTIAL`。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`、`tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`。
- 修复：测试输出复制现有 `docs/excel/*.md`；Roslyn 仅测试依赖从 Markdown 原文提取 `csharp` fence，逐个编译为独立 consumer assembly 并执行入口；覆盖 5 个文档文件的 9 个 fence，同时保留 package-only 上传和 stream ownership 测试。
- 验证：`DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` `1/1`，实际 fence 数量 `9`；完整 Docs Consumer `8/8`。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`、`tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`。
- 修复：eviction benchmark 直接比较生产 `ExcelMappingPlanFactory` 的首个 Plan 与容量溢出后的重建 Plan，并输出 `CACHE_EVICTION observed=True`；LOH 和 Peak Working Set 分别输出实际 bytes、硬 ceiling 和 `status=passed`；保留 100/500 PlanBuild、100/1000 Tenant、1/5 Unique 和 10K/100K 参数矩阵。
- 验证：eviction Dry 输出包含 `capacity=50 tenants=100` 和 `observed=True`；LOH 输出 `value=0 ceiling=536870912 status=passed`；Peak Working Set 输出约 `41697280 ceiling=1073741824 status=passed`；生产 cache eviction 单元测试通过；benchmark build 通过。

### Round 6 汇总

- MUST_FIX：`5`
- 已完成：`FIX-001`、`FIX-006`、`FIX-007`
- PARTIAL：`FIX-002`、`FIX-003`
- BLOCKED：无
- FAILED：无
- 回归验证：Unit net8/net6 各 `182/182`；Integration net8/net6 各 `10/10`；Docs Consumer `8/8`；Public API contract `6/6`；benchmark build 和专项 Dry smoke 通过；`git diff --check` 无 whitespace error。
- 包/发布：未执行 commit、push、PR、Tag、NuGet 发布。
- 下一步：交回独立 Reviewer 再次验收；重点复核 Core owning factory boundary 和七规则固定/动态 CSV/XLSX 完整矩阵。

### Round 6 终态

<!-- AI_EXECUTION_STATUS: PARTIAL -->
AI_TASK_ID: TASK-BING-OFFICES-20260821-MAPVAL-V2
AI_EXECUTION_FINISHED_AT: 2026-08-24T10:05:00+08:00

### Round 7

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 Round 7 的 `MUST_FIX`；未修改 `review.md`，未执行 commit、push、PR、Tag 或 NuGet 发布。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactoryProvider.cs`、`src/Bing.Offices.Core/Bing.Offices.Core.csproj`、`src/Bing.Offices.Npoi/Extensions/Extensions.Service.cs`、`src/Bing.Offices.Npoi/NpoiMappingPlanFactoryResolver.cs`、`tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`、`tests/Bing.Offices.Tests/PublicApiContractTest.cs`。
- 根因：NPOI 的 `AddNpoi()` 和默认构造兼容解析器直接调用 Core `ExcelMappingPlanFactoryProvider.CreateDefault()`，Core 默认实现创建责任跨越 owning boundary。
- 修复：新增 Core-owned `ExcelMappingPlanFactoryProvider.RegisterDefault(IServiceCollection)`，使用 `TryAddSingleton` 和 DI 服务集合创建默认 factory，保留用户预注册 `IExcelMappingPlanFactory` 的替换语义；`AddNpoi(): void` 改为调用 Core 注册入口。NPOI resolver 不再调用 `CreateDefault()`，仅为 current-major 无 DI 直接构造路径保留兼容性 `new ExcelMappingPlanFactory(...)` fallback；生产 DI 路径由 Core 注册负责，显式 factory 仍可注入 importer/exporter。
- 验证：默认 DI 与预注册 replacement 测试通过；现有显式 exporter factory 和默认 importer/exporter 构造路径通过；Provider SPI、production IVT、NPOI 泄漏检查和 API approval 通过；Public API net8 `6/6`。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDynamicValidationConfiguration.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDynamicColumnConfiguration.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingConfigurationCloner.cs`、`src/Bing.Offices.Abstractions/Bing/Offices/Providers/IExcelMappingPlan.cs`、`src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`、`src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`、`src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`、`src/Bing.Offices.Npoi/ExcelColumnPlan.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`src/Bing.Offices.Npoi/Imports/NpoiExcelExporter.cs`、`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`、`tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`。
- 根因：动态规则集合虽已预绑定，但缺少七类内置规则与固定/动态、CSV/XLSX 的直接等价证据；动态错误键仍可能回退到承载字典属性名。
- 修复：新增动态规则 descriptor，Plan build-time 将 `required`、`regex`、`date`、`maxValue`、`range`、`maxLength`、`unique` 转为内置 Attribute binding；descriptor 属性进入 clone、JSON/XML allowlist 和 cache key。动态 Unique 的 `IsUnique`/`UniqueIgnoreEmpty` 进入 provider-neutral Plan，CSV/NPOI 统一使用稳定动态 Key 和首行号。NPOI/CSV 错误路径对动态列使用稳定 Key。
- 验证：固定属性和动态列各建立七规则 CSV/XLSX 矩阵，断言正常行、负例、完整错误 Code、RowIndex、ColumnIndex、ColumnKey/PropertyName、失败行回滚、case-insensitive Unique 和 FirstRowNumber；两项矩阵 net8 `2/2`，Unit net8/net6 各 `185/185`，Integration net8/net6 各 `10/10`。动态 ValidationRules JSON/XML round-trip 通过。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`、`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/06-performance-gc.md`。
- 根因：上一轮只有 BenchmarkDotNet Dry 输出；LOH 读取的是强制 GC 后存量，缺少绑定到指定大输入参数的独立进程峰值和原始 artifact。
- 修复：新增 `--resource-probe` 父进程入口，按 `PlanBuildCount={100,500}`、`TenantCount={100,1000}`、`UniqueColumnCount={1,5}`、`UniqueRowCount={10000,100000}` 启动 16 个独立子进程；每个场景真实构建 Plan、运行 Unique journal、保持 90 KiB LOH 对象至测量点，记录 LOH/Peak Working Set、ceiling、exit code 和环境。BenchmarkDotNet `LohSizeBytes()` 同步绑定当前 Unique 行列参数并持有大对象负载。
- 验证：非 Dry 命令 `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --resource-probe artifacts/mapval-v2/resource-probe-round7.jsonl` 通过，原始 artifact 为 `artifacts/mapval-v2/resource-probe-round7.jsonl`，16/16 场景 `exitCode=0/status=passed`；最大 LOH `85,801,616` bytes，最大 Peak Working Set `133,619,712` bytes，分别低于 `512 MiB` 和 `1 GiB` ceiling。性能文档已记录命令、环境、参数、汇总数值和复核路径。

### Round 7 汇总

- MUST_FIX：`3`
- 已完成：`FIX-002`、`FIX-003`、`FIX-007`
- PARTIAL：无
- BLOCKED：无
- FAILED：无
- 回归验证：Unit net8 `185/185`；Unit net6 `185/185`；Integration net8 `10/10`；Integration net6 `10/10`；Docs Consumer `8/8`；解决方案 Release build 通过；Benchmark 资源探针 `16/16`；Public API approval 通过；`git diff --check` 通过，仅有既有 CRLF/LF 提示。
- 下一步：再次交回独立 Reviewer 验收；Review Fix Executor 不代表 Reviewer 已 PASS。

### Round 7 终态

<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: TASK-BING-OFFICES-20260821-MAPVAL-V2
AI_EXECUTION_FINISHED_AT: 2026-08-24T11:17:32.5822000+08:00

### Round 8

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 Round 8 的 `MUST_FIX`；未修改 `review.md`、`plan.md`，未执行 commit、push、PR、Tag 或 NuGet 发布。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`src/Bing.Offices.Npoi/NpoiMappingPlanFactoryResolver.cs`。
- 根因：NPOI 无 DI 兼容 resolver 直接 `new ExcelMappingPlanFactory(...)`，使 NPOI 仍承担 Core concrete factory 的创建责任。
- 修复：resolver 改为调用 Core-owned `ExcelMappingPlanFactoryProvider.CreateDefault(...)`，返回类型仍为 `IExcelMappingPlanFactory`；NPOI 源码不再直接构造 `ExcelMappingPlanFactory`，`AddNpoi()` 的 `TryAdd` replacement 语义保持不变。
- 验证：`MappingPlanFactory_DiDefaultAndReplacement_ShouldPreserveOwnershipBoundary` net8/net6 通过；NPOI 源码搜索确认无 `new ExcelMappingPlanFactory`。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`。
- 根因：既有七规则矩阵只有非边界阈值和单一动态 Unique 配置，缺少边界、空值、大小写比较和资源上限的 CSV/XLSX 直接证据。
- 修复：固定/动态矩阵增加 `MaxValue` 等于上限、`Range` 两端、`MaxLength` 等于上限样本；新增 `Import_DynamicUniqueOptions_ShouldMatchCsvAndXlsx`，同时验证 `IgnoreEmpty=true/false`、`OrdinalIgnoreCase`、首行号、失败行 rollback 和 `MaxTrackedUniqueValues` 上限，并在 CSV/XLSX 两侧断言错误行、稳定键和结果数量。
- 验证：四项定向测试 net8 `4/4`；同组 net6 `4/4`，包含固定/动态七规则矩阵、动态 Unique 选项、DI 默认/替换边界。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`benchmarks/Bing.Offices.Benchmarks/Program.cs`、`benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`、`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/06-performance-gc.md`。
- 根因：上一轮资源 probe 在 `planBuildCount=500, tenantCount=1000` 时没有实际装载 1000 个 tenant，且 LOH 字段把 retained 存量称为 peak；child 结果也以转义字符串嵌套。
- 修复：每个场景先构建 `tenantCount` 个不同 tenant 的 Plan，再执行额外 `planBuildCount` 次构建；LOH 在 tenant/Plan、Unique、payload 阶段采样 `GenerationInfo[3].SizeBeforeBytes`，分别记录 `lohSampledPeakBytes` 与 `lohRetainedBytes`；child `result` 改为结构化 JSON 对象。BenchmarkDotNet 方法同步改用 `LohRetainedBytes` 指标名称。
- 验证：非 Dry 命令 `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --resource-probe artifacts/mapval-v2/resource-probe-round8.jsonl` 通过；artifact `16/16` 场景 `exitCode=0/status=passed`，所有 `tenantCount=1000` 记录 `tenantPlanCount=1000`，最大采样 LOH `85,801,616` bytes、最大 retained LOH `85,801,616` bytes、最大 Peak Working Set `133,238,784` bytes，均低于 ceiling。

### Round 8 汇总

- MUST_FIX：`3`
- 已完成：`FIX-002`、`FIX-003`、`FIX-007`
- PARTIAL：无
- BLOCKED：无
- FAILED：无
- 回归验证：Round 8 定向 Unit net8/net6 各 `4/4`；Benchmark 项目 Release build 通过；资源探针 `16/16`；artifact 结构化解析通过；未执行 commit、push、PR、Tag 或 NuGet 发布。
- 下一步：再次交回独立 Reviewer 验收；Review Fix Executor 的 `COMPLETED` 不代表 Reviewer 已 PASS。

### Round 9

- Review 状态：`NEEDS_FIX`
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/review.md`
- 处理原则：仅处理 Round 9 的 `MUST_FIX`；未修改 `review.md`、`plan.md`，未执行 commit、push、PR、Tag 或 NuGet 发布。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：`tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`。
- 根因：固定列七规则 CSV/XLSX 矩阵缺少 `MaxValue=10`、`Range=1/10` 和 `MaxLength=5` 的成功阈值边界样本；动态矩阵已有对应覆盖，固定/动态证据不对称。
- 修复：在 `Import_FixedBuiltInValidationMatrix_ShouldMatchCsvAndXlsx()` 增加 `10/1/abcde` 和 `10/10/abcde` 两条有效行，覆盖 MaxValue 上限、Range 两端和 MaxLength 上限；同步将成功行数量更新为 3，并将七条失败断言行号更新为 5-11，保留 CSV/XLSX 错误码、列坐标、属性名和 Unique 首行号断言。
- 验证：固定/动态规则及动态 Unique 定向测试 net8/net6 各 `3/3`；Unit 全量 net8/net6 各 `186/186`；Integration 全量 net8/net6 各 `10/10`；所有命令退出码为 0。

### Round 9 汇总

- MUST_FIX：`1`
- 已完成：`FIX-003`
- PARTIAL：无
- BLOCKED：无
- FAILED：无
- 回归验证：定向矩阵 net8/net6 各 `3/3`；Unit net8/net6 各 `186/186`；Integration net8/net6 各 `10/10`。
- 下一步：再次交回独立 Reviewer 验收；Review Fix Executor 的 `COMPLETED` 不代表 Reviewer 已 PASS。
