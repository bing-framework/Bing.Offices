# API 易用性改进方案

Task ID: BO-API-USABILITY-20261025

## 授权与范围

用户要求“继续实现发现的问题，先做解决方案，然后再实现功能”。本计划承接本轮审查的六项发现；先完成计划文件，再进入实施，不重复请求执行许可。create-plan 的阶段停止点以用户本次连续授权为准。只增加兼容 API，不删除或改名现有公共成员，不改变包版本，不提交、推送、发布或发送外部通知。

基线：HEAD `7496bd40a29213e0d0ba03836943ffb02d5e7bc8` 加当前工作区既有功能修复及优化；已有未提交改动全部保留。旧任务 BO-IMPORT-CI-OPTIMIZATION 已结束，本任务单独记录证据。

## 当前实现和差距

| 发现 | 真实现状与证据 | 本任务目标 |
| --- | --- | --- |
| U-01 正则示例 | import-validation.md 的 verbatim 正则含双反斜杠，实测 ORD-123 不匹配；当前 Docs 测试只构造示例模型 | 修正示例，并执行真实有效/无效导入断言 |
| U-02 模板所有权 | ExcelExport.UseTemplate 默认 leaveOpen=false，NpoiExcelExporter finally 释放模板；async-io.md 却笼统声称所有 Stream 均保留 | 明确模板例外及默认值，保留既有释放行为 |
| U-03 默认语义 | ExcelSheetImportBuilder 默认 StopOnFirstFailure、MaxReadColumns=100；执行器逐行 continue | 明确 StopOnFirstFailure 仅当前行、宽表限制和坐标约定，直接测试多行行为及列上限 |
| U-04 接入入口 | Core CSV 默认实现已存在；注册藏在隐藏的 ExcelMappingPlanFactoryProvider.RegisterDefault；文档样例依赖测试预置变量 | 新增可发现的 AddBingOfficesCore，提供独立可运行的 Excel/CSV 快速开始 |
| U-05 单表导入 | IExcelImporter 只接受 Workbook Request，调用者必须声明根模型 | 在 Core 增加委托既有管线的单表扩展，直接返回 Items 与完整错误 |
| U-06 结果一致性 | Excel 有 IsSuccess/ErrorsTruncated；CSV 只有 IsTruncated；资源限制和错误截断含义不同 | 增加统一 IsSuccess/IsComplete/StopReason 读取方式，保留原属性与构造签名 |

主要功能已真实实现，非 stub。缺口集中在入口、语义和文档。性能热点、索引、异步 staging、文件提交、并发控制和映射缓存不在此次重构范围内。

## API 与行为设计

### 服务注册

在 `Bing.Offices.Extensions.BingOfficesServiceCollectionExtensions` 增加 `AddBingOfficesCore(IServiceCollection)`：注册默认校验规则及已有 Core 工厂、CSV 服务、配置加载和文件提交服务。全部复用 TryAdd/TryAddEnumerable，保持预注册扩展优先且重复注册幂等。NPOI 注册入口委托此方法，再注册 Excel Provider；旧 Provider SPI 保留。

### 单表入口

新增 `ExcelSingleSheetImportExtensions` 六个方法：`ImportSheet<T>`、`ImportSheetAsync<T>`、`ImportSheetFromFile<T>`、`ImportSheetFromFileAsync<T>`、`ImportSheetFromBytes<T>`、`ImportSheetFromBytesAsync<T>`。

参数依次为 importer、source/path/content、可选 selector、可选 `Action<ExcelSheetImportBuilder<T>> configure`、可选 `ExcelResourceLimits resourceLimits`、可选 CancellationToken。selector=null 表示物理索引 0；按名称/索引选择沿用 ExcelSheetSelector。通用复杂需求仍使用 Workbook Request。

Core 内部根模型初始化集合；构建一个 Sheet 的 Workbook Request，调用现有接口或文件/字节扩展，转换结果。新增 Abstractions `ExcelSingleSheetImportResult<T>`：Items、Errors、SourceRows（零基，逐项对应）、ErrorsTruncated、MaxErrors、IsSuccess、IsComplete、StopReason。其公开构造器隐藏于 IntelliSense，供 Core/第三方 Provider 适配，不增加生产友元程序集。普通流始终由调用方拥有，取消和异常原样传播。

### 统一结果摘要

新增公共 `ImportStopReason`：None、InvalidHeader、ResourceLimit、ErrorLimit、Unspecified。Excel Workbook、CSV 和单表结果都暴露 `IsComplete` 与 `StopReason`；CSV 增加 `IsSuccess`。

这是对既有结构化诊断的统一摘要：存在资源限制优先 ResourceLimit；否则错误收集达到上限为 ErrorLimit；否则结构拒绝为 InvalidHeader；无法归因的旧 CSV 截断为 Unspecified；其余 None。`IsComplete = StopReason == None`；普通行转换/校验错误可以 IsComplete=true、IsSuccess=false。None 不表示所有行有效，也不表示底层物理文件被读到 EOF；StopAtFirstEmptyRow 等显式范围终点算按配置完成。多 Sheet 的摘要按上述优先级合并，明细仍在 Errors 中；抛出的取消/IO/扩展异常不伪装成返回结果。

保留 ErrorsTruncated（仅错误收集）和 CSV IsTruncated（旧行为）全部语义。Excel MaxRows 耗尽且尚有请求 Sheet 时，目前循环直接 break 而无诊断；为避免新摘要错误声称完成，补充 ResourceLimit 错误，不再静默省略后续 Sheet。恰好到达最后一个 Sheet 的末行不额外报告超限。

不引入第二套解析器、状态缓存、全局可变默认值或泛化导入框架。

## Phase 1：语义修复

### U-01 / U-02 / U-03

- 目标：修复错误示例和所有权说明，公开默认校验范围。
- 已确认范围：docs/excel/import-validation.md、async-io.md，ExcelExport.cs、ExcelImport.cs、ValidateMode.cs，DocsConsumerTest.cs；新建职责测试文件。
- 步骤：正则改为单反斜杠；基于 Markdown 原文编译出的模型实际验证好/坏数据；模板文档说明默认关闭、leaveOpen:true 保留；增加默认值表，说明零基配置和一基错误坐标；多行校验不停止全表。
- 依赖：无。
- 验证：Docs 的有效/无效数据；XLS/XLSX × Sync/Async 多行失败后继续、100/101 列默认与显式扩容；模板现有职责测试复用并针对说明核对。
- 风险：修改文档 fence 可能影响已有按序号提取测试，应同步更新并保持真实语义断言。
- 验收：合法 ORD-123 成功，非法值产生结构化错误；无改变模板默认所有权、校验默认或列上限。

## Phase 2：接入和单表能力

### U-04

- 目标：Core-only CSV 注册不需要发现 Provider SPI，也不需要依赖 NPOI。
- 已确认范围：新增 Core 服务扩展；NPOI Extensions.Service.cs；新增 CoreServiceRegistrationTest.cs。
- 步骤：搬接现有默认注册；保留旧 SPI；验证重复调用及注册顺序；保持调用方自定义实现。
- 依赖：Phase 1。
- 验证：Core-only CSV 导入/导出完整内容、默认校验；注册两次和与 NPOI 两种顺序；预注册 importer/exporter/规则及工厂保留。
- 风险：缺省校验器重复或消失；测试应检查服务行为和注册数量。
- 验收：只有 Core 的入口即可使用 CSV；原 AddBingOfficesNpoi 行为不变。

### U-05

- 目标：单表导入无需声明根模型。
- 已确认范围：新增 ExcelSingleSheetImportResult.cs、ExcelSingleSheetImportExtensions.cs 及职责测试。
- 步骤：内部根模型、单 Sheet 请求、六入口薄封装、错误/源行/截断无损传递。
- 依赖：U-06 的统一摘要可在同一实现阶段集成。
- 验证：六入口在 XLS/XLSX 的真实成功/失败；默认首 Sheet、名称选择、映射与资源配置；完整错误、源行、流保持打开；取消原样传播；假 Provider 仅用于直接验证委托/异常，不替代真实 Provider 场景。
- 风险：重载歧义、丢失 Workbook 级错误、复制行集合和吞取消。
- 验收：`importer.ImportSheet<Order>(stream)` 直接得到 Items；缺 Sheet/缺列仍有完整错误；实际调用原 Workbook 管线。

## Phase 3：统一结果摘要

### U-06

- 目标：跨 Excel/CSV 的调用者可以用同名属性区分成功、已处理完但有行错误、结构/资源/错误上限导致的不完整。
- 已确认范围：ImportStopReason.cs、CsvImportResult.cs、ExcelWorkbookImportResult.cs、单表结果、NpoiExcelImporter.cs；新增 ImportResultStatusTest.cs。
- 步骤：实现只读摘要；保留旧构造器和旧属性；补齐 MaxRows 在 Sheet 边界的结构化错误。
- 依赖：核心结果类型。
- 验证：正常、行转换失败、InvalidHeader、MaxRows、MaxErrors、无原因旧截断、多原因优先级；XLS/XLSX × Sync/Async；跨 Sheet 耗尽和末 Sheet 恰好额度；CSV Sync/Async 真实输入；部分成功 Items 保留。
- 风险：把 ErrorsTruncated 错当所有截断、误报恰好上限、覆盖原有结果契约。
- 验收：统一摘要真实反映既有错误，不制造成功结果；旧字段所有断言继续通过。

## Phase 4：快速开始、契约和收口

### U-07 文档接入与契约验证（支撑 U-04/U-05/U-06）

- 目标：完成首次接入闭环，公开 API 增量可审计。
- 已确认范围：README.md、docs/excel/README.md、新增 quick-start.md、nuget-migration.md、Docs 测试、PublicApiContractTest.cs、build/api-snapshot-baseline.json、任务内 api-changes.md 与 symbol-test-map.md。
- 步骤：说明只用 CSV 安装 Core、Excel 安装 Npoi；提供包含全部 using、注册、模型、数据、导入导出和错误处理的完整程序；新快速开始代码原样提取编译运行，不补隐藏前置变量；首页和专题入口链接；新增类型分类与精确类型表；构建双 TFM 快照候选，核对只增加计划内成员后更新基线并记录增量。用户已授权上述 API 能力实施及项目强制配套快照维护；此处不作为批准其他删除/破坏性差异的依据。
- 依赖：所有实现。
- 验证：Release solution build；net6/net8 Unit、Integration；net8 Docs；API 成员和类型契约；docs-check；git diff --check。
- 风险：错误更新掉历史 API、样例只可编译不可运行；任何计划外 API 删除均拒绝纳入基线。
- 验收：原有公开成员没有删除/签名变化；新增成员逐个映射到职责测试；六项发现闭环。

## 验证命令与影响分析

已发现真实入口：Bing.Offices.sln；tests/Bing.Offices.Tests（net6.0/net8.0）；tests/Bing.Offices.Tests.Integration（net6.0/net8.0）；tests/Bing.Offices.Docs.Tests（net8.0）；build/ApiSnapshot；tools/docgen/run.py。

```powershell
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build --no-restore --filter "FullyQualifiedName~Usability|FullyQualifiedName~SingleSheet|FullyQualifiedName~CoreServiceRegistration|FullyQualifiedName~ImportResultStatus"
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-build --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net6.0 --no-build --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-build --no-restore
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --no-build --no-restore
python tools/docgen/run.py docs-check
git diff --check
```

ChangedProjects：Abstractions/Core/Npoi + Unit/Docs。ChangedPublicContracts：兼容新增。ChangedRuntimePaths：DI、结果适配、Sheet 边界限制诊断。ChangedProviders：NPOI、Core CSV。ChangedTFMs：无。ChangedBuildPackaging：仅快照，包依赖不变。ChangedBenchmarkHarness：无。RiskLevel：MEDIUM。每阶段先精准测试，公开 API 完成后才跑全套。没有 hot-path/IO staging/资源阈值变化，不重跑性能基准和重型资源矩阵；远端 CI/真实部署门槛保持外部待验证。

## 完成条件与任务治理

- 六项发现及支撑文档/快照全部完成，测试通过，新增 API 为兼容增量。
- 保存 execution.md、API 增量与成员测试映射；既有改动不回退。
- 只将真实可修问题标记 OPEN_ACTIONABLE；无此类问题后停止，不为外部 CI 再循环。
- 使用 task-state start 注册执行，task-finish --no-notify 收口，不发送外部通知。
- 默认职责清晰的小任务委派 luna_worker；主代理负责设计、结果整合与最终验证。
