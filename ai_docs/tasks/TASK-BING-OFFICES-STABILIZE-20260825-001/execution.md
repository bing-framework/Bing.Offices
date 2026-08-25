<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: TASK-BING-OFFICES-STABILIZE-20260825-001
AI_EXECUTION_FINISHED_AT: 2026-08-25T17:58:27.9573002+08:00

# 实施执行报告

## 执行结论

任务以 `PARTIAL` 终态收口，发布结论为 `no-release`。核心 P0 正确性修复、API 收敛、Profile Core 迁移、CSV/NPOI 低风险内部拆分和可支持目标框架验证已完成；批准计划仍有明确 P1 遗留项和外部环境阻断，因此不标记 `COMPLETED`。

未执行自动 git commit、git push、tag、PR、NuGet publish、reset、clean 或其它破坏性 Git 操作。完整最终审查见 [final-review.md](final-review.md)。

本轮已完成 Phase 1 P0 正确性修复、Phase 2 Mapping 方向和单次合并边界的主要路径、Phase 3 已批准 API 收敛、Phase 2.3 Profile Core 迁移，以及 Phase 6 Benchmark 治理的主要结构修正。Phase 4 已完成 NPOI 流复制、失败工作簿 writer 和 CSV header/type/property binding 协作者拆分；NPOI importer/exporter 其余职责拆分、完整 Stream 性能对比、Office 互操作和最终 release/no-release review 仍未完成。

## 已落地变更

- 统一 DI 与直接构造的默认校验规则，补齐 `MaxValueExcelValidationRule`。
- 按 `Disabled`、`ConfiguredRules`、`WorkbookRules`、`ConfiguredAndWorkbook` 分流配置规则与 Workbook 原生规则，修复 Continue 模式错误收集和 Unique journal 回滚。
- 失败工作簿按导入阶段实际解析的 Sheet identity 输出，修复 `ByIndex` 与非零表头行组合。
- 新增模板单元格 `PreserveTemplate`/`ReplaceTemplate` 覆盖策略，统一表头、自定义表头和正文写入行为。
- 将 Document/Request mapping snapshot 的最终合并集中到 Plan Factory；缺失方向默认失败，显式 `UseConventionFallback` 才允许 Convention fallback。
- 删除方向不安全的 JSON/XML facade 和伪异步 `ExportToBytesAsync`，保留 Document 级 JSON/XML 和同步字节导出 API。
- 新增只读 `IMappingProfileResolver` 和稳定 Profile 名称注册重载；保留 FullName 兼容行为，增加程序集部分加载容错。
- 将 Profile DI 注册扩展与 descriptor factory 从 Npoi 迁移到 Core；Npoi 公开面仅保留 `AddNpoi`，同步更新调用方、Docs consumer 和 API 快照。
- Benchmark 使用真实命名校验器，计划构建拆分 cold/cache hit/cache miss，Mapping 基准移除主动 LOH 假负载，资源压力保留在独立 ResourceProbe。
- Benchmark 两个基准类新增固定 `launchCount=1`、`warmupCount=2`、`iterationCount=3` Job；代表性 Mapping cache-hit 基线完成 16 个参数组合，输出 Mean、Allocated、Gen0/Gen1。固定 Job 不绑定 `RuntimeMoniker.Net80`，以兼容本机仅安装 SDK 10、但 net8 runtime 可用的环境。
- 发布包审计部分完成：重新 pack 后三个 2.0.0 包均含 `README.md`、`LICENSE`、`icon.png` 和对应 `.snupkg`；nuspec 含 README 元数据，SourceLink 不再作为运行时依赖。未执行 NuGet publish。
- 新增 `NpoiStreamCopier` 内部协作者，统一 importer/exporter 流复制、取消和输入大小限制逻辑。
- 新增 `NpoiFailureWorkbookWriter` 内部协作者，隔离失败摘要、ErrorRowsOnly 行复制、批注、合并区域、数据验证、图片和目标流写入逻辑；importer 只保留调用边界。
- 新增 `CsvHeaderBinder` 和 `CsvColumn` 内部协作者，隔离 CSV 表头/按位置绑定、重复表头、动态列和 HeaderMatch 校验；导入器继续负责行级读取、转换、校验与 Unique journal。
- 新增 `CsvDynamicTypeResolver` 内部协作者，统一 CSV 导入/导出动态列的受限类型解析和错误语义。
- 将 `CsvPropertyBinding` 移至独立 internal 文件，实体管线继续聚焦导入/导出流程，保留原有反射 fallback 和 compiled mapping 路径。
- 同步 API、Profile、JSON/XML、校验模式、模板策略和主路径文档。

## 验证证据

- net8 Unit：213/213 通过。
- net6 Unit：213/213 通过。
- net8 Integration：11/11 通过。
- net6 Integration：11/11 通过。
- Docs consumer net8：8/8 通过。
- Profile/API 聚焦：17/17 通过。
- 流/导入/导出/失败工作簿重构回归：133/133 通过；对应 Integration：11/11 通过。
- `NpoiStreamCopier` 直接单元测试：79/79 通过，覆盖数据复制、最大输入字节数、预取消和流所有权。
- 失败工作簿 writer 拆分聚焦 Unit：104/104 通过；对应 Integration：11/11 通过。
- CSV 表头绑定拆分聚焦回归：24/24 通过。
- CSV 动态类型解析拆分后，net8 CSV 24/24、net6 CSV 24/24、net8 Integration 11/11 通过。
- `CsvPropertyBinding` 独立文件后 net8 全量 Unit：213/213 通过。
- `CsvPropertyBinding` 独立文件后 net6 全量 Unit：213/213 通过。
- 解决方案 Release build：通过。
- 解决方案 pack：通过，未发布。
- Mapping Benchmark Dry smoke：208 个基准执行，无运行时异常；Dry job 仅报告迭代时间偏短。
- 固定 Mapping 性能基线：16 个参数组合完成；`DynamicPlanBuildCacheHit` Mean 约 262.8 us 至 1.4048 ms，Allocated 约 585.24 KB 至 2,949.93 KB；不作为发布阈值。
- ResourceProbe：16/16 场景通过。
- 源码残留检查：未发现已删除 API、生产 InternalsVisibleTo、伪异步 `Task.FromResult` 或同步等待残留。

## 环境阻断与剩余工作

- net5.0、net7.0、netcoreapp3.1 Unit 已编译，但当前机器缺少对应 runtime，testhost 无法启动；未安装运行时或绕过测试执行。
- Profile Core 迁移已完成；ReflectionTypeLoadException 尚未有真正构造异常的直接测试，稳定 alias 尚未扩展到程序集扫描元数据。由于 Core 没有测试友元，本轮未新增生产 `InternalsVisibleTo` 或测试出口。
- `NpoiExcelImporter`、`NpoiExcelExporter`、`ExcelMappingConfigurationLoader`、`CsvEntityPipeline` 仍需按批准计划继续拆分职责。
- 尚未完成固定 Job 的完整 Benchmark 基线、前后性能对比和高影响复制/cache key 优化。
- Excel/LibreOffice/WPS 互操作、老目标 runtime 运行验证、完整 Stream import/export 性能前后对比、本地包安装消费者链路和最终 release/no-release review 尚未完成。

本任务未宣称 Office 兼容性、旧 runtime 运行通过或完整 Stream 性能优化通过；相关证据边界和解除条件见 [final-review.md](final-review.md)。

## 终态审查

- 终态：`PARTIAL`。
- 发布：`NO-RELEASE`。
- 已完成：核心 P0 修复、可支持 net8/net6 Unit/Integration、Docs Tests、Benchmark smoke、Release build/pack、NuGet 元数据审计和低风险内部拆分。
- 未完成或阻断：NPOI importer/exporter 其余拆分、完整 Stream 性能前后对比、高影响路径优化、旧 runtime 运行、Office 互操作、本地 nupkg 消费者安装链路，以及 Profile 扫描异常直接测试。
- 用户可在解除上述条件后继续运行 `/execute-plan`；本次 finalizer 不执行任何提交或发布操作。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`
- 执行终态：PARTIAL

#### FIX-001

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `tests/Bing.Offices.Tests/MappingProfileRegistryTest.cs`
- 根因：Profile 扫描生产代码已有 `ReflectionTypeLoadException` 分支，但没有真实构造部分加载异常的直接测试。
- 修复：使用测试侧 `Assembly` 子类覆盖 `GetTypes()`，构造含可加载类型和 `LoaderExceptions` 的 `ReflectionTypeLoadException`；新增部分加载保留 Profile、零可加载类型抛出诊断异常两项测试。未新增 Core 生产 `InternalsVisibleTo`、可变生产 seam 或公共 API。
- 验证：
	- `MappingProfileRegistryTest` net8：13/13 通过。
	- net8 Unit 全量：215/215 通过。
	- net6 Unit 全量：215/215 通过。
	- Core 生产 `InternalsVisibleTo` 搜索：无新增结果。

#### FIX-002

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：无仓库消费者文件修改；使用系统 Temp 中的隔离消费者和 NuGet.Config 完成验证。
- 根因：原 Docs Tests 的 PackageReference 还原未证明使用本轮 `output/release` 包。
- 修复：重新 `dotnet pack Bing.Offices.sln -c Release --no-restore`；使用独立 `NUGET_PACKAGES` 目录、临时 NuGet.Config 和本轮 `output/release` 本地源，恢复 Abstractions/Core/Npoi 2.0.0；独立消费者通过 PackageReference 编译并完成 XLSX 导出/导入往返，输出 `PACKAGE_CONSUMER_OK`。
- 验证：
	- 隔离 restore：通过。
	- 独立消费者 Release build：通过。
	- 独立消费者运行时往返：通过。
	- assets 中三个包均为 `bing.offices.*\\2.0.0`，并记录最终包 hash。
	- 最终 pack：通过，未发布。

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`（无本轮代码修改，执行并采集固定基线）
- 根因：无 relation 的导入路径仍无条件创建并填充 `sourceLocations`，而完整 Importer/Exporter/Loader 拆分和高影响路径性能前后对比尚未完成。
- 修复：当 `request.Relations.Count == 0` 时不创建 `sourceLocations` 字典，也不记录成功项位置；存在关系请求时保留原有位置记录和关系错误定位行为。执行完整 Stream benchmark，覆盖 Import/Export 及 1,000、10,000、100,000 行，固定 `launchCount=1/warmupCount=2/iterationCount=3`，输出 Mean、Allocated、Gen0/Gen1/Gen2。
- 性能证据：优化前基准记录的 100,000 行 Import/Export 分配约为 1,634.26 MB/734.94 MB，Mean 约 2.133 s/1.419 s；最终基准约为 1,634.26 MB/734.94 MB，Mean 约 2.200 s/1.476 s。该局部优化未形成可归因的整体 benchmark 降幅，因此不宣称性能阈值达成或 streaming/zero-GC。
- 未完成范围：`NpoiExcelImporter`、`NpoiExcelExporter`、`ExcelMappingConfigurationLoader` 的其余职责拆分，以及高影响复制/cache key 的完整前后对比仍需后续任务处理。
- 验证：
	- Stream/导入/P0 聚焦回归：104/104 通过。
	- net8/net6 Unit：各 215/215 通过。
	- net8/net6 Integration：各 11/11 通过。
	- Docs consumer：8/8 通过。
	- 解决方案 Release build：0 error。
	- `git diff --check`：无空白错误。

### Round 1 汇总

- MUST_FIX：0
- 已完成：FIX-001、FIX-002
- PARTIAL：FIX-003
- BLOCKED：无新增外部阻断
- FAILED：无
- 回归验证：net8/net6 Unit 各 215/215；net8/net6 Integration 各 11/11；Docs 8/8；Release build、pack 和隔离 nupkg 消费通过。
- 下一步：由独立 Reviewer 复审本轮 FIX；FIX-003 剩余职责拆分和高影响性能优化不在本轮宣称完成范围内。

### Round 2

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`
- 执行终态：PARTIAL

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
	- `tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs`
	- `artifacts/review-fix-round2-resource-probe.json`
- 根因：导出器先将 NPOI 工作簿写入内部 `MemoryStream`，再将完整内容复制到调用方目标流，导致导出过程中额外保留一份完整工作簿缓冲；上一轮还未形成真实产品路径的前后性能对比。
- 修复：导出器改为通过非关闭包装直接将 NPOI 工作簿写入调用方目标流，移除导出后的整块 `MemoryStream` 和二次复制；保留目标流 ownership、模板流释放和取消检查语义。新增目标流保持打开且输出可被 NPOI 重开的回归测试。
- 性能证据：在相同固定 Job（`launchCount=1`、`warmupCount=2`、`iterationCount=3`）和 1,000/10,000/100,000 行矩阵下重新执行 Stream import/export。修改后 Export 分配约为 8.53 MB、70.63 MB、761.82 MB；修改前对应约为 8.68 MB、71.36 MB、约 734.94 MB 至 726.53 MB 的历史记录，受运行环境波动影响，不将耗时变化宣称为收益。修改后 100,000 行 Export Mean 为 1.413 s，Import Mean 为 2.112 s；Import 分配和 GC 保持产品路径原有 DOM 特征。未宣称 streaming 或 zero-GC。
- 未完成范围：`NpoiExcelImporter`、`NpoiExcelExporter`、`ExcelMappingConfigurationLoader` 的其余职责拆分，以及完整 cache-key 优化仍需后续处理；因此本 FIX 保持 PARTIAL。
- 验证：
	- `ExcelWorkbookRequestTest` + `StreamPipelineTest`：112/112，通过。
	- net8 Unit：216/216，通过。
	- net6 Unit：216/216，通过。
	- net8 Integration：11/11，通过。
	- net6 Integration：11/11，通过。
	- Stream benchmark：6/6 场景完成，固定 Job、三种规模、Mean、Allocated、Gen0/Gen1/Gen2 已生成。
	- ResourceProbe：16/16 场景通过，结果写入 `artifacts/review-fix-round2-resource-probe.json`。
	- 解决方案 Release build：通过，0 error。
	- VS Code 错误诊断：修改文件无错误。
	- `git diff --check`：通过，无空白错误。

### Round 2 汇总

- MUST_FIX：0
- 已完成：无新增完整 FIX；FIX-003 完成一项可归因的导出缓冲优化和对应验证。
- PARTIAL：FIX-003
- BLOCKED：无新增外部阻断。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 216/216；net8/net6 Integration 各 11/11；Stream benchmark 6/6；ResourceProbe 16/16；Release build 和 diff-check 通过。
- 下一步：重新进行独立 Review；继续处理 FIX-003 剩余职责拆分和 cache-key 优化前，任务不具备 release-ready 结论。

### Round 3

- Review 状态：NEEDS_FIX
- Fix Scope：must
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`，本轮只读，未修改。
- 执行开始：2026-08-25T07:24:03.520Z
- 执行完成：2026-08-25T07:32:17.199Z

#### FIX-004

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
	- `tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs`
- 修复：`NonDisposingStream` 持有 `CancellationToken`，在 `Write` 和 `Flush` 前检查取消；`Write` 完成后再次检查，保持调用方目标流不关闭。NPOI 对底层取消异常进行包装时，Exporter 在确认令牌已取消且根异常为 `OperationCanceledException` 的条件下恢复标准取消异常契约；未恢复完整中间 `MemoryStream`。
- 回归测试：新增 `AddNpoi_MidWriteCancelledExport_ShouldStopWritingAndPreserveDestination`，受控目标流首次底层写入后取消，断言抛出 `OperationCanceledException`、写入次数为 1 且目标流仍可写。
- 验证：
	- 中途取消集成专项 net8：1/1 通过。
	- net8 Unit 全量：216/216 通过。
	- net6 Unit 全量：216/216 通过。
	- net8 Integration 全量：12/12 通过。
	- net6 Integration 全量：12/12 通过。
	- Stream fixed benchmark：Import/Export × 1,000/10,000/100,000 共 6/6 场景完成；固定 Job 输出 Mean、Allocated、Gen0/Gen1/Gen2。最新 100,000 行结果约为 Import 2,339.04 ms、1,623.15 MB，Export 1,506.43 ms、726.8 MB；不宣称 streaming 或 zero-GC。
	- 解决方案 Release build：通过；仅保留既有弃用属性和目标框架支持警告。
	- VS Code 错误诊断：修改的生产文件和测试文件均无错误。
	- `git diff --check`：通过；CRLF/LF 提示为既有换行格式提示，不构成空白错误。

#### SHOULD_FIX 跳过

- FIX-003：本轮 `fixScope=must`，未处理其余 importer/exporter/loader 职责拆分和 mapping cache-key 优化；保留原有 `PARTIAL` 范围和 no-release 结论。

### Round 3 汇总

- MUST_FIX：FIX-004
- 已完成：FIX-004
- SHOULD_FIX：FIX-003 skipped by scope，未伪造 Reviewer 通过。
- 回归验证：net8/net6 Unit 各 216/216；net8/net6 Integration 各 12/12；中途取消专项 1/1；Stream fixed benchmark 6/6；Release build、错误诊断和 diff-check 通过。
- 执行终态：`COMPLETED`（本轮 MUST_FIX 范围完成）。
- 任务整体结论：仍为 `PARTIAL` / `NO-RELEASE`，等待独立 Reviewer 再次验收。

### Round 4

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`，本轮只读，未修改。
- 执行开始：2026-08-25T08:01:00.942Z
- 执行完成：2026-08-25T08:13:38.1343799Z

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiRelationBinder.cs`
	- `src/Bing.Offices.Npoi/Internals/NpoiWorkbookPlanKeyBuilder.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingTextReader.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`
- 根因：Importer 的关系绑定、Importer/Exporter 重复 workbook-plan key 构造和 Loader 的受限文本读取边界与主流程耦合；Mapping cache key 还存在可避免的字符串中间结果。
- 修复：
	- 抽取 `NpoiRelationBinder`，保留原关系键解析、导航属性写入、错误定位、取消和最大错误数语义。
	- 抽取共享 `NpoiWorkbookPlanKeyBuilder`，统一 Import/Export 的类型、Document、方向配置分组键，避免两份重复实现。
	- 将 Loader 的受限文本读取和 UTF-8 大小校验移至 `ExcelMappingTextReader`，保持 JSON/XML 公共入口、1 MiB 限制、调用方流所有权和异常边界。
	- 将 `ExcelMappingPlanFactory` cache key 的 JSON 序列化改为 `SerializeToUtf8Bytes`，移除字符串到 UTF-8 的额外转换；cache key 字段和 SHA256/Base64 结果契约保持不变。
	- 保留 NPOI DOM 所需的 bounded input `MemoryStream`；没有将该路径误称为 streaming 或 zero-GC。
- 验证：
	- net8 专项 Unit：132/132 通过。
	- net8 全量 Unit：216/216 通过。
	- net6 全量 Unit：216/216 通过。
	- net8 Integration：12/12 通过。
	- net6 Integration：12/12 通过。
	- `dotnet build Bing.Offices.sln -c Release --no-restore`：通过，0 error；仅有既有弃用和目标框架支持警告。
	- Mapping cache-hit benchmark：固定 `launchCount=1`、`warmupCount=2`、`iterationCount=3` 完成；当前矩阵 Mean 约 249.6 us 至 1.281 ms，进程退出码 0。
	- Stream benchmark smoke：完成并记录 Import/Export DOM 路径基线；不宣称 streaming 或 zero-GC。
	- `git diff --check`：通过；仅有既有 CRLF/LF 转换提示，无空白错误。
	- 修改生产文件 VS Code diagnostics：无错误。
- 未完成范围：Importer 的 NPOI DOM 输入缓冲仍是 WorkbookFactory 的有界可重读输入约束；Exporter/Loader 仍有更深层的写入、解析和校验职责，未进行高风险大范围重构；尚无完整、可归因的 cache-key/Stream 前后性能对比，也未完成 Office/LibreOffice/WPS 互操作和缺失旧 runtime 的运行验证。任务继续保持 `PARTIAL` / `NO-RELEASE`。

### Round 4 汇总

- MUST_FIX：0
- 已完成：关系绑定 collaborator、共享 workbook-plan key collaborator、Loader 受限文本 collaborator、cache-key 直接 UTF-8 序列化及相关回归验证。
- PARTIAL：FIX-003；保留 NPOI DOM 输入约束、更深层职责拆分、完整可归因性能前后对比和外部互操作残余。
- BLOCKED：net5/net7/netcoreapp3.1 testhost runtime、Office/LibreOffice/WPS 环境仍不可用。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 216/216；net8/net6 Integration 各 12/12；专项 Unit 132/132；Mapping cache-hit benchmark、Stream benchmark smoke、Release build、diagnostics 和 diff-check 通过。
- 执行终态：`PARTIAL`；未自动执行 git commit、git push、tag、PR 或发布。
- 下一步：重新执行独立 Review，复核 FIX-003 的范围裁决和性能证据；`review.md` 保持 Reviewer 独立证据，不由 Executor 修改。

### Round 5

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`，本轮只读，未修改。
- 执行开始：2026-08-25T08:23:40.554Z

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiImportPlanBuilder.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExportPlanBuilder.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
	- `ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/progress.md`
	- `ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/performance-baseline.md`
	- `ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/execution.md`
- 修复：新增 `NpoiImportPlanBuilder` 和 `NpoiExportPlanBuilder`，分别隔离 workbook plan 分组、方向化 key 使用、泛型反射创建及 `TargetInvocationException` 异常解包；Importer/Exporter 主流程继续负责 NPOI 管线编排，未改变公共 API、MappingDirection、Sheet 计划映射、模板处理或取消语义。
- 直接行为覆盖：复用现有 Mapping cache、Workbook Request、关系绑定和 Stream import/export 回归，未新增生产 `InternalsVisibleTo` 或测试出口。
- 验证：
	- net8 核心专项 Unit：132/132 通过。
	- 新增 plan builder 在 `Bing.Offices.Npoi` net8 Release 下编译通过。
	- VS Code diagnostics：Round 5 修改的四个生产文件无错误。
- 性能证据：Round 4 的 cache key 直接 UTF-8 序列化已在 `performance-baseline.md` 标记为已实施；本轮没有伪造旧/新实现配对数字。现有 Mapping fixed baseline 和 Stream smoke 仍只作为结构证据，不作为发布性能阈值。
- 未完成范围：Importer/Exporter/Loader 更深层职责拆分、导入有界 `MemoryStream` 与 NPOI DOM 资源约束、完整可归因 cache-key/Stream 前后性能对比、旧 runtime 运行、Office/LibreOffice/WPS 互操作和本地包消费者链路仍未完成。
- 本轮未修改 `review.md`，未执行 commit、push、tag、PR、NuGet publish 或破坏性 Git 操作。

### Round 5 汇总

- MUST_FIX：0
- 已完成：Importer/Exporter workbook plan builder 边界拆分及状态文件同步。
- PARTIAL：FIX-003，因批准计划定义的更深层职责拆分和可归因性能证据仍未完成。
- BLOCKED：net5/net7/netcoreapp3.1 testhost runtime、Office/LibreOffice/WPS 环境仍不可用。
- FAILED：无。
- 回归验证：net8 核心专项 132/132；生产文件 diagnostics 通过；未修改 Reviewer 独立证据。
- 执行终态：`PARTIAL`；任务整体保持 `PARTIAL` / `NO-RELEASE`。
- 下一步：回到 `code-reviewer` 进行再次验收，不自动提交或发布。

### Round 6

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`，本轮只读，未修改。
- 执行开始：2026-08-25T08:44:41.862Z

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiWorkbookValidationPipeline.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
	- `ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/progress.md`
	- `ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/performance-baseline.md`
	- `ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/execution.md`
- 修复：新增 internal `NpoiWorkbookValidationPipeline`，将 Workbook 原生 Data Validation 的规则解析、显式列表/范围读取、日期/数值/文本比较和错误收集从 Importer 主类移出；Importer 仅保留行处理编排，未改变错误顺序、UnsupportedFeaturePolicy、StopOnFirstFailure、取消或 Unique journal 行为。pipeline 复用 Importer 的无状态单元格读取和文本规范化方法，未复制第二套实现。
- 性能对照：Mapping benchmark 新增同一 cache-key payload 的旧式 `JsonSerializer.Serialize` + `Encoding.UTF8.GetBytes` 与当前 `JsonSerializer.SerializeToUtf8Bytes` 对照方法。固定 `launchCount=1/warmupCount=2/iterationCount=3`，64 个参数组合执行完成；代表性结果为旧路径约 `5.89 KB`、`2.3 us`，新路径约 `2.52 KB`、`2.15 us`。Dry 单次迭代低于 100ms，结果不作为发布阈值。
- 验证：
	- net8 核心专项 Unit：157/157 通过。
	- net8 Unit：216/216 通过。
	- net6 Unit：216/216 通过。
	- net8 Integration：12/12 通过。
	- net6 Integration：12/12 通过。
	- cache-key 对照 benchmark：64/64 参数组合完成，退出码 0。
	- Benchmark 项目 Release build：通过；仅有既有 `IgnoreNullValues` 弃用警告。
	- VS Code diagnostics：Round 6 修改生产文件无错误。
	- `git diff --check`：无空白错误；CRLF/LF 为既有提示。
- 未完成范围：Importer 的行 materialization/配置验证、Exporter 的 cell/style/width/merge/chart 写入和 Loader 的 JSON/XML parser/validator 更深层拆分；完整 Stream import/export 前后对比、NPOI DOM 输入约束、旧 runtime、Office/LibreOffice/WPS 互操作和 no-release 条件仍未解除。
- 本轮未修改 `review.md`，未执行 commit、push、tag、PR、NuGet publish 或破坏性 Git 操作。

### Round 6 汇总

- MUST_FIX：0
- 已完成：Workbook validation pipeline 边界拆分；cache-key 同 payload、同 Job 旧/新分配对照；状态与性能证据同步。
- PARTIAL：FIX-003，因批准计划要求的更深层职责拆分和完整 Stream 性能对比仍未完成。
- BLOCKED：net5/net7/netcoreapp3.1 testhost runtime、Office/LibreOffice/WPS 环境仍不可用。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 216/216；net8/net6 Integration 各 12/12；核心专项 157/157；cache-key benchmark 64/64 参数组合；Release benchmark build、diagnostics 和 diff-check 通过。
- 执行终态：`PARTIAL`；任务整体保持 `PARTIAL` / `NO-RELEASE`。
- 下一步：回到 `code-reviewer` 进行再次验收，不自动提交或发布。

### Round 7

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`，本轮只读，未修改。
- 执行开始：2026-08-25T09:08:25.467Z
- 执行完成：2026-08-25T09:42:00.000Z

#### FIX-003

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：PARTIAL
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `src/Bing.Offices.Npoi/Imports/NpoiImportRowMaterializer.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
	- `src/Bing.Offices.Npoi/Exports/NpoiExportSheetWriter.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
	- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingDocumentValidator.cs`
	- `benchmarks/Bing.Offices.Benchmarks/StreamPipelineBenchmarks.cs`
- 根因：Round 6 后 Importer 的行物化/配置校验、Exporter 的 Sheet 写入和 Loader 的文档验证仍与主流程混合；Exporter 还保留了已迁移写入逻辑的死代码，Stream 只有当前路径基线。
- 修复：
	- 新增并接入 `NpoiImportRowMaterializer`，承接图片索引、空行判断、原始值校验、动态值转换、实体创建、Unique 事务和资源索引。
	- 新增并接入 `NpoiExportSheetWriter`，承接表头、单元格、模板策略、批注、样式、列宽、换行、合并和图表写入。
	- 新增并接入 `ExcelMappingDocumentValidator`，承接 JSON 字段白名单、XML 结构、配置业务规则和文档校验；Loader 保留文本读取、反序列化和入口编排。
	- 删除 Exporter 中已由 Sheet Writer 接管的旧图表、Cell、样式、批注、布局和合并实现，删除 Importer 未使用的旧 `NormalizeCellValue`，避免第二套实现路径。
	- 为 Stream benchmark 增加同负载的 `ImportWithLegacySourceBuffer` 与 `ExportWithLegacyDestinationBuffer` 对照。两者是 benchmark-only 的显式缓冲适配器，复用当前生产实现，不声称重建了历史产品实现。
- 性能证据：固定 `launchCount=1`、`warmupCount=2`、`iterationCount=3`，覆盖 1,000/10,000/100,000 行，共 24 个测量项。当前路径与 benchmark-only 缓冲对照的代表性 Mean/Allocated 如下：
	- Import 100,000 行：`2.147 s / 1623.00 MB`；对照 `2.160 s / 1625.39 MB`。
	- Export 100,000 行：`1.382 s / 726.52 MB`；对照 `1.402 s / 726.52 MB`。
	- Import 10,000 行：`291.83 ms / 163.65 MB`；对照 `323.06 ms / 163.89 MB`。
	- Export 10,000 行：`130.45 ms / 70.63 MB`；对照 `156.31 ms / 70.63 MB`。
	- Dry 结果仅作结构性/分配证据；NPOI DOM 的有界输入缓冲仍存在，不宣称 streaming 或 zero-GC。
- 验证：
	- Loader JSON/XML/Mapping Patch/Profile 定向 Unit：`36/36` 通过。
	- net8 Unit 全量：`216/216` 通过。
	- net6 Unit 全量：`216/216` 通过。
	- net8 Integration：`12/12` 通过。
	- net6 Integration：`12/12` 通过。
	- NPOI 生产项目 `net8/net7/net6/netcoreapp3.1` Release build：通过；无编译错误。
	- Benchmark 项目 Release build：通过；仅有既有弃用警告。
	- Stream benchmark：24/24 测量项完成，进程退出码为 0。
	- 解决方案 Release build：通过，0 error。
	- VS Code diagnostics：本轮修改的 Loader、Importer、Exporter、benchmark 文件均无错误。
	- `git diff --check`：通过；仅有既有 CRLF/LF 转换提示，无空白错误。
- 未完成范围：
	- FIX-003 所对应的批准计划仍包含更细的高影响路径优化、NPOI DOM 内存约束设计、Office/LibreOffice/WPS 互操作、缺失旧 runtime 的实际测试，以及最终 release-ready 判定；本轮不将这些条件伪造为完成。
	- 由于 Reviewer 的独立验收尚未进行，不能将 `review.md` 的 `NEEDS_FIX` 改为通过。
- 任务结论：继续保持 `PARTIAL` / `NO-RELEASE`。

### Round 7 汇总

- MUST_FIX：0
- 已完成：Importer 行物化边界、Exporter Sheet 写入边界、Loader 文档验证边界、Exporter 死代码清理、Stream 同负载缓冲对照证据。
- PARTIAL：FIX-003；任务整体仍未达到 release-ready。
- BLOCKED：net5/net7/netcoreapp3.1 Unit testhost runtime、Office/LibreOffice/WPS 互操作环境仍不可用。
- FAILED：无。
- 回归验证：net8/net6 Unit 各 `216/216`；net8/net6 Integration 各 `12/12`；Stream benchmark 24/24；NPOI/Benchmark/解决方案 Release build、diagnostics 和 diff-check 通过。
- 执行终态：`PARTIAL`；未自动执行 git commit、git push、tag、PR、NuGet publish 或破坏性 Git 操作。
- 下一步：回到 `code-reviewer` 进行再次独立验收。

## Review 修复记录

### Round 8

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/TASK-BING-OFFICES-STABILIZE-20260825-001/review.md`
- 执行状态：COMPLETED

#### FIX-005

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
	- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
- 根因：`NpoiWorkbookValidationPipeline.Validate(...)` 返回 `false` 时，Importer 仅在 `StopOnFirstFailure` 模式跳过行物化；`Continue` 模式仍会进入 raw validation、转换、实体赋值和 Unique 路径，改变了拆分前 Workbook 校验失败行的行为边界。
- 修复：Workbook 校验只要返回失败，均回滚当前已开启的 Unique 事务并跳过 row materializer；保留 Workbook pipeline 在 `Continue` 模式下对当前行的错误收集能力。同步修正 Importer 行循环缩进，未修改公共 API、校验来源选择、生产 IVT 或取消语义。
- 测试：新增 `Import_WorkbookValidationFailure_Continue_ShouldSkipMaterialization`，使用真实 XLSX，覆盖 `WorkbookRules` 与 `ConfiguredAndWorkbook`；该行同时违反显式列表规则且不能转换为 `int`，断言只产生 `WorkbookValidation`，无 `ValueConversion`，且不产生成功实体。更新四模式矩阵，明确 Workbook 失败优先于配置校验。
- 验证：
	- net8 相关 Unit（P0/Workbook/Stream/Loader/Mapping/Profile）：`183/183`，PASS。
	- net6 相关 Unit：`183/183`，PASS。
	- net8 Integration：`12/12`，PASS。
	- `dotnet build Bing.Offices.sln -c Release --no-restore`：PASS，0 error。
	- VS Code diagnostics：修改的 Importer 和 P0 测试文件无错误。
	- `git diff --check`：PASS，无空白错误；仅保留既有换行格式提示。

### Round 8 汇总

- MUST_FIX：FIX-005。
- 已完成：FIX-005。
- PARTIAL：无本轮纳入范围的 FIX。
- BLOCKED：无新增阻断；任务原有旧 runtime、Office 互操作和 release-ready 残余仍保持记录。
- FAILED：无。
- 回归验证：net8/net6 相关 Unit 各 `183/183`；net8 Integration `12/12`；Release build、diagnostics 和 diff-check 通过。
- 下一步：回到 `code-reviewer` 进行再次独立验收；不得修改 `review.md` 伪造通过。
