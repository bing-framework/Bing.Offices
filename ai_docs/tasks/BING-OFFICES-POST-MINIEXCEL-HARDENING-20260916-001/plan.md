# Bing.Offices MiniExcel 后续整改与发布收口计划

## 任务信息

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 目标分支：`feat/miniexcel-provider`
- 角色：Planner；本文件只规划，不实施业务代码。
- 任务边界：修复当前 MiniExcel Provider 的真实正确性、TFM 兼容性、程序集边界、性能证据、测试、注释和产物治理；不重新实现 Provider，不自动修改版本，不 commit/push/publish。

## 当前状态与判断

- 已有 `Bing.Offices.Abstractions`、`Core`、`Npoi`、`MiniExcel` 四层；MiniExcel 已接入真实 `SaveAs/SaveAsAsync`、`Query/QueryAsync`、多 Sheet、动态列、Mapping、Converter、Validation、Unique、Relations、Stream/File 和 fail-fast unsupported。
- `Core` 目标框架仍为 `netstandard2.0`，但新增 `ExcelXlsxZipPreflight` 尚未由当前基线证明可编译；必须先复现具体编译器错误，不能沿用“已完成”历史结论。
- `Core/AssemblyInfo.cs` 当前向 NPOI 和 MiniExcel 暴露生产 `InternalsVisibleTo`；本任务目标为生产程序集间 IVT 为 0，除非形成维护者批准的唯一架构例外。
- `MiniExcelValueAdapter` 当前对日期直接使用 `DateTime.FromOADate`、`DateTime.Parse`、`DateTimeOffset.Parse`，尚未证明与 Core/NPOI 的 `ExcelCellValue`、1900/1904 日期系统合同一致。
- `MiniExcelExcelExporter.EnumerateRows` 当前把 converter 上下文 `rowIndex` 固定为 0；固定列和动态列均需修复并直接测试。
- Benchmark 入口未固定 BenchmarkDotNet `ArtifactsPath`；Consumer 项目默认硬编码 Bing 包版本；已有工作树包含前一任务变更，执行器必须先保存本任务自己的版本和产物基线。
- 当前版本基线（只读记录，不得修改）：`VersionPrefix=2.0.0`；`MiniExcelPackageVersion=[1.46.0]`；`NpoiPackageVersion=[2.7.4]`。`version.dev.props` 当前已有工作树改动，必须以任务开始时快照为准判断本任务是否改动。

## 不变约束

- 保留 `Core netstandard2.0`；不得改为 net6/net8 绕过兼容错误。
- 继续对 XLS、Chart、Comment、Image、Style、Template、Merge、Failure Workbook 等未完整支持能力 fail-fast，不得静默忽略。
- 不删除或改签名现有公共成员；若确需新增 Provider SPI，必须记录 API diff、消费者影响和迁移方案，并更新双 TFM snapshot。
- 不修改 `version.props`、`version.dev.props`、任何 `Version*`/`PackageVersion` 属性；只允许读取和核验。
- 所有正式验证产物统一写入仓库根 `artifacts/`；禁止在 `src/**`、`tests/**`、`benchmarks/**` 留下本任务生成的 artifacts、TestResults、packages、BenchmarkDotNet.Artifacts 或 probe 输出。
- 生产代码禁止 `Task.Run`、`.Result`、`.Wait()` 伪装异步；保持调用方 source/destination stream 所有权。

## Phase 0：真实基线、版本冻结和产物扫描

### H-000 基线快照

- 目标：冻结任务起点的 Git、SDK、TFM、版本文件、生产 IVT 和产物分布。
- 现状/证据：工作树非 clean，已有 MiniExcel 任务改动；不能用当前 diff 推断本任务改动。
- 修改范围：仅新增本任务 `baseline.md` 或写入 `execution.md` 的基线章节；不改生产文件。
- 实施步骤：读取 `AGENTS.md`、所有 `*.props`/`*.csproj`、CI、`.gitignore`；记录 `git status --short`、`git diff --stat`、当前 commit/branch；保存 `version.props`、`version.dev.props`、所有 csproj 版本属性的 UTF-8 内容或 hash；列出生产 `InternalsVisibleTo`；扫描项目目录中的 artifacts。
- 依赖：无。
- 验证：基线文件包含 Core=`netstandard2.0`、NPOI/MiniExcel=`net6.0;net8.0`、Benchmark=`net8.0`、版本值、IVT 清单和扫描结果。
- 风险：把前一任务已有 `version.dev.props` 或 `.gitignore` diff 误归因于本任务。
- 验收标准：能区分任务前已有改动与本任务新增改动，且版本文件快照可在最终复核重比对。

### H-001 真实编译失败定位

- 目标：重现并准确记录 `ExcelXlsxZipPreflight` 在 Core/netstandard2.0 下的错误 API、行号、错误码和调用链。
- 现状/证据：实现使用 `ZipArchive`、`ZipArchiveEntry`、字符串比较和 XML API，历史报告未提供 netstandard2.0 编译证据。
- 修改范围：仅基线/日志报告；不先修代码。
- 实施步骤：使用隔离输出目录执行 Core 和 solution Release build；从编译器输出定位具体 API/overload/type/member，分类为可用、兼容替代、无等价、需条件编译或应移到 Provider。
- 依赖：H-000。
- 验证：保存完整命令、退出码和错误文本；不得以文档声明替代编译结果。
- 风险：构建缓存或 NuGet 环境造成假阴性；需明确 `--no-restore` 与 restore 状态。
- 验收标准：报告包含至少一个可复现的真实错误或明确证明当前错误已不存在，并说明验证环境。

## Phase 1：Preflight 所有权和 netstandard2.0 兼容修复

### H-100 兼容方案决策

- 目标：在不升 Core TFM、不削弱安全合同的前提下确定唯一实现路径。
- 现状/证据：当前共享算法位于 Core，NPOI/MiniExcel 通过 IVT 调用 internal helper。
- 修改范围：`ExcelXlsxZipPreflight`、相关 Provider 适配器或最小内部 SPI；不新增大杂烩项目。
- 实施步骤：优先把不兼容 API 替换为 netstandard2.0 等价 API；若不存在可靠等价，保留 Core 的 limits/path/ratio/XML policy，把 Zip/XML runtime adapter 移到各 Provider，并以最小 Provider SPI 传递策略/结果；禁止复制两套完整算法。
- 依赖：H-001。
- 验证：Core netstandard2.0 build；NPOI/MiniExcel 双 TFM build；预检取消、流位置复位、DTD、路径穿越、重复 entry、entry/总量/压缩比/XML budget 测试。
- 风险：条件编译扩大分支或改变异常 provider/stage；必须保持原安全语义和异常类型。
- 验收标准：Core netstandard2.0 编译通过；两个 Provider 均执行相同安全 policy；所有预检回归通过。

### H-101 生产 IVT 清理

- 目标：消除 Core→Provider 生产友元依赖。
- 现状/证据：Core 当前有 `InternalsVisibleTo("Bing.Offices.Npoi")` 和 `Bing.Offices.MiniExcel`。
- 修改范围：Core/NPOI/MiniExcel `AssemblyInfo.cs` 及必要的最小 Provider SPI 所有权文件。
- 实施步骤：优先移动 owning implementation；若必须跨程序集调用，只暴露少量明确的 `Providers`/`Infrastructure` SPI，使用 `EditorBrowsable(Never)` 并记录契约；不得把几十个 internal 批量改 public；测试友元按需保留。
- 依赖：H-100。
- 验证：Public API/程序集属性测试明确断言生产 IVT=0；双 TFM API snapshot 和 package consumer 通过。
- 风险：过度扩大 public surface 或破坏测试访问；新增 SPI 必须有 api-diff 和迁移记录。
- 验收标准：生产程序集间 IVT 为 0，或存在维护者明确批准且限界、移除计划清晰的唯一例外。

## Phase 2：日期合同和 RowIndex 正确性

### H-200 MiniExcel 日期适配

- 目标：让 MiniExcel 日期导入走 Core 的 `ExcelCellValue`/`ExcelDateParser` 合同，覆盖 DateTime、DateTimeOffset、Excel serial、Culture、ExcelDateAttribute 和 1900/1904。
- 现状/证据：适配器直接 Parse/FromOADate，`CreateCell` 未携带 `IsDate1904`；MiniExcel API 是否暴露 workbook 1904 标志待 Spike。
- 修改范围：`MiniExcelValueAdapter`、必要的 MiniExcel raw-value/worksheet metadata 适配、日期测试 fixture。
- 实施步骤：先用真实 MiniExcel 读取 1900/1904 fixture，确认 API 可获取的日期系统信息；可获取则填入 `ExcelCellValue.IsDate1904`，不可获取则明确能力缺口并对受影响场景 fail-fast/标记不支持；禁止再造第二套日期解析规则。
- 依赖：H-100；先完成真实 API Spike。
- 验证：DateTime/DateTimeOffset、serial、culture、nullable、ExcelDateAttribute、1900/1904 与 NPOI 同一 request 的行为等价测试。
- 风险：第三方只返回已转换 DateTime，无法恢复原始 serial 或 date system；不得误称 parity。
- 验收标准：共同支持能力日期结果与 NPOI 一致；无法支持的 1904 场景有结构化边界和文档。

### H-201 Converter RowIndex 修复

- 目标：固定列和动态列的 converter 上下文行号与公共合同一致。
- 现状/证据：Exporter 每次枚举把 `rowNumber` 初始化为 0；Importer 已按物理行推进，需保留回归。
- 修改范围：`MiniExcelExcelExporter.EnumerateRows`、`MiniExcelValueAdapter` 相关直接测试。
- 实施步骤：按表头/数据行约定计算 1-based 或既有 contract 的行号；固定列、动态列、跳过行和多 Sheet 分别验证；不在热路径重复做反射元数据准备。
- 依赖：H-200 可并行，依赖公共行号合同确认。
- 验证：converter 根据 `context.RowIndex` 产生 Row1/Row2/Row3 不同输出的直接测试，双 TFM 执行。
- 风险：把零基物理索引和公共错误 RowIndex 混淆；以既有 NPOI/异常测试合同为准。
- 验收标准：固定/动态列 converter 行号真实递增，错误上下文和 NPOI 保持一致。

## Phase 3：异步、取消、流所有权和资源边界

### H-300 真实异步与取消矩阵

- 目标：验证中途取消、资源限制取消、源/目标流保持打开、临时文件清理和旧目标保护。
- 现状/证据：已有 pre-cancel 和基础 async 测试，但缺少 mid-flight 真实 IO 矩阵。
- 修改范围：MiniExcel unit/integration 测试、必要的 staging/commit 修复；不改变公共签名。
- 实施步骤：使用可控异步 source/enumerable/stream 在复制、预检、行枚举、SaveAsAsync/QueryAsync 各边界注入取消；验证 `OperationCanceledException` 原样传播、destination/source 未关闭、失败不替换旧文件、临时文件删除。
- 依赖：H-100、H-201。
- 验证：双 TFM unit/integration；资源矩阵输出到 `artifacts/resource-probe/`。
- 风险：MiniExcel v1 async 仍可能在返回后同步枚举；报告必须区分真实异步 IO 和异步包装。
- 验收标准：pre-cancel、mid-flight export/import、resource-limit cancellation、stream ownership、temp cleanup 全部有直接测试证据。

## Phase 4：Hot Path、GC 和关系算法

### H-400 性能热点审计

- 目标：在不改变行为的前提下降低明显分配或 CPU 热点。
- 现状/证据：已知热点包括反射 `GetProperty/GetValue/SetValue`、`Activator.CreateInstance`、`DynamicInvoke`、`ToDictionary`、重复 HashSet 和 MemoryStream；100K probe 曾显示约 1.28GB allocation。
- 修改范围：MiniExcel plan/row projection/import materialization；只做有 benchmark 前后证据的最小优化。
- 实施步骤：先建立相同 workload 的 NPOI vs MiniExcel Export/Import/Roundtrip 基线；优先缓存 property metadata、减少每行 schema/HashSet 重建、控制临时数组/字典；不得为了理论收益引入未经证明的池化或复杂并发。
- 依赖：H-201；需先完成 H-800 基线设计。
- 验证：MemoryDiagnoser、elapsed/allocated/GC/working set/rows/sec/output size；100K 为门禁，500K/1M 能跑则执行，否则 `NOT_VERIFIED`。
- 风险：优化改变映射、错误顺序、动态列或流生命周期；每个热点必须有行为回归。
- 验收标准：有真实前后数据；若无稳定改善，保留简单实现并在报告说明。

### H-401 关系复杂度确认

- 目标：确认关系绑定复杂度和容量边界，不无证据重写算法。
- 现状/证据：当前 MiniExcel relation binder 使用 parents 数组和 `FirstOrDefault` 查找，最坏为 O(P×C)。
- 修改范围：关系实现或 benchmark/report；保持 comparer、空键、重复键、未匹配和错误上下文语义。
- 实施步骤：用相同 request 测量当前复杂度；只有 benchmark 证明必要时按 comparer 建索引，补充重复/未匹配/失败后状态测试。
- 依赖：H-400。
- 验证：relation unit、cross-provider contract 和性能样本。
- 风险：哈希索引改变 comparer 或重复键语义。
- 验收标准：final report 明确复杂度；任何优化均有行为等价证据。

## Phase 5：专项测试与跨 Provider 合同

### H-500 MiniExcel 职责矩阵

- 目标：为每个生产符号建立直接测试，不以全项目总数替代。
- 现状/证据：已有 MiniExcel 基础 10 项测试，但缺少日期、RowIndex、中途取消、流所有权和完整 unsupported 矩阵。
- 修改范围：`tests/Bing.Offices.Tests/MiniExcelProviderTest.cs`、Integration、必要 fixture；测试项目保持双 TFM。
- 实施步骤：覆盖 DateTime/DateTimeOffset/1900/1904、固定/动态 converter RowIndex、多 Sheet、ValueMap、Validation、Unique、Relations、Async、Cancellation、resource limit、unsupported 零输出、DI、file commit。
- 依赖：H-200、H-201、H-300。
- 验证：net6/net8 MiniExcel filter、full unit/integration；输出 `unit-test-report.md`、`integration-test-report.md`。
- 风险：测试只断言片段或间接行为；必须断言完整 workbook、结构化错误、目标文件和流状态。
- 验收标准：`symbol-test-map.md` 将 Exporter、Importer、ValueAdapter、Preflight、DI、FileCommitter 映射到具体测试方法。

### H-501 Cross-provider contract

- 目标：以同一 request/model 比较 NPOI 与 MiniExcel 的共同能力行为，而非二进制文件相同。
- 现状/证据：当前交叉证据主要是各自测试，未形成完整日期/类型/错误合同矩阵。
- 修改范围：现有测试项目新增 contract fixture/测试；不新增运行时 resolver。
- 实施步骤：比较 Header、列序、Date、Decimal、Nullable、Enum、Bool、Dynamic Column、ValueMap、Converter、Validation、Relation、Error contract；对 unsupported 能力分别断言 MiniExcel fail-fast。
- 依赖：H-200、H-201、H-500。
- 验证：双 TFM contract tests，记录差异和明确允许的 Provider-specific 行为。
- 风险：把 NPOI 特有功能误纳入共同合同。
- 验收标准：共同能力行为等价；差异均出现在 capability matrix 和 migration notes。

## Phase 6：Artifacts、Benchmark、Consumer 和 CI 治理

### H-600 统一产物目录

- 目标：从源头将 Benchmark、test result、pack、consumer、resource、API 日志写入根 `artifacts/`。
- 现状/证据：CI 已部分使用 `artifacts/`，BenchmarkDotNet 入口尚未设置 `ArtifactsPath`；项目目录可能残留历史输出。
- 修改范围：BenchmarkDotNet `ManualConfig/ArtifactsPath`、CI 命令和 `.gitignore`；不重写 bin/obj 基础布局。
- 实施步骤：配置 `artifacts/benchmarks/`；测试命令使用 `--results-directory artifacts/tests/...`；pack 使用 `artifacts/packages`；consumer 临时项目/cache 使用 `artifacts/consumers`；resource/API 日志归档到对应子目录；扫描并报告历史残留，只有本任务生成的残留才清理。
- 依赖：H-000。
- 验证：扫描 `src/**/artifacts`、`tests/**/artifacts`、`benchmarks/**/BenchmarkDotNet.Artifacts`、项目目录 TestResults/packages；目标计数为 0。
- 风险：误删用户已有产物；只处理已确认由本任务生成且范围明确的临时文件。
- 验收标准：所有新验证输出集中根 `artifacts/`，正式 Markdown 证据保留在任务目录；`.gitignore` 覆盖 `/artifacts/`。

### H-601 Consumer/版本发现治理

- 目标：Consumer 消费本次 pack 的实际包，不依赖硬编码 `2.0.0`。
- 现状/证据：两个 Consumer 的 `BingOfficesPackageVersion` 默认值为 `2.0.0`。
- 修改范围：Consumer csproj、CI pack/restore 命令、package report；禁止修改版本文件。
- 实施步骤：从 `artifacts/packages` 唯一匹配并读取实际 nupkg/nuspec 版本，或通过 CI 属性注入当前维护者版本；保留 PackageReference 而非 ProjectReference；验证 MiniExcel 无 NPOI 依赖。
- 依赖：H-600。
- 验证：net6/net8 consumer build/run，输出 `package-consumer-report.md`；版本文件 hash 不变。
- 风险：多版本包残留导致错误选择；每次验证前清理/隔离 `artifacts/packages` 中的目标包集合，不触及用户文件。
- 验收标准：两 TFM 消费实际 pack 包成功，无硬编码新版本策略，无版本文件 diff。

### H-602 API Snapshot 与 CI 门禁

- 目标：使 API 变化、包内容和版本治理独立可审计。
- 现状/证据：已有双 TFM snapshot；baseline 的自动 additive 标记不能替代维护者审批。
- 修改范围：API snapshot 工具、baseline/报告、CI 验证步骤；不自动批准 baseline。
- 实施步骤：若 SPI/public API 改变，生成 `api-diff.md`（Added/Removed/Changed/Reason/Impact/Migration），等待维护者审批；CI 加入版本文件未改、生产 IVT、产物目录、MiniExcel nuspec/NPOI 依赖检查。
- 依赖：H-101、H-601。
- 验证：Release build、双 TFM API capture/compare、identity self-test、pack 内容、`git diff --check`。
- 风险：把 XML 注释变化误判为签名 breaking，或用自动 marker 伪造审批。
- 验收标准：所有 API 变化可追溯，未批准 baseline 不标记发布通过。

## Phase 7：中文注释治理与文档同步

### H-700 应用 chinese-comments

- 目标：按项目既有技能检查本任务实际修改的 C# 文件，补充准确中文 XML 文档，不改变行为和签名。
- 现状/证据：部分新类有 summary，但复杂 internal 算法、异步/取消、资源/安全和日期语义注释需要按真实行为复核。
- 修改范围：本任务实际修改的 `ExcelXlsxZipPreflight`、Provider SPI、MiniExcel adapter/exporter/importer、相关 Options/Exceptions；不批量扫描全仓库。
- 实施步骤：先完成正确性实现和测试，再按 `chinese-comments` 规则检查 public/protected/SPI/复杂 internal、参数、返回值、异常和兼容性说明；实现成员优先 `<inheritdoc />`，不写逐行或无信息注释；检查 XML 可解析和引用合法。
- 依赖：H-100、H-200、H-300、H-500。
- 验证：Docs tests、XML documentation build/package inspection、人工抽查注释与实际 async/unsupported/TFM 行为一致。
- 风险：复制旧注释造成错误承诺；不能宣称完整 streaming、1904 支持或性能保证，除非代码和测试证明。
- 验收标准：注释覆盖本任务修改范围，未引入签名/行为变化；未覆盖的全仓库缺口列为 follow-up。

### H-701 文档与发布说明

- 目标：同步 README、Excel provider 文档和能力矩阵，准确区分支持、unsupported、NOT_VERIFIED 和发布阻塞。
- 现状/证据：已有 MiniExcel provider 文档和前一任务报告，部分性能/审批限制仍需本任务结果覆盖。
- 修改范围：`README.md`、`docs/excel/*`、本任务报告。
- 实施步骤：更新日期合同、1904 边界、异步/流语义、性能数据、artifact 路径、API/迁移说明和版本建议；不修改版本配置。
- 依赖：H-200、H-400、H-600、H-700。
- 验证：Docs tests、文档内容审查、报告与真实命令/数据逐项对照。
- 风险：把未跑的 500K/1M、生产 2C4G 或 GitHub Actions 写成 PASS。
- 验收标准：文档只引用真实证据，未验证项统一 `NOT_VERIFIED`。

## Phase 8：最终验证、报告和发布判定

### H-800 发布收口

- 目标：生成完整任务证据并给出是否达到可发布状态的明确结论。
- 现状/证据：前一任务已有 execution/review/report，但本任务必须重新基于实际 diff 和验证结果判断。
- 修改范围：`baseline.md`、`execution.md`、`review.md`、`api-diff.md`、`unit-test-report.md`、`integration-test-report.md`、`package-consumer-report.md`、`benchmark-report.md`、`resource-report.md`、`final-report.md`。
- 实施步骤：执行当前 csproj/CI 解析出的 Release build、net6/net8 unit/integration、docs、API snapshot、package consumer、resource matrix、100K benchmark；500K/1M 可运行则执行，否则 `NOT_VERIFIED`；扫描版本、IVT、artifacts、diff；由独立 reviewer 对 plan/execution/diff/证据审查。
- 依赖：H-100 至 H-701 全部完成。
- 验证命令：以当前仓库真实配置为准，至少包括 `dotnet build Bing.Offices.sln -c Release`、双 TFM测试项目、Docs、API Snapshot identity/capture/compare、四包 pack 到 `artifacts/packages`、net6/net8 consumer、Benchmark list/probe、资源矩阵和 `git diff --check`；测试结果用 `--results-directory artifacts/tests/...`。
- 风险：环境限制、NuGet SSL、机器容量或未获 API 审批；必须保留原始退出码和限制，不降低状态标准。
- 验收标准：仅当 Core netstandard2.0、生产 IVT、日期合同、RowIndex、async/resource/stream、专项和 cross-provider 测试、100K benchmark、artifact 治理、中文注释、API/consumer、Release build、diff check 全部满足，且版本文件未变、未 commit/push/publish，才能标记 `COMPLETED`；否则明确 `NEEDS_FIX` 或 `NOT_VERIFIED`。

## 最终报告必须回答

1. `ExcelXlsxZipPreflight` 在 netstandard2.0 失败的具体 API、最终兼容方案及安全语义是否保持。
2. 生产 IVT 是否为 0；若非 0，批准人、范围和移除计划是什么。
3. MiniExcel/NPOI 日期合同、Date1904 能力和 Converter RowIndex 是否一致。
4. 100K 优化前后 elapsed/allocated/GC/working set/rows/sec/output size；Relation 复杂度。
5. 所有 artifacts 的最终路径及项目目录残留扫描结果。
6. chinese-comments 实际处理文件、仍需后续治理的范围。
7. API 变化、消费者迁移、发布阻塞项和当前是否真正可发布。
8. `version.props`、`version.dev.props`、DLL/NuGet version 是否均未被 Agent 修改。

## 计划级 Definition of Done

- 已生成本计划；执行阶段不得在未完成上述证据前宣称完成。
- Core 继续 `netstandard2.0` 且真实 build 通过，或完成有证据的 ownership 重构。
- 生产 IVT 清理、日期合同、RowIndex、异步/资源/流语义和专项 cross-provider 测试均有直接证据。
- Benchmark 输出、test results、packages、consumer、resource/API 日志集中在根 `artifacts/`；100K 完成，500K/1M 明确 PASS 或 `NOT_VERIFIED`。
- 中文注释遵守 `chinese-comments`，文档与真实能力一致；双 TFM API、PackageReference consumer、Release build、`git diff --check` 通过。
- 版本文件未改变；不自动 commit、push、tag、PR、NuGet publish。
