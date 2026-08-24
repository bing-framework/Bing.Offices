# TASK-BING-OFFICES-20260821-MAPVAL-V2 实施计划

> Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`  
> Task Name：Bing.Offices Mapping Profile、校验体系、JSON/XML 与 NuGet 边界改造  
> 计划日期：2026-08-21  
> 状态：`TODO`（供 `/execute-plan` 执行；本文件不是实施完成报告）  
> Compatibility Mode：`MIGRATION_CURRENT_MAJOR`

## 1. 范围与决策

### 1.1 已确认仓库事实

| 项目 | 目标框架 | 包/依赖关系 | 证据 |
| --- | --- | --- | --- |
| `Bing.Offices.Abstractions` | `netstandard2.0` | 包身份保持不变 | `src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj` |
| `Bing.Offices.Core` | `netstandard2.0` | 引用 Abstractions；CsvHelper `33.1.0` | `src/Bing.Offices.Core/Bing.Offices.Core.csproj` |
| `Bing.Offices.Npoi` | `net8.0;net7.0;net6.0;netcoreapp3.1` | 引用 Core；NPOI `[2.7.4]` | `src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj`、`version.dev.props` |
| 版本 | `1.0.0` | 无 next-major 标识 | `version.props` |
| 测试 | Unit 多目标；Integration `net6.0;net8.0`；Docs Consumer `net8.0` | xUnit，现有真实 NPOI 集成测试 | `tests/*/*.csproj` |
| Benchmark | `net8.0` | BenchmarkDotNet `0.14.0` | `benchmarks/Bing.Offices.Benchmarks` |

当前引用 `refs/heads/master`，静态 HEAD 为 `73883f709f8eb9e58cd948db0bf90e82ca44a661`。本 Planner 只能进行只读仓库分析，未运行 `git status`、restore、build、test、pack、benchmark；执行开始时必须重新记录分支、HEAD、`git status --short --branch`、当前差异和命令输出。不得把历史 TRX、已有 artifacts 或相邻任务计划的结果当作本 Task 的基线。

选择 `MIGRATION_CURRENT_MAJOR` 的依据是当前版本 `1.0.0`、三个已发布包身份和用户明确“当前 major 保持兼容”的约束。除非执行 Phase 0 发现明确的 next-major 分支/发布策略并在 `03-decisions.md` 建 ADR，本任务不得删除已发布 API、改变 `AddNpoi()` 返回类型，或建立第二条永久生产执行链。

### 1.2 指定输入文档的可用性与冲突

- 未找到 `ai_docs/Bing.Offices-validation-mapping-refactor-solution-v2-20260821.md`、`Bing.Offices-implementation-review-20260821.md`，也未找到同名变体。执行 Phase 0 必须再次搜索并在 `03-decisions.md` 记录实际路径或缺失结论。
- 已存在相邻计划 `ai_docs/tasks/bing-offices-excel-import-export-enhancement-v2/plan.md`。其中部分工程事实可作为候选线索，但它以“未发布 API 可破坏收敛”为前提，和本 Task 的兼容迁移要求冲突；本计划的兼容策略优先。
- 当前源码已经实现 Workbook Request、动态列基本读写、读取范围、独立 Header/Body 空白策略、图片导入/失败工作簿/部分 Workbook Validation、模板和样式等能力。不得仅依据相邻计划的历史“未实现”判断覆盖当前行为。

### 1.3 强制架构目标

1. 保留三个 NuGet 包及依赖方向 `Abstractions <- Core <- Npoi`。
2. `Abstractions` 不反向依赖 Core/Npoi；普通公共 API 不暴露 NPOI 类型。
3. Attribute、Profile、JSON/XML、Request Fluent 均被规范化、按固定优先级合并并编译为一套不可变 Plan；Excel 与 CSV 使用该 Plan。
4. 优先级固定为 `Convention/Default < Attribute < Mapping Profile Fluent < JSON/XML Mapping < Request Fluent`；Import 和 Export 独立合并。
5. 移除全部生产程序集友元关系；跨程序集只通过 `Bing.Offices.Providers` 下最小、只读、隐藏于 IntelliSense 的 SPI 通信。
6. 所有执行和证据文档必须包含完整 Task ID；执行期间不执行 Commit、Push、PR、Tag 或 NuGet 发布。

## 2. 当前实现、缺口与完成度

### 2.1 已经真实进入生产链的能力

- `ExcelTypeMapFactory` 已按 Request Configuration > 单模型 `ExcelMappingProfile<T>` > Attribute/Convention 的局部顺序覆写固定列；Profile 会深拷贝部分字段。
- JSON 与 XML 均能读取平铺 `ExcelMappingConfiguration`，XML 已使用 `DtdProcessing.Prohibit` 和 `XmlResolver = null`。
- `NpoiExcelImporter` 与 `CsvEntityImporter` 均支持 Attribute、命名验证器、转换器和基础映射；NPOI 已有 internal `ExcelColumnPlan`。
- 当前 Excel Request 已具有动态列 Key/Alias、转换器、读取范围、空白策略、图片策略、资源限制、失败输出和部分 Workbook Validation；这证明之前的 Workbook Request 改造不是骨架。
- Unit、Integration、Docs Consumer、Public API 快照和 Benchmark 项目均存在；CI 真实运行 locked restore、Release build、net6/net8 Unit/Integration 及三个本地 pack。

### 2.2 未达到本 Task 目标的关键缺口

| 领域 | 当前证据 | 结论 |
| --- | --- | --- |
| Mapping Profile | 只有 sealed `ExcelMappingProfile<T>`，没有双模型/方向 Builder/Profile Registry/程序集扫描或显式注册 | 未实现 |
| 统一编译 | 仅 NPOI 有 `ExcelColumnPlan`；Core CSV 仍直接消费 `ExcelTypeMap` 并各自执行转换/校验 | 需要重构 |
| 配置合并 | 没有 `MappingSourceKind`、三态 Ignore、Rule Key Remove/Clear 或集合 Replace/Append 契约 | 未实现 |
| 不可变性 | 旧 Profile Clone 漏掉已存在的扩展字段风险；外部 `ExcelMappingConfiguration` 及动态定义需要统一 snapshot | 部分完成 |
| 校验 | Required/Regex/Range/MaxLength/DateTime/Duplication 是 Attribute Rule；Regex 无 timeout；无 MaxValue；Range 仅 decimal/int 路径；上下文公开可变 DuplicateValues | 需要重构 |
| Unique 性能 | CSV 每行调用 `CloneDuplicateValues`，复制全量集合，接近 O(n^2) 分配；NPOI 是局部 rollback，但未实现统一资源上限/descriptor | P0 性能问题 |
| JSON/XML | 未有 normalized `ExcelMappingDocument`、v2、version/model alias、限制、精确路径、未知字段策略或 v1 迁移诊断 | 未实现 |
| DI | 只有 `AddNpoi(): void`；无 `AddBingOffices`、Profile 注册或重复 Registry 检测 | 未实现 |
| NuGet 边界 | `Abstractions/AssemblyInfo.cs` 对 Core/Npoi 有生产 IVT；Npoi 仅有测试 IVT | 未完成 |
| 可见性 | Core 将 `ExpressionExtension`、`PropertyInfoExtensions`、`TypeExtensions` 等内部 helper 列入公共 API；NPOI low-level extensions 候选需审计 | 需要收敛 |
| 性能证据 | Benchmark 仅 1K Excel Import/Export；没有映射、校验、缓存、JSON/XML 或 GC 矩阵 | 未完成 |

### 2.3 完成度与质量判断

本 Task 的核心目标（统一 Plan、方向 Profile、v2 JSON/XML、六类校验、SPI/IVT、兼容迁移）目前约为 **25%**：底层 Workbook/CSV、局部映射和若干输入能力已存在，但中心编译模型、公共契约、注册和安全配置格式均缺失。该数字只衡量本 Task，不否定现有 Excel 导入导出功能。

主要性能风险是 CSV Unique 的全量状态复制、每个执行器独立的绑定/验证路径和配置加载无输入配额。主要维护性风险是 `NpoiExcelImporter`、`NpoiExcelExporter`、`CsvEntityPipeline`、`ExcelTypeMapFactory` 与 `ExcelColumnPlan` 的职责重叠，以及公共 API 将实现辅助类型暴露给消费者。已有 `PublicApiContractTest` 会阻止未经批准的公开面变化，但其基线必须在迁移中改造成 consumer/provider 两套 approval，不能仅更新哈希掩盖破坏。

## 3. 执行协议与产物

执行 Agent 的首个写入动作必须在 `ai_docs/tasks/TASK-BING-OFFICES-20260821-MAPVAL-V2/` 建立以下文件，统一状态值只能为 `TODO`、`IN_PROGRESS`、`BLOCKED`、`DONE`；任一时间只能一个主 Phase 为 `IN_PROGRESS`：

| 文件 | 责任 |
| --- | --- |
| `00-baseline.md` | Git、版本、TFM、包、实际命令与当前测试/基准结果 |
| `01-plan.md` | 本计划的执行镜像、阶段状态与依赖 |
| `02-progress.md` | 每轮修改、失败、下一步；上下文切换前先更新 |
| `03-decisions.md` | ADR、源码证据、API/兼容取舍 |
| `04-compatibility.md` | 三包、旧 API、JSON/XML v1/v2 矩阵 |
| `05-verification.md` | Build、Unit、Integration、Docs Consumer、API、包/互操作证据 |
| `06-performance-gc.md` | Benchmark、分配、GC、LOH、Working Set |
| `07-final-report.md` | 完成项、风险、迁移、未执行发布声明 |

不得改写本 `plan.md` 来伪造执行状态；实施记录写入上述执行文档。所有新增文本文件必须 UTF-8。

## 4. 分阶段任务

### Phase 0 - 基线与设计冻结（P0）

#### MAPVAL-000 基线、状态文档与兼容模式确认

- **依赖**：无。**状态**：`TODO`。
- **目标**：创建第 3 节的状态文件，确认当前分支/HEAD/工作区改动、包版本、TFM、Public API 和实际验证状态；不覆盖用户改动。
- **已确认范围**：解决方案、三个 csproj、`version*.props`、`common.props`、CI、各 test/benchmark csproj、`PublicApiContractTest.cs`。
- **步骤**：运行 `git status --short --branch`、`git rev-parse HEAD`、`git diff --stat`；搜索缺失方案/评审文档；记录所有与本 Task 重叠的未提交文件；读取 package lock；将版本模式判定写入 `00-baseline.md` 和 `04-compatibility.md`。
- **验证**：先执行 CI 的 locked restore 和 Release build；其后按第 7 节完整记录 Unit、Integration、Docs Consumer、pack 和现有 Benchmark smoke。任何原有失败保留原始输出与归属。
- **验收**：存在八个带 Task ID 的证据文件；兼容模式有源码/版本证据；没有实施代码变更混入基线。

#### MAPVAL-001 API、调用链和数据格式盘点

- **依赖**：MAPVAL-000。**状态**：`TODO`。
- **目标**：形成 Profile/Mapping/Validation/JSON/XML/DI/Provider API 的可追溯清单，防止错误删除已发布符号。
- **步骤**：搜索仓库消费者、XML 文档、Docs Consumer 与测试；检查三个程序集的 `InternalsVisibleToAttribute`；采集 JSON/XML v1 fixture；建立“生产符号 -> 测试方法”表；对相邻旧计划与当前源码差异建立 ADR。
- **API 决策**：保留 `IExcelMappingConfigurationLoader` 并通过新增 document/load/diagnostic 重载演进；保留 `AddNpoi(): void`；确认是否存在真实 `AddBingOffices`，不存在则仅在需要的 owning package 添加 additive 注册入口。
- **验收**：`03-decisions.md` 明确已有、候选废弃和必须保留的 API；`04-compatibility.md` 有三包与 v1 基线。

### Phase 1 - Normalized Mapping 与不可变 Plan（P0）

#### MAPVAL-100 映射描述符、来源合并和不可变编译器

- **依赖**：MAPVAL-001。**状态**：`TODO`。
- **目标**：在 Abstractions 增加最小的消费者契约，在 Core 实现 internal Mapping Compiler；输出 immutable `WorkbookPlan`、`SheetPlan`、`ColumnPlan` 和 internal Rule/Converter bindings。
- **修改范围**：`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/`、`Providers/`（仅 SPI）、`src/Bing.Offices.Core/Bing/Offices/Mappings/`、现有 Mapping Builder/请求对象；按职责新增一类型一文件，禁止 `Models.cs` 等大杂烩。
- **步骤**：
  1. 引入显式 `MappingSourceKind` 与每个来源的不可变 snapshot；Import/Export 独立 merge。
  2. 用 PropertyName 识别固定列、Key 识别动态列；实现标量高优先级覆盖、三态 Ignore、Validation Rule Key 合并、`RemoveValidation(key)`、`ClearValidations()`，以及 Alias/ValueMap 的 `Replace`/`Append` 显式策略。
  3. Attribute、旧 `ExcelMappingProfile<T>`、旧 `ExcelMappingConfiguration` 和 Request Fluent 适配进 compiler；旧 API 只委托新 compiler。
  4. Plan 构建时复制 List/Array/alias/value map/profile 状态，预编译 Getter/Setter、Regex、Converter 和 Named Validator；不得持有外部可变集合。
  5. tenant cache key 必须包含租户/模型/方向/配置版本，设置容量和淘汰，缓存值只读。
- **测试矩阵**：五层优先级；Include 覆盖低层 Ignore；Remove/Clear；Replace/Append；固定/动态一致性；构建后修改输入；多线程/多租户缓存命中未命中与隔离；ImageMultiplicity 不丢失。
- **风险**：netstandard2.0 不可使用高版本 API；以条件编译或兼容实现保持包 TFM。
- **验收**：Excel/CSV 只接收同一 compiler 的方向 Plan；原始配置变动不能改变已构建结果。

#### MAPVAL-101 将 Excel、CSV 与固定/动态列接入单一 Plan

- **依赖**：MAPVAL-100。**状态**：`TODO`。
- **目标**：删除业务语义重复，而非强行共用 NPOI I/O；Core 负责 provider-neutral Plan 执行，Npoi 只做 Cell/Workbook 适配。
- **修改范围**：`CsvEntityPipeline.cs`、`ExcelTypeMapFactory.cs`、Npoi `ExcelColumnPlan.cs`、Importer/Exporter 的 plan 创建路径及关联测试。
- **步骤**：将 CSV 的 header binding、conversion、validation、dynamic dictionary key 使用 compiled Plan；NPOI 以只读 Provider SPI 消费 Plan；移除或 internal 化旧 TypeMap 作为 compatibility adapter；在计划阶段绑定命名 Validator/Converter，禁止逐 Cell 枚举 DI 集合。
- **测试**：同一 Mapping 同时驱动 CSV/XLSX；不同动态列 Key 与固定 PropertyName；Converter/Validator 只绑定一次；失败行不 Setter；并发复用 Plan。
- **验收**：没有并行的 converter/rule-name lookup；Excel/CSV 的等价规则结果可由直接测试证明。

### Phase 2 - Validation V2 与 Unique 资源治理（P0）

#### MAPVAL-200 Rule Descriptor、Attribute 迁移和执行顺序

- **依赖**：MAPVAL-101。**状态**：`TODO`。
- **目标**：实现 Required、Regex、Date、MaxValue、Range、Unique，并保留独立 MaxLength；所有规则由 Plan Descriptor 执行。
- **修改范围**：Abstractions Validation contracts、Core `Attributes/Filters` 与 Validation internal 实现、Npoi/CSV import path。
- **步骤**：
  1. 新增目标命名 Attribute：`ExcelRequiredAttribute`、`ExcelRegexAttribute`、`ExcelDateAttribute`、`ExcelRangeAttribute`、`ExcelMaxValueAttribute`、`ExcelUniqueAttribute`、`ExcelMaxLengthAttribute`。
  2. 当前 Attribute 标记 `[Obsolete]`，仅委托统一 Descriptor Factory，不能保留旧 Rule 运行路径。
  3. 固定执行序列为 raw Cell -> whitespace -> Required/Regex -> conversion -> Date/MaxValue/Range/custom -> Unique -> 成功后 Setter；错误 Code 区分 `MaxLength`/`MaxValue`。
  4. Regex 在构建时编译并配置 timeout；Date 支持 culture/format、DateTime/DateTimeOffset/可支持的 DateOnly 与 provider-neutral serial；Range 只允许批准的数值/日期类型且 build 时拒绝 min > max；MaxValue 使用无下界 Range Descriptor。
  5. 公开 `ExcelValidationContext` 删除/迁移可变 DuplicateValues，保留只读坐标、值、类型与 culture。
- **测试矩阵**：每规则正常/负例/边界；Regex timeout；日期格式/culture/serial；开闭区间；非法类型、非法 range 与转换失败；动态/固定同规则；命名 Validator 在 Build 绑定。
- **验收**：六类校验从真实 Excel 与 CSV 导入主链执行；旧 Attribute 使用同一 Descriptor，不存在第二套执行逻辑。

#### MAPVAL-201 Unique committed state 与 pending journal

- **依赖**：MAPVAL-200。**状态**：`TODO`。
- **目标**：移除 CSV `CloneDuplicateValues` 的全量复制，统一 Sheet scoped unique，确保失败行不占用唯一值。
- **步骤**：实现 internal `UniqueTracker`，记录 committed state 和 current-row pending journal；成功整行后提交，错误/取消回滚 pending；支持 `Ordinal`/`OrdinalIgnoreCase`、IgnoreNull、IgnoreEmpty、FirstRowNumber、`MaxTrackedUniqueValues`。
- **测试**：首行坐标、大小写、空/null、失败后相同值可用、多列、动态列、达到上限、100K 行线性行为。
- **性能验收**：避免每行克隆既有集合；在 `06-performance-gc.md` 对比改造前后 Allocation、Gen0/1/2、LOH、Working Set。

### Phase 3 - Profile V2、Registry 与 DI（P0）

#### MAPVAL-300 方向型双模型 Profile API

- **依赖**：MAPVAL-100。**状态**：`TODO`。
- **目标**：新增 `IMappingProfile<TImport,TExport>`、同模型 `IMappingProfile<T>` 和 `FluentSetting<TImport,TExport>`，严格隔离 Import/Export Builder。
- **步骤**：Import Builder 只暴露 Header/Alias/Index/Converter/Whitespace/Ignore/Image/Validation/Dynamic Import；Export Builder 只暴露 Header/Index/Placement/Formatter/Scale/Style/Layout/Comment/Merge/Image/Dynamic Export；Profile Configure 无请求级可变状态。
- **兼容**：保留旧 `ExcelMappingProfile<T>` 并 `[Obsolete]`，转换为新 Profile/Document source；不删除旧构造器或 Mapping overload。
- **测试**：不同 Import/Export DTO、同模型简化 Profile、方向 API 编译约束、Plan 不可变、旧 Profile 真实委托。
- **验收**：不能通过 Import Builder 调用 export-only API，也不能从 Export Builder 注册 import-only validation。

#### MAPVAL-301 Profile Registry 与注册扩展

- **依赖**：MAPVAL-300。**状态**：`TODO`。
- **目标**：实现显式和程序集扫描注册，并防止 DI 顺序决定结果。
- **修改范围**：Core 或 Npoi 中实际拥有 DI 依赖的扩展；不得把 `Microsoft.Extensions.DependencyInjection` 引入 Abstractions/Core 的 netstandard2.0 公共面。
- **步骤**：实现 `AddMappingProfile<TProfile,TImport,TExport>()`、`AddMappingProfilesFromAssembly()`、`AddMappingProfilesFromAssemblies()`；仅扫描非抽象、非开放泛型且实现批准接口的类型；Registry key 为 `(ProfileName, Direction, ModelType)`；重复默认/Nam​​ed key 在启动或 Build 失败；显式和扫描走同一 Registry；缓存有界且多租户隔离。
- **兼容**：`AddNpoi()` 保持 `void`，使用独立扩展完成分步注册；不得用返回 `IServiceCollection` 替换现有签名。
- **测试**：批量/显式、多程序集、命名、重复、开放泛型/抽象忽略、扫描顺序无影响、解析并发、外部 Docs Consumer。
- **验收**：没有“最后注册覆盖”；Registry 失败在首次 Plan Build 前可诊断。

### Phase 4 - JSON/XML v2、v1 兼容与配置安全（P0）

#### MAPVAL-400 单一 normalized document 与 Loader 演进

- **依赖**：MAPVAL-100、MAPVAL-300。**状态**：`TODO`。
- **目标**：JSON/XML 先归一化到同一个 `ExcelMappingDocument`，经同一 validate/compiler 主链构造方向 Plan；静态/DI loader 均是薄 facade。
- **步骤**：设计 v2 `version`、`profile`、稳定业务 model alias、`import`、`export`，各方向含 columns/dynamic columns/converter/validation/style/layout；禁止程序集限定 CLR 类型名；保留 `IExcelMappingConfigurationLoader` 现有 API，通过重载返回 document/diagnostic；新输出只写 v2。
- **v1 策略**：当前平铺 `Columns` 或缺省 version 视为 v1，仍可读取，生成 migration diagnostic 默认不中断；提供 v1 -> v2 converter/example；Obsolete 周期独立于 v1 文档生命周期。
- **验收**：JSON/XML 只存在一条 parser-validator-compiler 主链，旧 loader 返回的兼容 configuration 也来自 document。

#### MAPVAL-401 安全限制、错误路径与互操作测试

- **依赖**：MAPVAL-400。**状态**：`TODO`。
- **步骤**：继续 XML 禁用 DTD 和 `XmlResolver = null`；为 JSON/XML 流限制文件大小、深度、字符串、列、alias、validation、regex 长度；拒绝未知字段或按版本记录 diagnostics；Validation type 使用 allowlist，禁止任意 CLR type 反射；所有错误输出 JSONPath/XML path（例如 `import.columns[0].validations[1]`）。
- **测试**：JSON/XML v1 与 v2 等价、round trip、未知字段、损坏输入、DTD/XXE、超限、非法 Regex、未知 model alias、精确路径、流不关闭、UTF-8。
- **验收**：配置安全错误不会触发任意类型加载、外部实体解析或无界内存读取。

### Phase 5 - Provider SPI、IVT 和公共 API 收敛（P0）

#### MAPVAL-500 只读 Provider SPI 与生产 IVT 清理

- **依赖**：MAPVAL-101、MAPVAL-200、MAPVAL-400。**状态**：`TODO`。
- **目标**：以窄只读 SPI 取代 `Abstractions -> Core/Npoi` 的 production IVT。
- **步骤**：在 `Bing.Offices.Providers` 定义 provider 实际需要的只读 Plan view；应用 `[EditorBrowsable(EditorBrowsableState.Never)]`；使用 `IReadOnlyList`/不可变值，不含 NPOI、PropertyInfo setter、DI 容器或外部配置；Core internal Plan 实现 SPI，Npoi 仅消费 SPI；删除 Abstractions 的 Core/Npoi friend，只保留测试白名单。
- **测试**：扫描所有发布程序集的 `InternalsVisibleToAttribute` 并断言仅测试程序集；SPI API reflection 审计；NPOI signature 无 NPOI 泄露至 provider-neutral API；consumer/provider 独立 approval。
- **验收**：生产 IVT 数量为零；Npoi 不需要访问 Core internal 类型。

#### MAPVAL-501 Extension allowlist 与 API approval 更新

- **依赖**：MAPVAL-500。**状态**：`TODO`。
- **目标**：保留用户需要的 Stream File/Bytes、`AddNpoi()`、实际存在的 `AddBingOffices()` 和 Mapping Registration；将仅内部使用的 NPOI/Reflection/Expression/Type extensions internal 或 obsolete migration。
- **步骤**：先用仓库引用与 Docs Consumer 决定每个候选 public extension；禁止机械 internal 化；必要的低层 NPOI API 另行 ADR，不能隐式扩大主包面；更新 `PublicApiContractTest` 为 consumer/provider allowlist+成员签名 approval，保留旧 API 在 migration mode 的 Obsolete forwarding 验证。
- **验收**：三个 NuGet 包身份不变；公共面有明确 allowlist，旧公开 API 的兼容路径可执行且不是复制实现。

### Phase 6 - 文件职责、文档、验证与性能（P1）

#### MAPVAL-600 职责拆分与兼容迁移闭环

- **依赖**：Phase 1-5。**状态**：`TODO`。
- **目标**：按 Mapping、Validation、Providers、Planning、Npoi/Internal 拆分；每个 public top-level type 单独文件；移除大杂烩和无用兼容代码。
- **步骤**：先建立 symbol-to-test 映射，再移动；Core 放 compiler/descriptors/attributes adapters，Npoi 放 workbook adapters，Abstractions 放契约；以 `[Obsolete]` + forwarding 保留 current-major API；不存在真实迁移周期时不删除 public 类型；更新 XML comments、approval、Docs Consumer。
- **验收**：没有 `Interfaces.cs`/`Models.cs`/`Policies.cs`/`ValidationRules.cs` 新大杂烩；每个生产 public symbol 可追溯到测试。

#### MAPVAL-601 文档与可编译示例

- **依赖**：MAPVAL-600。**状态**：`TODO`。
- **目标**：新增或更新 `docs/excel/README.md`、`mapping-profile.md`、`mapping-json-xml.md`、`import-validation.md`、`dynamic-columns.md`、`nuget-migration.md`，及 ASP.NET Core 上传/失败文件返回示例。
- **必含内容**：双 DTO/同 DTO Profile、扫描/显式注册、JSON/XML v2、v1 -> v2、优先级冲突、六类校验、动态列校验、Stream ownership、兼容模式和 major 升级说明。
- **验证**：将所有 C# 示例加入 Docs Consumer 编译运行测试；不得只做 Markdown 文本检查。
- **验收**：Docs Consumer 仅引用本地 pack 的三个包时可以编译运行，且不依赖 NPOI implementation type。

#### MAPVAL-602 Benchmark、GC 与最终验证

- **依赖**：MAPVAL-601。**状态**：`TODO`。
- **目标**：提供可比较的真实前后数据，不能宣称零 GC 除非测量证实。
- **新增基准**：10K/100K 行、1/5 Unique 列；10K 多规则；100/500 dynamic Plan Build；100/1000 tenant cache 及淘汰；JSON/XML v1/v2 parse；assembly scan 与 explicit registration。
- **记录**：Wall Time、Operations/sec、Allocated Bytes、Gen0/Gen1/Gen2、LOH、Peak Working Set、硬件/Runtime/命令；对 CSV Unique 确认不再发生逐行全量 clone。
- **最终验收**：第 7 节命令通过，重开真实 `.xlsx` 检查 Profile、JSON/XML、动态列、六类错误坐标和输出内容；外部 Office/LibreOffice 不可用时标记待验证，给出命令、夹具和环境要求。

## 5. 测试用例矩阵

| 分类 | Given | When | Then |
| --- | --- | --- | --- |
| 优先级 | 五层对同一字段/Ignore/Rule Key 配置 | Build Import/Export Plan | 按固定顺序合并，加载/DI 顺序不影响结果 |
| Profile | Import DTO 与 Export DTO 不同 | 注册、解析并执行真实 XLSX | 双方向使用各自模型与 Builder；重复 key 失败 |
| 不可变性 | Build 后修改数组、List、Alias、ValueMap、Profile | 使用既有 Plan 并发导入导出 | 输出不变，租户/请求隔离 |
| 配置格式 | JSON/XML v1 与 v2 相同语义 | load/validate/compile/round-trip | 等价 Plan；v1 有非阻断 diagnostic；未知/超限有精确路径 |
| 校验 | Required/Regex/Date/MaxValue/Range/Unique 正常、边界、非法输入 | CSV/XLSX 固定和动态列导入 | 相同 Error Code/坐标；MaxLength 与 MaxValue 完全独立 |
| Unique | 第一行成功、第二行重复或第一行失败 | Row pending -> commit/rollback | 失败行不占位，首行号/比较/空值/上限准确 |
| 边界 | 空/损坏/不可读流、不可写失败流、取消、Dispose | 真实 Provider 调用 | 无吞异常，stream ownership 和结果/异常契约正确 |
| NuGet | `dotnet pack` 后空消费者 | 仅引用公开包 API 编译运行 | 三包完整，普通 API 无 NPOI，SPI/IVT approval 通过 |

只 Mock 时间、随机、I/O 或外部依赖；不得 Mock 被测 compiler、Plan 或 provider 内部调用来替代行为断言。单元测试方法使用英文 `Method_State_Expected()`，测试目的使用中文 XML 注释，采用 AAA。

## 6. 文件清单

### 已确认将修改

- `src/Bing.Offices.Abstractions/AssemblyInfo.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/*`
- `src/Bing.Offices.Abstractions/Bing/Offices/Validations/*`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/*`、`Exports/*`、`Csv/*`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelTypeMapFactory.cs`
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
- `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs` 与现有 Filter Attribute
- `src/Bing.Offices.Npoi/Extensions/Extensions.Service.cs`
- `src/Bing.Offices.Npoi/ExcelColumnPlan.cs`、`Imports/NpoiExcelImporter.cs`、`Exports/NpoiExcelExporter.cs`
- Unit、Integration、Docs Consumer、Benchmark、`PublicApiContractTest.cs` 与 `docs/excel/*`

### 候选新增

- `src/Bing.Offices.Abstractions/Bing/Offices/Providers/*`：只读 SPI 契约。
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/IMappingProfile*.cs`、`FluentSetting*.cs`、document/diagnostic 契约。
- `src/Bing.Offices.Core/Bing/Offices/Mappings/Planning/*`、`Validation/*`、`Configurations/*`：internal compiler、descriptor、merger、unique tracker、format parser。
- `src/Bing.Offices.Npoi/Planning/*`、`Internal/*`：SPI adapter；仅实际需要时增加，禁止 NPOI 类型穿透。
- 目标 Task 状态文件、精确职责单测、真实 xlsx fixture、Docs Consumer profile/configuration examples 与 Benchmark classes。

## 7. 真实验证命令

执行顺序以 Phase 0 记录的实际 SDK/lock 状态为准。下列命令来自 `.github/workflows/ci.yml` 或现有项目文件；不得把未运行写为通过。

```powershell
dotnet restore Bing.Offices.sln --locked-mode
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore
dotnet pack src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release --no-build -o artifacts/packages
dotnet pack src/Bing.Offices.Core/Bing.Offices.Core.csproj -c Release --no-build -o artifacts/packages
dotnet pack src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-build -o artifacts/packages
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore
git diff --check
```

执行 Agent 必须另外以临时目录创建 package consumer，仅使用 `artifacts/packages` 的三个 `.nupkg` 恢复并编译运行 Docs Consumer 示例。不得调用 `Publish.bat`，不得将包上传到外部服务。

## 8. 完成门槛与风险

Task 仅在 JSON/XML v1 读取和 v2 双方向、四来源同 compiler、五层优先级、双模型 Profile、扫描/显式注册、六类校验、固定/动态 Excel/CSV 等价、MaxValue/MaxLength 分离、三包保留、生产 IVT 为零、扩展 allowlist、兼容旧 API forwarding、Build/Unit/Integration/Docs/API/pack consumer、Benchmark/GC、文档和所有状态证据均完成时标为 `DONE`。

已知风险是 netstandard2.0 对 API 选择的限制、公开 API approval 变更的二进制兼容性、NPOI 对复杂 XML/Workbook 重写的互操作差异，以及大量用户未提交改动可能与本任务重叠。发现不可推断且会改变数据语义的问题时，写入 `03-decisions.md` 标记 `BLOCKED`，继续不依赖任务；不得以删测试、降低断言、catch 吞异常、跳过验证或修改版本发布来制造通过。

最终 `07-final-report.md` 必须逐项声明：Task ID、最终状态、Compatibility Mode、实现/API/JSON/XML/NuGet/IVT、测试、Benchmark/GC、文档、待验证风险、变更文件，并明确写明“未执行 Commit、Push、PR、Tag 或 NuGet 发布”。
