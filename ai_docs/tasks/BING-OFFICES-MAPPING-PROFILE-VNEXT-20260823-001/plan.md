# BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001 实施计划

> 计划状态：READY_FOR_EXECUTION  
> 计划范围：下一主版本 MappingProfile、JSON/XML 映射文档、配置合并、公开 API 与程序集边界收口。  
> 本轮限制：仅完成计划；不得实施、修改测试/配置或更新任务状态文件。后续由“开始实施”/`execute-plan` 接手。

## 1. 输入、基线与结论

### 1.1 已读取的输入

- 根目录 `AGENTS.md`、`.github/copilot-instructions.md`、`.github/skills/chinese-comments/SKILL.md`、`.github/prompts/create-plan.prompt.md`。
- 需求指定的 `ai_docs/codebase-analysis/2026-08-24-bing-offices-implementation-review.md` 与 `ai_docs/architecture/2026-08-23-bing-offices-mapping-profile-and-comments-solution.md` 在工作区按精确路径和文件名检索均不存在。实施前不应虚构其结论；本计划以当前源码、测试和 `docs/excel/mapping-profile.md` 为证据，并将缺失文档列为交接风险。
- 发布项目为 `Bing.Offices.Abstractions`、`Bing.Offices.Core`、`Bing.Offices.Npoi`，依赖方向为 `Abstractions <- Core <- Npoi`。前两者目标为 `netstandard2.0`，NPOI 使用 `net8.0;net7.0;net6.0;netcoreapp3.1`；当前版本为 `1.0.4`。
- 当前工作区存在上一 Mapping/Validation 任务的大量未提交改动。不得清理、回退、混入或假定这些变动归本任务所有；实施开始必须先记录 `git status --short`、`git diff --stat` 与基线 commit，并将本任务修改逐文件区分。

### 1.2 当前实现完成度

| 范围 | 已实现证据 | 判定 | vNext 缺口 |
|---|---|---|---|
| 双向 Profile | `IMappingProfile<TImport,TExport>`、`ExcelMappingProfile<TImport,TExport>`、`FluentSetting<TImport,TExport>` 已存在 | 部分完成 | 单向导入/导出契约不存在；同模型接口只是继承双向接口，无法表达真正单向 Profile。 |
| Registry | `MappingProfileRegistry` 使用 `(名称,方向,模型)` 键，能够注册双向快照 | 部分完成 | 快照模型强制同时具有 Import/Export 类型，注册冲突判断也按双向耦合，不能容纳四种形状。 |
| DI | `AddMappingProfile<TProfile,TImport,TExport>` 与程序集扫描可工作 | 部分完成 | 调用方必须给出三个泛型参数；扫描只取第一个 `IMappingProfile<,>`；命名不符合目标 API，且没有统一四形状处理。 |
| JSON/XML v2 | `ExcelMappingDocument` 有 Import/Export，加载器有深度、大小、未知字段和 XXE 防护，v1 会迁移 | 部分完成 | `Profile`、`ModelAlias` 仍在文档根节点；v1 被隐式迁移为 Import，方向决策没有显式迁移契约。 |
| 合并 | `MappingConfigurationMerger` 按列合并，支持部分 Validation clear/remove/replace 与 ValueMapping replace/append | 部分完成 | 普通标量没有显式 reset；集合 `null`/空/清空语义不完整；动态列、Style、Layout 只有“非空整体覆盖”，不能表达 patch、clear/remove。 |
| 计划构建 | `ExcelMappingPlanFactory` 已能按方向查 Registry、核对别名、缓存计划，NPOI 使用 Core 工厂 | 部分完成 | 读取根级 Profile/Alias；缓存键与归一化模型耦合现有文档形状；优先级链未以统一可审计的配置层实现。 |
| API/边界 | `PublicApiContractTest` 已检查顶层类型、NPOI 泄漏和生产 IVT；生产 IVT 目前只见测试友元 | 部分完成 | 基线仍批准 legacy `ExcelMappingProfile<T>`、`IMappingProfile<T>` 和旧 DI API；需要 next-major 精确基线与程序集职责重审。 |
| 测试、消费者、文档、性能 | Unit/Integration/Docs/Benchmark 项目均存在；Docs Tests 使用已打包的 `1.0.4` 包并编译 Markdown fences | 部分完成 | 现有用例围绕旧 v2 形状；尚无四 Profile 形状、方向级元数据、完整 patch 真值表、v1 显式迁移、v2 包消费者和回归基线证据。 |

**整体判断：约 45% 的目标能力已有可用骨架或前置基础，但四形状 API、方向级文档模型和完整 patch 语义是结构性未完成项，不能以增量兼容补丁完成。** 当前机制的复杂度集中在“一个双向快照兼容所有情况”，DI、Registry、文档、缓存和工厂均依赖该假设；vNext 应统一成单方向 `ProfileDescriptor`，降低耦合并让失败消息可定位。

### 1.3 目标架构与不可变约束

1. 仅保留并正式支持四种互斥、显式形状：
   - `IImportMappingProfile<TImport>`；
   - `IExportMappingProfile<TExport>`；
   - `IMappingProfile<TModel>`；
   - `IMappingProfile<TImport,TExport>`。
2. 禁止通过 sentinel 类型、方向 bool、空方向、`object`、nullable `Type` 或每种形状各自的注册 API 表达方向。
3. 内部规范化为方向单一的 `ProfileDescriptor`：`Name`、`Direction`、`ModelType`、`Configuration`、`Source/ProfileType`；Registry 唯一键严格为 `(ProfileName, Direction, ModelType)`。
4. DI 只暴露 `services.AddMappingProfile<TProfile>()` 与 `services.AddMappingProfiles(typeof(TProfile).Assembly)`；扫描必须处理一个类型实现的所有受支持 Profile 契约，抽象、开放泛型和非法/冲突实现必须有确定异常。
5. JSON/XML 新格式在每个 `import`/`export` 节点内定义 `profile`、`modelAlias`、配置；根节点只保留文档级字段（版本、租户、配置版本等）。当某方向未提供配置时，应有明确、可测试的默认/缺失行为。
6. 优先级固定为 `Convention < Attribute < MappingProfile Fluent < JSON/XML Mapping Document < Request Override`。所有合并均经一个方向化、可测试的 patch 引擎，保留 unset、set、clear/reset、append/replace/remove 的区别。
7. 当前 major 的 legacy 类型和入口不作为 v2 运行时兼容层保留。v1 JSON/XML 只由独立、显式的迁移工具处理，必须要求调用方选择目标方向或提供映射策略，禁止静默默认导入。
8. 公开发布程序集不得新增生产 `InternalsVisibleTo`；NPOI 不能出现在 Abstractions/Core 公开签名中。所有本任务变更 C# 声明、公开成员、参数、返回值、异常行为及复杂私有状态按中文 XML 注释 skill 完成中文 XML 文档；实现接口时优先 `/// <inheritdoc />`。

## 2. 变更边界

### 已确认的修改文件/区域

- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingProfileContracts.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingProfile.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingProfileV2.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocument.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/IMappingProfileRegistry.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingProfileRegistry.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingConfigurationMerger.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingConfigurationCloner.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingDocumentCloner.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDocumentFactory.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingConfiguration.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelColumnConfiguration.cs`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`
- `src/Bing.Offices.Npoi/Extensions/MappingProfileServiceCollectionExtensions.cs`
- `tests/Bing.Offices.Tests/MappingProfileRegistryTest.cs`
- `tests/Bing.Offices.Tests/MappingProfileV2Test.cs`
- `tests/Bing.Offices.Tests/StreamPipelineTest.cs`
- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
- `tests/Bing.Offices.Docs.Tests/DocsConsumerTest.cs`
- `tests/Bing.Offices.Docs.Tests/DocsExamples.cs`
- `docs/excel/mapping-profile.md`
- `docs/excel/mapping-json-xml.md`
- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
- `version.props`、Docs Tests 包版本/本地源配置及相关 lock files（仅在版本和 package-consumer 策略确认后修改）。

### 候选文件/区域（先检索再决定）

- 导入/导出 builder、请求对象、CSV 管线和 NPOI 归一化入口：确认是否仍接受 legacy `object profile` 或根级元数据。
- `IExcelMappingPlan`、`ExcelTypeMapFactory`、模型别名注册表及计划缓存相关实现：确认方向级 descriptor 是否需要直接参与缓存身份和别名校验。
- `ExcelMappingDynamicColumnConfiguration`、Style/Layout 配置、验证和值映射枚举：在 patch 模型落地时决定是否引入明确操作 DTO/枚举。
- `tests/Bing.Offices.Tests.Integration/**`：为真实 NPOI + DI + JSON/XML 端到端路径补集成验证。
- `tests/Bing.Offices.ProfileFixtures/**`：为外部程序集、单向/同模型/多接口 Profile 消费者增加无测试程序集耦合的 fixture。
- `docs/excel/README.md`、根 README、包 README/打包目标与 CI workflow：仅在实际链接、包版本、restore 源或门控命令受影响时修改。

## 3. 分阶段实施计划

### Phase 0 - 基线隔离与设计校准

#### P0-001：建立可追溯实施基线

- **依赖**：无。
- **目标**：在不改变现有脏工作区的前提下，记录 vNext 开始时的仓库和发布基线，防止上一任务改动被误归属或回退。
- **证据**：当前版本为 `1.0.4`；Docs Tests 从 `artifacts/packages-1.0.4` 消费包；工作区已知含先前任务未提交变更。
- **步骤**：
  1. 记录 `git status --short`、`git diff --stat`、`git rev-parse HEAD`、当前 `version.props`，并建立“现有改动/本任务改动”逐文件边界。
  2. 读取或重新定位缺失的架构/评审文档；若仍缺失，在实施记录中明确标注为 unavailable，禁止将其假设写成事实。
  3. 抽取当前公开 API 基线、生产 IVT 列表、NPOI 导出类型列表和 package lock 状态。
  4. 定义 v2 版本号、预发布策略、包输出目录和 Docs consumer 要消费的本地产物；版本变更必须统一三包与 docs consumer。
- **测试/验证**：只执行当前基线的 restore/build/test 诊断；不将失败自动归因于本任务，失败需与脏工作区区分。
- **风险**：已有源码与旧计划/评审不一致；缺失设计文档不能由推测替代。
- **验收**：实施记录可指出每项新改动相对哪个基线发生；版本、包源、目标框架和禁止触碰文件明确。

#### P0-002：冻结 v2 迁移契约与 API 删除清单

- **依赖**：P0-001。
- **目标**：先决定所有 breaking change 和 v1 迁移输入/输出，再修改模型，避免半兼容 API 滞留。
- **步骤**：
  1. 制定并评审 API 表：新增四 Profile 契约、统一注册 API、方向 descriptor/registry API、v1 显式迁移器；删除/不再公开的 `ExcelMappingProfile<T>`、legacy `IMappingProfile<T>` 定义方式、`AddMappingProfile<TProfile,TImport,TExport>`、`AddMappingProfilesFromAssembly`、`AddMappingProfilesFromAssemblies` 和接受 `object profile` 的兼容入口。
  2. 为 v1 JSON/XML 迁移器定义 `targetDirection` 必填行为，或采用具名策略接口；v1 平铺配置不能自动猜测 Import。迁移输出必须是方向节点持有 Profile/Alias 的 v2 文档，并给出 JSON/XML 示例。
  3. 决定 direction node 缺失、profile 缺失、alias 缺失、同名不同模型、同类型不同方向以及一个 Profile 多契约的完整异常语义与错误消息关键字段。
  4. 制定公开 API 基线更新策略：先让 Contract Test 显式暴露删除/新增差异，再以审阅后的 v2 允许列表收口，禁止宽松 Contains 式断言。
- **测试/验证**：以测试矩阵评审：每个被删除 API 有替代路径；每个 v1 输入必须要求显式方向；每个新公开 API 有外部消费者编译用例。
- **风险**：`IMappingProfile<T>` 名称在目标中仍需表示“同模型双向”，与当前“继承双模型”的含义相近但实现方式需重建，不能简单删除后导致歧义。
- **验收**：所有 API 的“保留/替换/删除”与版本策略可在发布说明中逐项追踪；不存在未决的静默 v1 默认方向。

### Phase 1 - Abstractions：四形状 Profile 与方向化注册表

#### P0-101：重建四种 Profile 契约及方向配置 Fluent API

- **依赖**：P0-002。
- **目标**：使类型系统直接表达导入、导出、同模型双向和异模型双向四种形状，不引入方向 flag 或哨兵泛型。
- **修改范围**：`MappingProfileContracts.cs`、`FluentSetting.cs`、`ImportMappingBuilder.cs`、`ExportMappingBuilder.cs`、`ExcelMappingProfile*.cs` 与相关 builders。
- **步骤**：
  1. 定义 `IImportMappingProfile<TImport>` 与 `IExportMappingProfile<TExport>`，各自只暴露对应方向的 `Configure` 方法和 builder。
  2. 重定义 `IMappingProfile<TModel>` 为同模型双向的明确契约；保留 `IMappingProfile<TImport,TExport>` 作为异模型双向契约。两者均以明确双方向配置表达，不通过接口继承让 scanner 误判重复契约。
  3. 将 profile snapshot/构造器重构为可生成一个或两个不可变方向 descriptor；删除将单向 Profile 伪装成空另一方向的状态。
  4. 逐个更新/补充中文 XML 注释；所有配置方法的参数、异常条件和方向含义需明确。
- **单元测试矩阵**：
  - Given 四种独立 profile，When 规范化，Then 生成正确数量、方向、模型类型与不可变配置。
  - Given 同一 profile 类型实现多个合法契约，When 规范化，Then 每个方向 descriptor 均出现且不漏扫描。
  - Given 重复或冲突的同一方向/模型配置，When 规范化，Then 在注册前抛出带 Profile/Direction/Model 的确定异常。
  - Given 配置后修改 builder 输入集合，When 读取 descriptor，Then 快照不变。
- **验收**：没有 `object`、nullable `Type`、空方向或 `T=object` 用于表达未支持方向；所有四形状都能生成可查询 descriptor。

#### P0-102：以 `ProfileDescriptor` 收口 Registry 与模型别名关系

- **依赖**：P0-101。
- **目标**：Registry 仅存方向单一不可变 descriptor，并以 `(ProfileName, Direction, ModelType)` 精确判重、查询和报错。
- **修改范围**：`IMappingProfileRegistry.cs`、`MappingProfileRegistry.cs`、`ExcelModelAliasRegistry.cs`、必要的内部 key/descriptor 文件。
- **步骤**：
  1. 定义内部或按必要性公开的 `ProfileDescriptor`，固定 `Name`、`Direction`、`ModelType`、`Configuration`、`Source/ProfileType`；对 null/空名称、非法 enum、非 class/new 约束和可变集合输入做边界校验。
  2. Registry 使用 descriptor 的单键注册/解析，不再依赖“同一个 snapshot 被两方向引用”的 `ReferenceEquals`。
  3. 冲突异常输出完整 key 和两个来源类型，避免 DI 注册顺序决定结果。
  4. 将模型别名解析与方向 descriptor 配对：别名类型必须等于请求模型；若别名指定 profile，则必须等于方向节点 Profile；不得读取根级文档字段。
  5. 确认并保留线程安全和快照不可变性；不为性能引入生产 IVT。
- **单元测试矩阵**：正常解析、名称大小写策略、同名不同方向/模型可共存、同键冲突、并发读取、别名类型不匹配、别名 profile 不匹配、注册失败后 Registry 不污染。
- **验收**：所有 Registry 查询和错误信息均可从唯一键追溯；不再存在双向 snapshot 的一致性假设。

### Phase 2 - 文档模型、显式迁移与 Patch 语义

#### P0-201：将 JSON/XML 设计改为方向节点元数据

- **依赖**：P0-102。
- **目标**：重构 `ExcelMappingDocument` 为方向节点拥有 `Profile`、`ModelAlias` 和 Mapping Configuration，根节点不再承载方向业务元数据。
- **修改范围**：`ExcelMappingDocument.cs`、相关 clone/factory、`ExcelMappingConfigurationLoader.cs`、计划工厂及请求归一化调用点。
- **步骤**：
  1. 引入明确的 direction node DTO（或等价强类型结构），防止 Import/Export 重用可变同一配置实例；根节点只保留 version、tenant、configuration version 等文档范围字段。
  2. 更新 JSON camelCase 与 XML 元素白名单、反序列化、序列化、克隆、文档验证及诊断位置。必须继续禁止 DTD/外部实体、限制深度/字节数、拒绝未知字段和 CLR 类型名 alias。
  3. 从计划工厂、CSV、NPOI import/export、builder/request 中删除根级 `Profile`/`ModelAlias` 读取，改为 `direction` 节点读取。
  4. 对无该方向节点、空 Profile、空 alias 的语义执行 P0-002 已决定的规则，不做隐式 fallback。
- **单元测试矩阵**：JSON/XML round-trip、独立 import/export profile+alias、跨方向 alias/profile 不串扰、未知根/节点字段、非法 alias、大小/深度边界、流不关闭、序列化不写根级 Profile/Alias。
- **验收**：生成的 JSON/XML 只在 import/export 节点出现业务 `profile`/`modelAlias`；两方向的 plan 身份和缓存键独立。

#### P0-202：提供独立 v1 到 v2 显式迁移器

- **依赖**：P0-201。
- **目标**：将兼容从主加载器移出，避免 current-major 静默“v1 等于 Import”的行为进入 vNext。
- **修改范围**：新建明确迁移 API/类型（位置由 Phase 0 API 表决定）、loader、文档、legacy API 删除点。
- **步骤**：
  1. 新迁移器接收 v1 JSON/XML 及调用方明确的目标方向或具名映射策略；产出 v2 direction node 文档和结构化诊断。
  2. v2 主加载器只接受 v2 结构；收到 v1 根对象/根元素时，返回说明迁移入口和方向要求的异常，而非静默转换。
  3. 迁移器保持现有安全限制、UTF-8 流处理与输入大小限制；不接受程序集限定类型名作为 alias。
  4. 文档提供 import/export 两种迁移样例及升级前后 JSON/XML 对照。
- **单元测试矩阵**：v1 JSON/XML 到 Import、到 Export、缺失 direction、非法 target、迁移诊断、v2 loader 拒绝 v1、迁移后序列化/反序列化一致。
- **验收**：任一 v1 迁移方向均由调用方可见地选择；运行时没有默认 Import 迁移路径。

#### P0-203：实现统一五层 precedence 与可判定 Patch 引擎

- **依赖**：P0-101、P0-201。
- **目标**：把现有分散的 null/空集合判断替换为一个方向化合并规则，完整区分未设置、设置、显式 reset/clear、append、replace、remove。
- **修改范围**：`ExcelMappingConfiguration.cs`、`ExcelColumnConfiguration.cs`、动态列/Style/Layout DTO、`MappingConfigurationMerger.cs`、cloner/factory、缓存键和相关 builder。
- **步骤**：
  1. 为标量、列、别名、校验规则、值映射、动态列、Style、Layout 定义明确 patch 表达。避免以 `null` 或空列表猜测操作；必要时增加 operation enum/patch DTO。
  2. 实现严格顺序：Convention、Attribute、Profile Fluent、Document、Request。每层输出独立快照；低层不得被高层输入对象共享引用。
  3. 列字段支持 set/reset；集合支持 unset、clear、replace、append/remove，并明确空字符串、空集合和 `null` 的合法性。
  4. DynamicColumns、Style、Layout 必须从现有“非空整体替换”升级为可表达 clear/reset 和按稳定 key 合并/删除的语义；若某项不支持部分 patch，显式限制并在文档/异常中说明，不留隐式行为。
  5. 更新缓存键，使它反映最终规范化的方向配置及所有对结果有影响的 patch 状态，防止不同 patch 产生相同 key；测试 hit/miss/重复渲染和租户隔离。
- **单元测试矩阵**：
  - 每一层覆盖前一层的正常字段；
  - null/unset 不覆盖、空集合的预定语义、clear/reset、append/replace/remove；
  - 规则/值映射的大小写、重复、删除不存在项；
  - 动态列/Style/Layout 的 set/clear/merge；
  - 五层组合的最终完整配置断言；
  - 原输入对象修改不影响结果；不同最终配置 cache miss、相同规范化配置 cache hit。
- **Mock 边界**：只 mock 外部 converter/validation 服务；合并器、descriptor、registry、cache key 直接以真实对象断言行为，不验证内部调用次数。
- **验收**：所有组合以完整对象/完整 SQL 类似的精确值断言，而非 `Contains`；合并行为可由一张公开的真值表复现。

### Phase 3 - Core/NPOI 调用链与公开 API 收口

#### P0-301：让 Plan Factory、请求和 Provider 只消费方向 descriptor

- **依赖**：P0-201、P0-203。
- **目标**：贯通真实 CSV/XLSX/NPOI 管线，使 Registry、Document 和 Request Override 以同一方向模型生成计划。
- **修改范围**：`ExcelMappingPlanFactory.cs`、`ExcelTypeMapFactory.cs`、`ExcelMappingDocumentFactory.cs`、Import/Export builders 与 requests、CSV pipeline、NPOI normalized document 方法、`IExcelMappingPlan*` 如需要。
- **步骤**：
  1. 删除或迁移 `Create<T>(object profile, ...)` 等 legacy 入口，替换成类型安全 document/direction/override 流程。
  2. 由方向节点的 profile/alias 查询 descriptor，执行别名、请求模型、Profile key 三者一致性校验，再按统一 precedence 生成不可变 plan。
  3. 审核所有 import/export/CSV 路径，确保没有直接读取过时 root 属性或绕过合并器；异常必须包含方向、模型、profile/alias 和失败原因。
  4. 复查 `ExcelMappingPlanFactoryProvider` 的 Core 所有权，NPOI 只能通过 Abstractions 契约及 Core provider 构造，不新增 Core/NPOI 相反依赖。
- **测试**：Unit 覆盖每个入口的方向选择、Profile 未找到、alias 不匹配、request 最高优先级、plan 只读；Integration 用真实 XLSX 和 CSV 验证 import/export 四形状及 error 坐标。
- **验收**：真实 NPOI 和 CSV 都不依赖 legacy API，所有计划身份来自单方向配置且缓存不会串方向。

#### P0-302：统一 DI 注册 API，扫描所有合法契约

- **依赖**：P0-101、P0-102。
- **目标**：外部消费者只使用两个无歧义 DI 入口。
- **修改范围**：`MappingProfileServiceCollectionExtensions.cs`、Profile fixtures、Registry/DI tests、公开 API 基线。
- **步骤**：
  1. 实现 `AddMappingProfile<TProfile>()`：反射/规范化该 profile 的所有合法契约，注册具体类型和 descriptors；可选 profile name 只能采用已决、稳定的命名规则，避免第三泛型参数 API。
  2. 实现 `AddMappingProfiles(Assembly)`：按 `FullName` 稳定排序，扫描所有接口而非 `FirstOrDefault`；抽象、接口、开放泛型跳过；不支持的实现形状与重复 key 失败。
  3. 删除旧扫描方法和显式三泛型注册 API，更新所有源码、文档、fixtures、fences 与 consumer 调用。
  4. 覆盖同类多契约、多个程序集顺序无关、重复程序集、同 key 冲突、DI 生命周期、构造器依赖、注册失败前后服务集合一致性。
- **验收**：所有四类 profile 在显式和程序集扫描下都能解析；公共 API 只保留两个注册入口，不存在每种 shape 一个扩展方法。

#### P1-303：公开面、程序集边界与中文注释审计

- **依赖**：P0-301、P0-302。
- **目标**：将 next-major 删除项真正从 public surface 收口，避免 compatibility dead code 和低层依赖泄漏。
- **步骤**：
  1. 更新 `PublicApiContractTest` 为 v2 精确 top-level types/member signatures；删除 legacy 类型/方法的期望值并加入四 profile、方向 document node、迁移器与统一 DI API。
  2. 精确检查三发布程序集的 exported types、公开方法/构造器参数和返回值；NPOI 仍只开放注册扩展，Core/Abstractions 不泄漏 NPOI。
  3. 检查所有 `InternalsVisibleTo`，生产 friend 必须为零；测试友元仅在确有必要时保留。
  4. 搜索并处理本范围 `TODO`、`NotImplementedException`、fallback、obsolete compatibility 分支：与 Mapping vNext 相关者删除、实现或显式移出范围；无关项只记录，不进行搭车重构。
  5. 根据中文 comments skill 对本任务所有变更 C# API 进行 XML 文档审计，确保摘要、泛型、参数、返回、异常和继承关系准确。
- **验收**：Public API 精确测试、NPOI 泄漏测试、生产 IVT 测试均通过；编译输出 XML 文档中不存在本任务新增 public API 的缺失说明。

### Phase 4 - 测试、外部包消费者、文档与性能证据

#### P0-401：补齐 Unit 与 NPOI Integration 证据矩阵

- **依赖**：Phase 1-3。
- **目标**：测试真实行为而非仅验证接口存在，并覆盖正常、边界、异常和并发/缓存行为。
- **修改范围**：现有 `Bing.Offices.Tests`，必要时新增职责明确的测试类；`Bing.Offices.Tests.Integration`。
- **步骤**：
  1. 将四 profile shapes、registry/descriptor、统一 DI、direction node、显式 v1 migration、patch precedence、cache identity 拆成职责测试类，避免 `ExcelWorkbookRequestTest` 继续成为大杂烩。
  2. 每个测试采用英文 `Method_State_Expected()`、中文 XML 测试目的和 AAA；对于 JSON/XML 与 SQL 输出同等级的格式需求，断言完整规范化结果或完整序列化字符串，不仅断言片段。
  3. Unit 不访问网络、真实 DB/缓存/文件系统；流测试使用 `MemoryStream`。Integration 仅使用真实 NPOI/XLSX、DI 与本地临时流，不引入外部服务、sleep 或随机等待。
  4. 所有目标集执行 Unit：`net8.0;net7.0;net6.0;net5.0;netcoreapp3.1`；Integration 执行 net6/net8。对因为 v2 删除 API 导致的旧测试失败，必须迁移断言，不允许把测试删除为“通过”。
- **验收**：每一项 P0 architecture decision 均至少有一个直接 Unit 测试；真实 CSV/XLSX + DI 覆盖四形状和文档方向元数据。

#### P0-402：以新发布包运行 package-consumer 与 Markdown 文档测试

- **依赖**：P0-301、P0-302、P1-303。
- **目标**：验证外部消费者仅引用构建出的 v2 三包，文档与 API 没有同仓源码引用假象。
- **修改范围**：Docs Tests csproj/lock 文件、`DocsConsumerTest.cs`、`DocsExamples.cs`、docs Markdown、打包/本地源配置（按实际方案）。
- **步骤**：
  1. 先 pack 三个 v2 项目到隔离本地 feed，再让 Docs Tests restore 该 feed 的同版本三包；核对其没有 ProjectReference 到源码。
  2. 外部 consumer 同时运行 explicit/assembly scan 的四种 profile、方向级 JSON/XML、v1 迁移器与 NPOI `AddNpoi` 主链。
  3. 更新并扩展 docs fence runner，使 docs/excel 所有 C# fence 逐段编译并实际执行；维护 fence 数量断言，使文档新增代码不能静默绕开测试。
  4. 文档明确 next-major breaking changes、完整迁移表、方向节点 schema、五层 precedence/patch 真值表、DI 统一入口和安全限制。
- **验收**：清理测试输出后仍能从本地 v2 packages restore、编译和执行 consumer；Markdown code fences 全量通过，无过时 API 示例。

#### P1-403：性能、资源与复杂度回归证据

- **依赖**：P0-203、P0-301、P0-401。
- **目标**：确认 descriptor/patch 重构没有引入明显计划创建、注册扫描、缓存或 LOH 回归。
- **修改范围**：`benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`、`Program.cs` 资源 probe、基线说明/制品目录。
- **步骤**：
  1. 为四形状 descriptor 构建、程序集扫描、五层 merge、方向 document plan build、cache hit/miss 加入可重复 benchmark 场景；固定数据规模、预热与基线参数。
  2. 复用上一任务的独立 child-process resource probe 模式，分别记录 sampled LOH peak、retained LOH、peak working set、实际 tenant/profile plan cardinality，不得混淆“采样值”和“峰值”。
  3. 与 P0-001 的可比基线比较，记录绝对值、命令、环境、输入规模和允许阈值；若无可比旧场景，明确“不适用”原因而非伪造改善百分比。
  4. 如 profile descriptor 或缓存 key 是热路径，进行 allocation/retained-memory 分析，避免通过无限缓存或隐藏静态集合换取表面吞吐。
- **验收**：基准和资源结果可复跑、原始数据可追溯；无未解释的明显时间/分配/LOH/工作集回归，缓存容量边界仍有效。

### Phase 5 - 发布前收敛

#### P0-501：跨框架完整验收与变更审查

- **依赖**：Phase 4。
- **目标**：在发布前将 API、行为、包消费、文档和性能证据收敛为可审查结论。
- **步骤**：
  1. 恢复依赖并构建主解决方案，核对所有 lock file 的有意变更；不得使用 `--no-restore`、`--no-verify` 或跳过安全检查作为通过手段。
  2. 执行以下真实项目命令，输出到本任务制品目录并保留失败详情：
     ```powershell
     dotnet restore .\Bing.Offices.sln
     dotnet build .\Bing.Offices.sln -c Release --no-restore
     dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore
     dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net7.0 --no-restore
     dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore
     dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net5.0 --no-restore
     dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f netcoreapp3.1 --no-restore
     dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -c Release -f net6.0 --no-restore
     dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-restore
     dotnet pack .\src\Bing.Offices.Abstractions\Bing.Offices.Abstractions.csproj -c Release --no-build
     dotnet pack .\src\Bing.Offices.Core\Bing.Offices.Core.csproj -c Release --no-build
     dotnet pack .\src\Bing.Offices.Npoi\Bing.Offices.Npoi.csproj -c Release --no-build
     dotnet test .\tests\Bing.Offices.Docs.Tests\Bing.Offices.Docs.Tests.csproj -c Release
     ```
  3. 在实际 target framework 可用性受 SDK/runtime 限制时，记录缺失运行时和替代 CI 证据；不得把 framework 排除从项目中删除。
  4. 审查 diff：不包含上一任务无关改动、无新增 production IVT、无 NPOI 公共签名、无 legacy public API、无未更新 XML doc、无未同步 docs/package version。
- **风险**：本地环境可能缺少 net5/coreapp3.1 runtime；Docs Tests 只会验证其 restore 到的 v2 包，必须在 pack 后重 restore。
- **验收**：所有可运行 target 全绿，无法本地运行 target 有明确 CI 证据；三 package 及 consumer 验证使用同一 v2 version；未执行 commit/push/publish。

## 4. 关键验收追溯矩阵

| 生产符号/行为 | 直接测试责任 | 主要项目 |
|---|---|---|
| 四种 `I*MappingProfile` 契约与 descriptor 正规化 | 四形状、同类多契约、冲突、不可变快照 | `Bing.Offices.Tests` |
| `MappingProfileRegistry` 精确三元键 | 同名异方向/异模型、同键冲突、并发、失败不污染 | `Bing.Offices.Tests` |
| `AddMappingProfile<TProfile>()` / `AddMappingProfiles(Assembly)` | 显式、扫描、排序、重复、抽象/开放泛型、DI 依赖 | `Bing.Offices.Tests` + ProfileFixtures |
| 方向节点 JSON/XML 与 alias/profile 校验 | JSON/XML round-trip、unknown 字段、跨方向隔离、安全限制 | `Bing.Offices.Tests` |
| v1 显式迁移器 | import/export 指定、缺方向失败、诊断、v2 loader 拒绝 v1 | `Bing.Offices.Tests` |
| 五层 Patch 合并与 cache key | 组合真值表、set/reset/clear/append/replace/remove、hit/miss | `Bing.Offices.Tests` |
| Core plan 与 CSV/NPOI 链 | 四形状真实 CSV/XLSX、请求覆盖、错误坐标、alias mismatch | Unit + `Bing.Offices.Tests.Integration` |
| v2 public surface/边界 | 精确 exported type/member baseline、NPOI leak、生产 IVT | `Bing.Offices.Tests` |
| 包消费者与文档 | 仅 NuGet package 引用、Markdown fence compile/execute、迁移样例 | `Bing.Offices.Docs.Tests` |
| 性能与资源 | descriptor/merge/cache benchmark、子进程 LOH/working set evidence | `Bing.Offices.Benchmarks` |

## 5. 风险与处理原则

- **Breaking change 风险**：这是有意的 next-major 收口。使用迁移指南、v1 显式迁移器和 package major 版本标记管理，不保留隐式 runtime legacy fallback。
- **脏工作区风险**：不得 `git reset --hard`、`git checkout --`、删除不熟悉文件或批量格式化；每次验证需解释失败是否可复现于 P0 基线。
- **反射扫描风险**：必须按所有实现接口规范化并按稳定排序；反射异常和重复 key 不得被吞掉或由最后注册覆盖。
- **配置含义风险**：Patch DTO/enum 的 wire format 一经发布难以调整，先以测试真值表和文档 schema 固化，再接入 loader。
- **安全风险**：保留 JSON 深度/大小限制，XML 禁止 DTD 与外部 resolver，别名只接受注册的业务 alias；不得从配置解析或加载 CLR Type。
- **跨框架风险**：新公共 API 必须兼容 netstandard2.0，禁止引入仅新 TFM 可用的反射、序列化或集合 API。
- **性能风险**：不可将 benchmark 结果视为功能测试；需分别测 plan build、cache hit/miss、扫描和 retained memory，禁止无限缓存。

## 6. Definition of Done

1. 四 Profile 形状、单方向 descriptor、三元 Registry key 和统一两项 DI API 已实现并经 Unit/Integration/consumer 验证。
2. JSON/XML 方向节点持有 Profile/ModelAlias；主 loader 不再静默升级 v1；独立迁移器要求调用方明确方向。
3. 五层 precedence 和所有 patch 操作拥有精确真值表、不可变快照和缓存 hit/miss 测试。
4. legacy public APIs 和兼容逻辑已按 v2 删除，PublicApiContractTest、NPOI 泄漏检查和 production IVT 检查更新并通过。
5. 所有本任务 C# 改动符合中文 XML 注释规范，测试方法为英文且具中文测试目的和 AAA。
6. 最新 v2 三包可被 Docs consumer 从隔离本地 feed restore，docs/excel fences 均可编译并运行。
7. Unit 全目标框架、Integration net6/net8、build、pack、docs consumer 和 benchmark/resource evidence 通过或有可审计的环境阻塞说明。
8. 实施过程中未修改/回退无关脏改动，未提交、推送、发布或创建 PR。

## 7. 执行交接

按顺序执行 `P0-001 -> P0-002 -> Phase 1 -> Phase 2 -> Phase 3 -> Phase 4 -> P0-501`。每完成一项任务，在执行记录中写入实际修改文件、测试命令、通过数、残留风险和上述追溯矩阵更新；不要修改独立 review 文档。开始实施时使用本任务 ID 进入 `execute-plan`，不得在本 Planner 回合继续编码。
