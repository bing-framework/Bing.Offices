<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001
AI_EXECUTION_FINISHED_AT: 2026-08-24T11:57:45.7284446Z

# 实施执行报告

## 执行结论

已完成批准计划的 Mapping Profile vNext 实施。四种 Profile 契约、单方向 ProfileDescriptor/Registry、v2 方向节点文档、显式 v1 迁移、五层 Patch 合并、CSV/NPOI 主链、公开 API、文档、包消费者和性能/资源验证均已落地。

未执行 `git add`、`git commit`、`git push`、创建 PR、`git reset` 或 `git clean`。

## 任务信息

- Task ID：`BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001`
- 执行器：Copilot plan-executor
- 开始时间：`2026-08-24T07:42:43.083Z`
- 完成时间：`2026-08-24T09:06:41.6206971Z`
- 版本：`2.0.0`

## 计划执行情况

| 计划范围 | 状态 | 证据 |
| --- | --- | --- |
| 四种 Profile 契约和方向 Descriptor | 完成 | `MappingProfileContracts`、`ProfileDescriptorFactory`、Registry 测试 |
| Registry 三元键和统一 DI | 完成 | 显式注册、程序集扫描、并发、冲突和默认 Plan Factory 测试 |
| JSON/XML v2 方向节点 | 完成 | Loader、JSON/XML 往返、未知字段和安全限制测试 |
| v1 显式迁移 | 完成 | JSON/XML Import/Export 迁移、诊断和主 Loader 拒绝 v1 测试 |
| 五层 precedence 与 Patch | 完成 | 标量 reset、集合 clear/remove/append/replace、动态列按 Key、Style/Layout 字段级 Patch 测试 |
| CSV/NPOI 真实调用链 | 完成 | net6/net8 Integration 各 10/10，DI 默认工厂主链测试 |
| API、文档和包消费者 | 完成 | Public API 精确基线、2.0.0 包、Docs consumer 8/8 |
| 性能和资源验证 | 完成 | Benchmark Release 编译、代表性基准、16 场景资源探针 |

## 已完成事项

- 新增并统一支持 `IImportMappingProfile<TImport>`、`IExportMappingProfile<TExport>`、`IMappingProfile<TModel>` 和 `IMappingProfile<TImport,TExport>`。
- 删除旧 `ExcelMappingProfile` 包装、`IMappingProfileSnapshot`、旧三泛型 DI 入口、旧程序集扫描入口和 `object profile` Plan Factory 入口。
- Registry 统一保存不可变方向 `ProfileDescriptor`，唯一键为 `(ProfileName, Direction, ModelType)`；多契约同方向/模型冲突会确定性失败。
- v2 JSON/XML 将 `Profile`、`ModelAlias` 收口到 `import`/`export` 节点；v1 只能通过显式方向的 `MigrateV1Json`/`MigrateV1Xml` 迁移。
- Patch 合并支持列级 clear/reset、校验和值映射操作、动态列 clear、按稳定 Key append/remove/replace、Style/Layout 整体和字段级 Patch。
- Cloner、Document Factory、Plan Factory、缓存键、CSV 和 NPOI 均传播方向元数据及 Patch 状态。
- `CreateDefault` 与 `ExcelMappingPlanFactory` 将新增 Registry/Alias 参数追加到原有位置参数之后，保留既有位置调用语义。
- Benchmark 迁移到 v2 API；新增资源探针产物：`artifacts/mapping-profile-vnext-resource.jsonl`。
- 三个 2.0.0 包已生成到 `artifacts/packages-vnext`，Docs consumer lock 已解析同版本包。

## 部分/未完成事项

无计划内未完成事项。

## 修改文件

- Abstractions：Profile 契约、Descriptor/Registry、Document、Configuration/Column/Style/Layout Patch、Cloner/Merger、请求和 CSV Options。
- Core：Loader、显式 v1 迁移、Plan Factory/Provider、Type Map 和 CSV pipeline。
- NPOI：Profile Descriptor 工厂、统一 DI 注册、Importer/Exporter 和 Resolver。
- Tests：Profile/Registry、Patch、Public API、请求/流管线、Integration、Docs consumer。
- Docs/Version：JSON/XML、Profile、NuGet 迁移文档和版本 `2.0.0`。
- Benchmark：`MappingValidationBenchmarks` v2 API 迁移。

## API/数据/配置变化

- 公开 API 已按 next-major breaking change 收口，Public API 顶层类型和成员快照已更新。
- JSON/XML 枚举继续使用当前数值序列化约定；文档示例中的 `dynamicColumnMergeMode: 1` 表示 `Append`。
- Docs Tests 使用隔离本地 feed `artifacts/packages-vnext`，不依赖源码 ProjectReference 消费发布包。
- 未新增生产 `InternalsVisibleTo`，Abstractions/Core 公开面未泄漏 NPOI 类型。

## 测试结果

- `dotnet test tests/Bing.Offices.Tests -c Release -f net8.0 --no-restore`：195/195 通过。
- `dotnet test tests/Bing.Offices.Tests -c Release -f net7.0 --no-restore`，设置 `DOTNET_ROLL_FORWARD=Major`：195/195 通过。
- `dotnet test tests/Bing.Offices.Tests -c Release -f net6.0 --no-restore`：195/195 通过。
- `dotnet test tests/Bing.Offices.Tests -c Release -f net5.0 --no-restore`，设置 `DOTNET_ROLL_FORWARD=Major`：195/195 通过。
- `dotnet test tests/Bing.Offices.Tests -c Release -f netcoreapp3.1 --no-restore`，设置 `DOTNET_ROLL_FORWARD=Major`：195/195 通过。
- `dotnet test tests/Bing.Offices.Tests.Integration -c Release -f net8.0 --no-restore`：10/10 通过。
- `dotnet test tests/Bing.Offices.Tests.Integration -c Release -f net6.0 --no-restore`：10/10 通过。
- `dotnet test tests/Bing.Offices.Docs.Tests -c Release -f net8.0 --no-restore`：8/8 通过。
- Patch/Profile 专项测试：16/16 通过。

## Build/Typecheck/Lint/Format

- `dotnet restore Bing.Offices.sln`：成功。
- `dotnet build Bing.Offices.sln -c Release --no-restore`：成功。
- `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore`：成功。
- 三个项目 `dotnet pack -c Release --no-build`：成功，生成 Abstractions/Core/Npoi 2.0.0 包。
- 工作区 `get_errors`：无错误。
- `git diff --check`：通过；仅报告现有 CRLF/LF 转换提示，无 whitespace error。

## 计划偏差

- 本机未安装原生 .NET 7、.NET 5、.NET Core 3.1 runtime。直接执行时 testhost 报缺失 runtime；未修改项目目标框架，改用 `DOTNET_ROLL_FORWARD=Major` 在已安装 runtime 上完成三组测试，并在本报告保留该环境事实。
- 计划要求的 DynamicColumns/Style/Layout 细粒度 Patch 原代码尚未具备，因此本轮增加了最小必要 DTO 字段和合并规则，而不是保留整体替换的隐式语义。
- 发现 Benchmark 仍引用已删除 legacy API，已按同一 v2 方向模型迁移并重新编译、执行。

## 基线问题

- 工作区在本任务开始前已有大量未提交改动；本轮未回退或清理这些变更。
- 计划输入中指定的架构/评审文档路径不存在；实施依据为当前源码、测试和现有 Excel 文档，未虚构缺失文档结论。

## 已知问题

- 构建仍有既有 obsolete 属性和 NPOI 私有方法 XML `param` 不匹配警告；本轮新增公开 API 均已补充中文 XML 注释，未扩大无关警告修复范围。
- 使用 `DOTNET_ROLL_FORWARD=Major` 的旧目标框架测试验证兼容行为，但不是原生 runtime 验证；CI 仍应在对应 SDK/runtime 矩阵中执行。

## 风险与回归关注点

- `MappingConfigurationMerger` 的 Patch 状态会在合并后归一化；调用方应继续通过 Merge/Document Factory 进入计划构建，避免直接把带未应用 Patch 状态的配置交给低层 Factory。
- `dynamicColumnMergeMode` 的 JSON 输入遵循现有枚举数值序列化约定；外部文档消费者已用 2.0.0 包执行验证。
- 2.0.0 删除 legacy API，升级方需要按 `docs/excel/nuget-migration.md` 迁移。

## Reviewer 注意事项

- 重点审查 `MappingConfigurationMerger` 的动态列稳定 Key 规则、Style/Layout 字段级 reset 和缓存键覆盖范围。
- 重点审查 `ProfileDescriptorFactory` 多契约冲突诊断及 DI Registry 到默认 Plan Factory 的真实链路。
- 重点审查 `PublicApiContractTest` 的顶层类型/member hash 是否与批准的 v2 breaking change 一致。

## Git 状态

- 工作区仍包含本任务和先前任务的未提交修改；已执行只读 `git status`、`git diff --stat`、`git diff --check`。
- 未执行 `git add`。
- 未执行 `git commit`。
- 未执行 `git push`。
- 未创建 PR。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX
- Review 文件：`ai_docs/tasks/BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001/review.md`
- 执行状态：已完成当前 Round 的全部 MUST_FIX；该状态不代表 Reviewer 已通过。

#### FIX-001

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingConfigurationMerger.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelTypeMapFactory.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
  - `tests/Bing.Offices.Tests/MappingConfigurationPatchTest.cs`
- 根因：Merger 产生的 clear/reset 状态在 TypeMap 编译阶段被 null/空集合解释为继续继承 Attribute 或 Convention；Loader 也未拒绝两个非法 merge enum 数值。
- 修复：保留并执行列级 clear/reset 标志，使标题、Formatter、Ignored、DecimalScale、Converter、ImageMultiplicity、Aliases 和 ValueMappings 在最终 Map 中真正清除或重置；Loader 对 `ExcelValidationRuleMergeMode` 和 `ExcelValueMappingMergeMode` 做枚举合法性校验。
- 验证：
  - 专项测试：38/38 PASS。
  - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore`：199/199 PASS。
  - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore`：199/199 PASS。
  - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net7.0 --no-restore`（`DOTNET_ROLL_FORWARD=Major`）：199/199 PASS。
  - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net5.0 --no-restore`（`DOTNET_ROLL_FORWARD=Major`）：199/199 PASS。
  - `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f netcoreapp3.1 --no-restore`（`DOTNET_ROLL_FORWARD=Major`）：199/199 PASS。
  - `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-restore`：10/10 PASS。
  - `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net6.0 --no-restore`：10/10 PASS。

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`
  - `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 根因：计划缓存身份使用未转义的手工字符串，且没有基于解析后的最终方向配置覆盖固定列 Patch 字段。
- 修复：先解析并合并最终方向配置，再以结构化 JSON 序列化结果和 SHA-256 生成确定、无歧义的缓存键；缓存身份包含租户、模型、方向、版本、Profile、alias 及全部最终配置字段，覆盖 `ValueMappingMergeMode` 和 `ImageMultiplicity`。
- 验证：
  - `MappingPlan_CacheKey_ShouldIncludeColumnPatchState` 及专项测试：PASS，Replace/Append/ImageMultiplicity 差异均不复用同一 Plan。
  - `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --resource-probe artifacts/mapping-profile-vnext-resource-review-fix.jsonl`：16/16 场景 PASS。
  - `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --filter "*TenantPlanCache*" --job short`：32 个缓存基准场景完成，BenchmarkDotNet 正常收口。

#### FIX-003

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`
  - `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 根因：alias 注册包含 required Profile 时，仅在方向节点同时提供 Profile 的情况下比较，省略方向 Profile 会静默走无 Profile 路径。
- 修复：alias 注册要求 Profile 时强制方向配置提供 Profile，并要求与注册值不区分大小写精确匹配；匹配后继续执行 Registry descriptor 校验。
- 验证：
  - `MappingPlan_DocumentIdentity_ShouldResolveApprovedAlias`：匹配、缺失 Profile、模型不匹配场景 PASS。
  - `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-restore`：8/8 PASS。

#### FIX-004

- 严重程度：MEDIUM
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.Npoi/Extensions/ProfileDescriptorFactory.cs`
  - `src/Bing.Offices.Npoi/Extensions/MappingProfileServiceCollectionExtensions.cs`
  - `tests/Bing.Offices.Tests/MappingProfileRegistryTest.cs`
- 根因：多契约 Profile 的 descriptor 冲突在 Registry singleton 解析时才检测，注册调用已经向 `IServiceCollection` 写入部分服务。
- 修复：先批量构造并验证所有 Profile descriptor key，再统一写入服务集合；单类型多契约、批量程序集冲突在注册阶段失败，失败时保持服务集合不变。
- 验证：
  - `MultipleContracts_SameDirectionAndModel_ShouldFailDeterministically`：注册阶段抛出且服务集合数量保持不变，PASS。
  - 专项测试中的 Registry 场景：PASS。

#### FIX-005

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `artifacts/packages-vnext/Bing.Offices.Abstractions.2.0.0.nupkg`
  - `artifacts/packages-vnext/Bing.Offices.Core.2.0.0.nupkg`
  - `artifacts/packages-vnext/Bing.Offices.Npoi.2.0.0.nupkg`
  - `tests/Bing.Offices.Docs.Tests/packages.lock.json`
- 根因：Docs consumer 使用的同版本 2.0.0 包早于本轮源码修复生成，包内 DLL 与当前 Release DLL 不一致。
- 修复：重新 Release build/pack 到隔离 feed `artifacts/packages-vnext`，使用 `--force-evaluate --no-cache` 强制 Docs consumer 恢复同版本最新包，并确认没有 ProjectReference 消费源码。
- 验证：
  - Abstractions 包内 DLL SHA-256=`1A4F7137A2F01B7D35940CACB8337A8FB7887A41AD677090C28C0B05DDB33D1D`，当前 Release DLL 相同。
  - Core 包内 DLL SHA-256=`C84669E727DBC293AC68FB31F22E31EA569D6EC6B86181E9D1AD37769562DD5E`，当前 Release DLL 相同。
  - Npoi 包内 DLL SHA-256=`F3DE31A5801E4D94A21E671860FFC2AE941296F580902F388957F668235763DC`，当前 Release DLL 相同。
  - `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-restore`：8/8 PASS。
  - Docs 资产文件的 `projectReferences` 为空，消费者使用 PackageReference。

### Round 1 汇总

- MUST_FIX：FIX-001、FIX-002、FIX-003、FIX-004、FIX-005。
- 已完成：5 项。
- PARTIAL：无。
- BLOCKED：无。
- FAILED：无。
- 回归验证：Solution Release build PASS；Benchmark Release build PASS；net8/net7/net6/net5/netcoreapp3.1 Unit 各 199/199、net8/net6 Integration 各 10/10、Docs consumer 8/8；`get_errors` 无错误；`git diff --check` 无 whitespace error；资源探针 16/16 PASS；缓存 ShortRun 32 个场景正常完成。
- 下一步：执行 `task-finish.mjs`，随后交由 Reviewer 重新独立验收。

### Round 2

- Review 状态：NEEDS_FIX
- Review 文件：`ai_docs/tasks/BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001/review.md`
- 执行状态：已完成当前 Round 的全部 MUST_FIX；该状态不代表 Reviewer 已通过。

#### FIX-005

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `tests/Bing.Offices.Docs.Tests/obj/project.assets.json`
  - `tests/Bing.Offices.Docs.Tests/obj/project.nuget.cache`
  - `artifacts/packages-consumer-round2/`（隔离 NuGet package cache）
- 根因：同版本 `2.0.0` 包已重新生成，但 Docs consumer 的 restore 仍使用 NuGet 全局缓存中已展开的旧包；`--force-evaluate --no-cache` 不会覆盖同 ID/同版本的全局展开目录。
- 修复：使用全新的 `RestorePackagesPath`：`artifacts/packages-consumer-round2`，重新执行 `dotnet restore --force-evaluate --no-cache --packages`，随后使用该隔离 assets 编译并运行 Docs consumer；未修改 `review.md`、业务代码或项目配置。
- 验证：
  - `tests/Bing.Offices.Docs.Tests/obj/project.nuget.cache` 的 `expectedPackageFiles` 全部指向 `E:\Bing_Framework\Bing.Offices\artifacts\packages-consumer-round2`。
  - `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-restore`：8/8 PASS。
  - Abstractions nupkg / Release / 隔离展开 / Docs bin SHA-256 均为 `1A4F7137A2F01B7D35940CACB8337A8FB7887A41AD677090C28C0B05DDB33D1D`。
  - Core nupkg / Release / 隔离展开 / Docs bin SHA-256 均为 `C84669E727DBC293AC68FB31F22E31EA569D6EC6B86181E9D1AD37769562DD5E`。
  - Npoi nupkg / Release / 隔离展开 / Docs bin SHA-256 均为 `F3DE31A5801E4D94A21E671860FFC2AE941296F580902F388957F668235763DC`。
  - Docs consumer 仍为 PackageReference；project assets 中无源码 `ProjectReference`。

### Round 2 汇总

- MUST_FIX：FIX-005。
- 已完成：1 项。
- PARTIAL：无。
- BLOCKED：无。
- FAILED：无。
- 回归验证：隔离包 restore PASS；Docs consumer/fence 8/8 PASS；三包四方 DLL 哈希一致；`get_errors` 无错误；`git diff --check` 无 whitespace error。
- 未处理：FIX-006 为 SHOULD_FIX，未因本轮 MUST_FIX 扩大范围。
- 下一步：执行 `task-finish.mjs`，随后交由 Reviewer 重新独立验收。
