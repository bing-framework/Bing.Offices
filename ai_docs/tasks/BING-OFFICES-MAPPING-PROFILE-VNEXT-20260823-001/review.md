<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001
AI_REVIEWED_AT: 2026-08-24T13:56:01.8218797Z

# 独立复审报告

## 验收摘要

最终结论：`PASS_WITH_ISSUES`。

本轮是 Review Fix Round 2 后的独立复审，优先逐项验证上一轮 `FIX-001` 至 `FIX-005`。`FIX-001` 至 `FIX-004` 的生产实现和直接回归测试未发生回归；`FIX-005` 已通过全新的独立 NuGet 包目录重新验证。Docs consumer 从 `artifacts/packages-consumer-review3` 恢复后，8/8 测试通过，且三个包的 nupkg、统一 Release 输出、独立展开目录与 Docs 输出 DLL 的 SHA-256 全部一致。

上一轮保留的 `FIX-006` 是 `SHOULD_FIX`，涉及更完整的真实集成、文档名称和 Benchmark 矩阵，不构成当前计划交付的阻断项。本轮未发现新的 `BLOCKER` 或 `HIGH` 问题。Reviewer 未修改业务代码、测试、`plan.md` 或 `execution.md`。

## Review 边界

- 计划覆盖 Abstractions、Core、NPOI、Unit/Integration/Docs Tests、文档、版本、包消费者与 Benchmark。
- 工作区在任务开始前已有大量未提交改动；当前 Diff 仍无法以 Git 基线完全区分预存改动与本任务改动。本轮未回退或清理任何文件。
- 当前变更主题与 Mapping Profile vNext 一致。任务目录、生成包和 NuGet restore 资产属于本任务验证证据；不将其误判为业务源码改动。
- 本轮重新读取 `plan.md`、当前 `execution.md`、旧 `review.md`、Docs 项目配置/资产、相关生产实现和回归测试，而非仅采信执行报告。

## 上一轮 FIX 复验

| FIX | 复验状态 | 结论与证据 |
| --- | --- | --- |
| FIX-001 | RESOLVED | `MappingConfigurationMerger` 继续传播列级 clear/reset 状态；`ExcelTypeMapFactory` 路径由 `MappingConfigurationPatchTest` 直接验证 Profile 与 Attribute 低优先级字段被清除。Loader 对非法 Validation/ValueMapping merge enum 仍拒绝。net8 Unit 199/199 通过。 |
| FIX-002 | RESOLVED | `ExcelMappingPlanFactory` 在解析最终方向配置后，以结构化 JSON 和 SHA-256 生成缓存键；租户、模型、方向、配置版本和最终 Configuration 均进入身份。`MappingPlan_CacheKey_ShouldIncludeColumnPatchState` 继续验证 Replace、Append、ImageMultiplicity 不复用 Plan。 |
| FIX-003 | RESOLVED | alias 注册要求 Profile 时，`ResolveDocument` 对方向节点缺失 Profile 和 Profile 不匹配均抛出确定异常。`MappingPlan_DocumentIdentity_ShouldResolveApprovedAlias` 覆盖匹配、模型不匹配和缺失 Profile。 |
| FIX-004 | RESOLVED | `AddMappingProfile`/`AddMappingProfiles` 在写入 `IServiceCollection` 前先构造并校验全部 registration key；`MultipleContracts_SameDirectionAndModel_ShouldFailDeterministically` 直接断言注册失败后集合数量不变。 |
| FIX-005 | RESOLVED | 初始 `project.nuget.cache` 曾指向全局同版本包缓存，因此未直接沿用旧结果。本轮使用 `dotnet restore --force-evaluate --no-cache --packages artifacts/packages-consumer-review3` 独立恢复；更新后的 assets/cache 指向该目录，Docs 8/8 通过。Abstractions、Core、Npoi 的 nupkg、Release、独立展开和 Docs 输出哈希各自四方一致。 |
| FIX-006 | PARTIAL | 仍是上一轮定义的 `SHOULD_FIX`。未发现其在本轮造成发布包、主链或已验证行为失败；测试/文档/Benchmark 完整性缺口保留为非阻断问题。 |

## 计划验收矩阵

| 计划项 | 结论 | 当前证据 |
| --- | --- | --- |
| P0-001 基线隔离 | PARTIAL | 脏工作区已在实施/复审记录中说明，但缺少可机器比较的逐文件起始快照。 |
| P0-002 v2 API/迁移契约 | PASS | 四种契约、legacy 删除与显式 v1 迁移仍由 Unit/Docs consumer 编译和运行验证。 |
| P0-101/P0-102 Profile、Descriptor、Registry | PASS | 四形状 descriptor、三元 key、并发读取、冲突和 alias 约束的 Unit 继续通过。 |
| P0-201/P0-202 方向节点与显式迁移 | PASS | JSON/XML v2、方向 metadata 与明确 Import/Export v1 迁移路径仍由回归和 Docs 测试覆盖。 |
| P0-203 Patch 与缓存 | PASS | clear/reset 贯通最终 Plan，最终配置参与结构化缓存身份。 |
| P0-301/P0-302 调用链与统一 DI | PASS | 默认 Plan Factory 的 DI 主链和原子注册失败测试通过；net8 NPOI Integration 10/10 通过。 |
| P1-303 API/程序集边界 | PASS | 编辑器静态检查无错误；未发现本轮修复引入 NPOI 反向公开泄漏或生产 IVT。 |
| P0-401 Unit/Integration 矩阵 | PARTIAL | 本轮 net8 Unit 199/199、Integration 10/10。四形状与 Patch 的真实 CSV/XLSX 覆盖仍可加强。 |
| P0-402 包消费者与文档 | PASS | Docs 项目仅含 PackageReference，无 ProjectReference；独立 restore 后 Docs/fence 8/8 且实际输出 DLL 与 feed 包一致。 |
| P1-403 性能/资源 | PARTIAL | 现有执行报告记录资源探针和缓存基准；计划矩阵中的四形状 descriptor、五层 merge、独立 cache hit/miss 基准仍不完整。 |
| P0-501 发布前收敛 | PASS | 本轮有独立 package-consumer 证明，Unit、Integration、Docs 和静态检查均成功；旧框架与基准的执行记录保留为历史证据。 |

## 功能、API 与架构 Review

- Patch clear/reset 已从 Merger 贯通到 TypeMap/Plan，不再回退 Attribute/Profile 默认值。
- Plan Factory 先解析方向配置与 Profile，再计算结构化缓存身份；原固定列字段遗漏和分隔符拼接风险已消除。
- alias required Profile 缺失时不再静默走 Convention/Attribute 路径。
- DI 类型级 key 冲突在服务写入前检测；Profile 构造器依赖仍按计划留在 Provider 阶段。
- `Abstractions <- Core <- NPOI` 依赖方向未改变；缓存仍为有界 `ConcurrentDictionary` 加队列淘汰，未引入无界常驻缓存。

## 包消费者专项证据

| 包 | nupkg / Release / 独立 restore / Docs 输出 SHA-256 |
| --- | --- |
| `Bing.Offices.Abstractions` | `1A4F7137A2F01B7D35940CACB8337A8FB7887A41AD677090C28C0B05DDB33D1D` |
| `Bing.Offices.Core` | `C84669E727DBC293AC68FB31F22E31EA569D6EC6B86181E9D1AD37769562DD5E` |
| `Bing.Offices.Npoi` | `F3DE31A5801E4D94A21E671860FFC2AE941296F580902F388957F668235763DC` |

- Docs `project.nuget.cache` 的 Bing.Offices package entries 指向 `artifacts/packages-consumer-review3`，而非全局 NuGet 包目录。
- Docs `project.assets.json` 将三包标记为 `type: package`；项目文件也没有 `ProjectReference`。
- 同版本本地包会受全局展开缓存影响，因此发布/CI 复验必须继续使用唯一版本或独立 `RestorePackagesPath`。本轮验证采用后一种方式。

## 测试与验证

- `dotnet restore tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj --force-evaluate --no-cache --packages artifacts/packages-consumer-review3`：成功。
- `dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-restore`：8/8 通过。
- `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build --no-restore`：199/199 通过。
- `dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-restore`：10/10 通过。
- `get_errors`（src/tests/benchmarks）：无错误。
- `git diff --check`：无 whitespace error；仅输出既有 CRLF/LF 转换提示。

## 剩余问题

### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 当前状态：OPEN
- 对应计划项：P0-401、P0-402、P1-403、P0-501
- 问题：真实 Integration 尚未直接覆盖四形状 Profile、Patch clear/reset 的 CSV/XLSX 路径；Docs 的方向文档示例使用短 Profile 名称但默认 DI 注册使用 `FullName`；`nuget-migration.md` 仍保留 `MIGRATION_CURRENT_MAJOR`；Benchmark 未完全覆盖计划指定的 descriptor、五层 merge 和独立 cache hit/miss 场景。
- 影响：这些缺口降低测试、文档和性能证据的完整性，但本轮没有证明其导致发布包或已验证主链行为错误。
- 建议：在后续独立改进任务中补充代表性真实集成、包端默认 Plan Factory、可解析 Profile 命名文档、版本占位符替换，以及对应 Benchmark 场景。
- 验证方式：Integration net6/net8、隔离包 Docs Tests、Markdown fences、Benchmark Release 编译与代表性场景。

## 最终验收 Checklist

- [x] 已读取当前 plan、execution、旧 review 和项目规则。
- [x] 已优先复验 FIX-001 至 FIX-005。
- [x] FIX-001 至 FIX-004 均保持修复状态。
- [x] FIX-005 已由独立 NuGet 包目录、Docs 8/8 和四方哈希闭环验证。
- [x] Docs consumer 无源码 ProjectReference。
- [x] net8 Unit、net8 Integration、Docs consumer 和静态检查通过。
- [x] `git diff --check` 无 whitespace error。
- [ ] FIX-006 的增强性测试、文档和 Benchmark 完整性尚未收敛（SHOULD_FIX）。

不存在未解决的 `MUST_FIX`，因此最终结论为 `PASS_WITH_ISSUES`。Reviewer 不进入修复阶段。
