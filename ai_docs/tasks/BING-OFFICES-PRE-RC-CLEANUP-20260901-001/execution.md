<!-- AI_EXECUTION_STATUS: BLOCKED -->
AI_TASK_ID: BING-OFFICES-PRE-RC-CLEANUP-20260901-001
AI_EXECUTION_FINISHED_AT: 2026-09-03T00:58:56.3600735Z

# 实施执行报告

## 执行结论

状态：`BLOCKED`。发布判定：`No-Go`。

本轮持续执行已批准的 `BING-OFFICES-PRE-RC-CLEANUP-20260901-001`，完成用户批准的正式 API baseline 收口及可运行 TFM 回归。API compare、net8 Unit 和 net6 Unit 均已通过；整体仍因缺失 TFM runtime、完整性能/资源门禁、长路径 package consumer、剩余 API 治理和 clean-clone 交付证据保持 `BLOCKED/No-Go`。未修改 `review.md`。

## 正式 API baseline 收口批次（2026-09-03）

- 用户批准：将当前 Release API 纳入正式 baseline；批准范围为已记录并完成迁移验证的 Breaking Change 和 API 收敛，不扩展到尚未批准的剩余 execution detail 治理。
- 输入构建：`dotnet build .\Bing.Offices.sln -c Release --no-restore`，退出码 `0`，`0 error / 15 warning`；SDK `10.0.400`。
- 当前 Release 快照：`artifacts/api-snapshot-formal-20260903/api-snapshot-*.json`。
- 正式 baseline：`artifacts/api-snapshot-formal-baseline-20260903.json`。
- API compare：`dotnet run --project .\build\ApiSnapshot\ApiSnapshot.csproj -c Release --no-build -- --root .\output\release --baseline .\ai_docs\tasks\BING-OFFICES-PRE-RC-CLEANUP-20260901-001\artifacts\api-snapshot-formal-baseline-20260903.json --output .\ai_docs\tasks\BING-OFFICES-PRE-RC-CLEANUP-20260901-001\artifacts\api-snapshot-formal-20260903`，退出码 `0`；`netcoreapp3.1`、`net6.0`、`net8.0` 全部通过。
- 正式成员 hash：Abstractions `7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943`（723 members）；Core `B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6`（194 members）；NPOI `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`（1 member）。三个受控 TFM 的逻辑程序集快照一致。
- 与 Round 5 快照的逐成员差异仅为已批准的 Settings 删除、Core execution detail internal 化及 Provider `cacheCapacity` 尾参数收敛；未发现未批准的公共 API 新增。
- Unit net8：`384 total / 384 passed / 0 failed`，TRX：`tests/Bing.Offices.Tests/TestResults/api-baseline-net8-final-rerun.trx`。
- Unit net6：`384 total / 384 passed / 0 failed`，TRX：`tests/Bing.Offices.Tests/TestResults/api-baseline-net6-final-rerun.trx`。
- 最终 Release DLL 身份：Abstractions SHA-256 `3F3A6E9740A1B55B95738ED56CC731B1FD50B0FBFBB612714747A82C7A74D409`，Core `C9DA149B5C7D82E60C075C182597DE79FBD0CEE978EB9BA843F95AA0E6D8F79F`，NPOI netcoreapp3.1 `A1BA3472FE3A5C04A96CFF2BF456C0361B96EC3C4E5B78533EF0A30A1F52135D`、net6 `0C26DF3310A5ADA8176B8F852275838C5BB5D69348212A5FF3DB1CA38A5A87CE`、net8 `3BBDB9551E5C8C050D797F5139196935E983157915CFD885541BC2A82C1B2D1E`。
- 本批次关闭正式 API baseline 和可运行 TFM Unit hash 阻断；不改变整体 RC 判定。

## Round 6 追加证据

- 执行范围：`P0-02`、`P2-01`、`P2-02`、`P3-01`、`P3-02`、`P3-03`。
- 生产代码已完成编译级修复：NPOI Sheet executor 的动态列命名空间/转换器集合、NPOI 默认 Plan provider 引用，以及 Benchmark/ResourceProbe 对 internal Plan factory 的公开 Provider SPI 迁移。
- `dotnet build Bing.Offices.sln -c Release --no-restore`：解决方案全量构建成功，包含 Abstractions、Core、NPOI（net8.0/net6.0/netcoreapp3.1）、Tests（net8.0/net6.0/netcoreapp3.1）、Integration、Docs、Benchmark、ResourceProbe；结果为 `0 error / 28 warning`。
- Benchmark `Program.cs` 与 `MappingValidationBenchmarks.cs` 已从 internal `ExcelMappingPlanFactory` 迁移至 `ExcelMappingPlanFactoryProvider.CreateDefault()`；剩余警告主要为 netcoreapp3.1 依赖支持提示、已有 nullable/obsolete/analyzer 警告，未将其伪装为零警告。
- API formal baseline 已在本执行回合按用户批准更新并通过 compare/Unit；`review.md` 未修改；RC 仍保持 `No-Go`。

## 任务信息

- Task ID：`BING-OFFICES-PRE-RC-CLEANUP-20260901-001`
- 执行模式：plan-execution
- 执行器：Copilot
- 计划：`plan.md`
- 独立 Review：`review.md`，状态 `BLOCKED`（本轮保持不变）
- 发布清单：`release-checklist.md`
- 最终报告：`final-report.md`

## 计划执行情况

| 阶段 | 状态 | 说明 |
| --- | --- | --- |
| Phase 0 基线与初始门禁 | `PARTIAL` | 环境、项目、输入、Release build、正式 API compare、可运行测试和 pack 已取证；缺失 runtime 和交付复现仍阻断 |
| Phase 1 正确性/安全/资源 | `PARTIAL` | P0/P1 业务修复和 Failure Workbook 已验证；ResourceProbe 不能证明任意 DOM/Failure Workbook 硬内存上限 |
| Phase 2 API 删除/收敛 | `PARTIAL` | 已批准 Breaking Change 已落地并完成正式 baseline/成员级 diff；剩余 legacy/public execution detail 治理未完成 |
| Phase 3 目录/职责/异步重构 | `TODO/PARTIAL` | 大规模拆分和完整 ADR 未执行，避免在发布门禁未收口时扩大行为风险 |
| Phase 4 测试/包证据 | `PARTIAL` | Integration/Docs/专项、正式 API compare 和可运行 TFM Unit 有效；长路径 consumer 和缺失 TFM 仍有限制 |
| Phase 5 Benchmark/资源 | `PARTIAL` | BDN `9/9`、ResourceProbe `16/16` 和尾延迟已记录；完整矩阵与正式预算未批准 |
| Phase 6 文档/Review/RC | `BLOCKED` | `review.md`、`release-checklist.md`、`final-report.md` 已创建；No-Go 阻断仍未解除 |

## 已完成事项

### 生产与测试修复

- Mapping Document request configuration 按方向合并并进入最终 Plan；
- v1 JSON/XML 迁移的非目标方向保持 `null`，保留 XML DTD/外部实体防护和流所有权合同；
- Relation Binder 解包 `TargetInvocationException` 并重抛原异常；
- CSV 公式防护覆盖 BOM、控制字符、ASCII/Unicode whitespace，保留 `Escape`/`None` 语义和负数行为；
- `CsvImportOptions` 增加 unique limit/comparison 边界校验；
- ResourceProbe 改为独立 child process，不预先打开同一 Workbook；
- `AddNpoi` 删除，唯一 DI 入口为链式 `AddBingOfficesNpoi`；
- 未知 Workbook/Sheet 能力改为 `NotSupportedException`；
- `MaxBytes` 更名为 `MaxSerializedBytes`，完成 Failure Workbook 输出限制、取消、临时文件、清理诊断和目标流保护测试；
- 直接验证 NPOI 2.7.4 provider：HSSF 的 `Hidden`/`Collapsed` 不支持而 `ZeroHeight` 支持，XSSF 三项均支持；三处窄范围 capability fallback 合法保留。

### 已取得的运行证据

- Release solution build：成功；
- Integration net6/net8：`30/30`；
- Docs net8：`11/11`；
- Failure Workbook 专项：`14/14`；
- DI 定向回归：`6/6`；
- Excel ResourceProbe：`7/7` child scenarios；mapping/unique benchmark ResourceProbe：`16/16`；
- StreamPipeline BenchmarkDotNet：`9/9` ShortRun；
- package-only consumer：`artifacts/package-consumer-rerun2` 无 `ProjectReference`，使用 Round 5 `packages-round5` 的本地 2.0.0 nupkg 和短路径 `C:\nupkg-cache-round5` 时 restore/build/run 均退出码 `0`，输出 `package-consumer-ok`。

## 部分/未完成事项

1. **缺失 TFM runtime。** netcoreapp3.1、net5.0、net7.0 runtime 未安装，相关 Unit 未运行。
2. **性能门禁。** ShortRun 仅为证据；完整 Benchmark 矩阵未形成，尾延迟 `budgetStatus=UNAPPROVED`，100,000 行 Import 记录 `3` 个异常需复核。
3. **资源边界。** 已补充七个真实 Excel child-process 场景，并保留 mapping-plan/unique-tracker 子进程场景；仍未覆盖任意真实 XLS/XLSX、压缩输入和 Failure Workbook 双 DOM，不能宣称硬内存上限。
4. **Package consumer 环境。** 短路径缓存成功；任务深路径缓存 restore 后 build 触发 SDK/MSBuild `MSB3106`，随后出现 DI Abstractions 解析错误。该环境失败已如实保留，不能写成无条件 consumer 通过。
5. **API 分层/删除候选。** Round 5/6 已完成已批准的 legacy converter、旧 validation attributes、CSV 全局状态/旧隐式重载、Office exceptions、Settings 及部分 Core execution detail 收敛；DataTable 显式兼容类、UniqueTracker 和剩余 public execution detail 仍需逐符号治理与迁移闭环。
6. **P3 重构/ADR。** Import/CSV/Mapping 拆分和完整异步/所有权 ADR 未全部完成，保持 TODO/partial。

## 修改文件

生产代码、测试和 Benchmark/Probe 修改集中于计划范围；本轮继续补充/修正以下任务证据：

- `package-consumer-report.md`：改为最新 `package-consumer-rerun2` 证据，补充 nupkg SHA-256、短路径成功和长路径 `MSB3106`；
- `progress.md`、`test-matrix.md`、`unit-test-report.md`：同步 consumer 和 capability 审计结果；
- `decisions.md`、`deprecated-removal.md`：记录 NPOI capability fallback 决策；
- `review.md`：独立 Review，结论 `BLOCKED`，本轮未修改；
- `release-checklist.md`：RC 发布门禁和解除条件；
- `final-report.md`：最终执行摘要和证据索引；
- `execution.md`：本报告及合法终态元数据。

## API/数据/配置变化

- `ExcelImportFailureOptions.MaxSerializedBytes` 替代 `MaxBytes`，语义仅为 Failure Workbook 序列化输出限制；
- `ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(IServiceCollection)` 替代旧 `AddNpoi`，返回同一 `IServiceCollection`；
- 未对其它候选兼容 API 机械删除，等待逐符号治理和维护者批准；
- 未执行数据库、共享基础设施或生产数据变更。

## 测试结果

| 验证 | 结果 |
| --- | --- |
| Unit net8 | `384/384`，正式 API baseline 匹配 |
| Unit net6 | `384/384`，正式 API baseline 匹配 |
| Integration net8/net6 | `15/15` + `15/15`，合计 `30/30` |
| Docs net8 | `11/11` |
| Failure Workbook 专项 | `14/14` |
| DI 定向回归 | `6/6` |
| Excel ResourceProbe | `7/7` child scenarios，非完整 DOM 安全证明；mapping/unique benchmark probe 另有 `16/16` |
| Package-only consumer | 短路径缓存成功；长路径缓存 `MSB3106` 阻断 |
| 缺失 TFM | netcoreapp3.1/net5/net7 未执行 |

## Build/Typecheck/Lint/Format

- Release solution build：`PASS`；
- `get_errors`：源码、测试和构建相关诊断为 0；
- `git diff --check`：通过，仅有 CRLF/LF 提示；
- 未发现独立 lint 入口，未伪造 lint 通过；
- package-only consumer 短路径 restore/build/run 均为退出码 `0`。

## 计划偏差

- P3 大规模职责拆分未执行：先完成 P0/P1 修复、测试和发布证据，避免在 API/性能门禁未收口时扩大行为风险；
- API formal baseline 已按用户批准更新；成员差异与当前 Release DLL 身份见本报告“正式 API baseline 收口批次”和 `api-diff.md`。剩余未批准 API 治理仍保持独立阻断。
- package consumer 使用短路径 NuGet cache 作为最小可逆环境隔离，同时保留深路径 `MSB3106` 失败证据；
- ResourceProbe 没有扩展为未经验证的 DOM 硬上限声明。

## 基线问题

- 用户指定的 `ai_docs/codebase-analysis/` 评审与方法论文件当前不存在，以源码、项目配置、真实命令和任务 artifacts 为证据；
- 当前工作树不是 clean clone，任务证据和原始 artifacts 尚未形成授权提交输入；
- API formal baseline 已取得用户批准并完成 compare/Unit 验证；
- 缺失 TFM runtime 和长路径 SDK/MSBuild 限制均需外部环境处理。

## 已知问题

- API baseline 收口前的固定 hash 失败已保留为历史证据，当前可运行 TFM Unit 已全绿；
- 多个 legacy/public execution-detail 类型尚未最终收敛；
- Benchmark 矩阵不完整、正式预算未批准；
- ResourceProbe 不覆盖任意 Workbook DOM、压缩输入或 Failure Workbook 双 DOM；
- 长路径 NuGet cache 的 `MSB3106`；
- 100,000 行 Import 的 `3` 个 Benchmark 异常待复核。

## 风险与回归关注点

- 不得把 `MaxInputBytes` 写成解压、NPOI DOM、实体对象图或 Failure Workbook 峰值的硬上限；
- 直接写目标 Stream 失败可能部分写入，File API 仍需验证原子提交；
- 旧 `artifacts/package-consumer` 失败后旧二进制输出不作为成功证据；
- 不得将 ShortRun 或 mapping-only ResourceProbe 写成正式性能/资源门禁；
- API hash 已在批准 Breaking Change 后更新，并已重跑 Unit/API compare；后续 API 变化仍必须重复该门禁；
- 缺失 TFM 安装后需补跑对应测试，不能从 net6/net8 推广结论。

## Reviewer 注意事项

- 复核正式 API baseline 是否与当前 Release DLL、快照和成员差异绑定；
- 确认 API hash 通过来自批准变更和真实 compare，不是忽略或条件跳过；
- 保留 100k Import 的 3 个异常和尾延迟 `UNAPPROVED` 状态；
- 检查 ResourceProbe 不被误读为任意 Workbook 内存硬上限；
- 检查 package consumer 使用最新 nupkg、无 ProjectReference，并区分短路径成功和长路径 `MSB3106`；
- 复核 public execution detail/legacy compatibility 的逐符号删除或保留记录。

## Git 状态

- 未自动执行 `git add`；
- 未自动执行 `git commit`；
- 未自动执行 `git push`；
- 未自动创建 PR、tag 或发布；
- 未执行 `git reset`、`git clean` 或覆盖未知用户修改；
- 当前工作树包含本任务的生产/测试/文档修改和任务证据目录，需由授权人员后续审阅、staging 和提交。

## 任务收口条件

只有在补齐可运行 TFM、批准性能预算并补齐资源边界证据、完成 API 治理和交付复现后，才可重新独立 Review 并判断是否进入 Go。正式 API baseline 已关闭，但当前 `BLOCKED` 终态仍不能进入发布。

## Review 修复记录

### 本次 API baseline 收口

- 用户批准：正式 API baseline 纳入当前 Release。
- 已更新：`tests/Bing.Offices.Tests/PublicApiContractTest.cs`、`api-diff.md`、`unit-test-report.md`、`test-matrix.md`、`progress.md`、`release-checklist.md`、`final-report.md` 和本报告；新增正式 baseline/快照 artifacts。
- 验证：最终 Release build `0 error / 15 warning`；API compare `0`；net8 Unit `384/384`；net6 Unit `384/384`。
- 未修改：`review.md`。
- 结论：正式 API baseline 阻断已关闭；缺失 runtime、性能/资源、package consumer、剩余 API 治理和 clean-clone 仍保持整体 `BLOCKED/No-Go`。

### Round 8

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 执行结论：已完成本轮纳入范围的 `FIX-007`；未修改 `review.md`。执行终态保持 `BLOCKED`，RC 保持 `No-Go`。

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
	- `benchmarks/Bing.Offices.Benchmarks/Program.cs`
	- `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`（Round 8 前已增加回归，未在本次最终验证中重新修改）
	- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`（删除 EOF 多余空行）
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/benchmark-plan.md`
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/benchmark-report.md`
- 根因：租户淘汰基准此前计算了容量但未传入实际 Provider Factory，日志和被测缓存容量不一致；同时历史 Benchmark 结果未能作为修复后数值基线。
- 修复：Provider 通过显式 `cacheCapacity` 创建实际 Factory；`TenantPlanCacheEviction()` 使用同一 `capacity` 创建 Factory、校验容量小于租户数时确实发生重建，并输出同一事实来源的容量。`TenantPlanCache()` 显式固定使用默认容量 `256`。新增 `PropertyAccessorBenchmarks` 纳入已存在 Benchmark 入口，补充 compiled getter/setter 与 reflection 对照证据。
- 直接回归：`MappingPlan_ProviderCacheCapacity_ShouldDriveEviction()` 覆盖容量为 `2` 时的命中、淘汰、重建和租户配置正确性；`MappingPlan_ProviderDefaultCache_ShouldRetainThenEvictPlans()` 覆盖默认容量命中、租户隔离和超容量重建。
- 验证：
	- `dotnet build .\Bing.Offices.sln -c Release --no-restore`：PASS，退出码 `0`，`0 error / 27 warning`；缺失 `netcoreapp3.1` runtime 仍未解除。
	- `dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter "FullyQualifiedName~ReviewFixRegressionTest.MappingPlan_ProviderCacheCapacity_ShouldDriveEviction|FullyQualifiedName~ReviewFixRegressionTest.MappingPlan_ProviderDefaultCache_ShouldRetainThenEvictPlans" --logger "trx;LogFileName=round8-cache-final.trx"`：PASS，`2/2`；TRX：`tests/Bing.Offices.Tests/TestResults/round8-cache-final.trx`。
	- Tenant cache Benchmark：PASS，退出码 `0`；日志：`BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.TenantPlanCacheBenchmarks-20260902-214530.log`；100 tenant 输出 `capacity=50 observed=True`，1000 tenant 输出 `capacity=256 observed=True`；结果：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.TenantPlanCacheBenchmarks-report-github.md`。
	- Property accessor Benchmark：PASS，退出码 `0`；日志：`BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.PropertyAccessorBenchmarks-20260902-214303.log`；结果：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.PropertyAccessorBenchmarks-report-github.md`。
	- `git diff --check`：PASS；仅保留 CRLF/LF 转换提示，无 EOF 空白错误。

### Round 8 汇总

- MUST_FIX：无新增未解决项；历史 `FIX-001` 至 `FIX-003` 继续为 `RESOLVED`。
- SHOULD_FIX：`FIX-007` 已完成。
- OPTIONAL：`FIX-006` 继续 `DEFERRED`，任务目录 Git 跟踪、clean-clone 交付和 `.agents/runtime/current-task.json` 交付问题不在本轮推荐范围内，且未执行 staging 或不可逆 Git 操作。
- PARTIAL/BLOCKED：formal API baseline、缺失 `netcoreapp3.1` runtime、完整性能/资源矩阵、Failure Workbook 双 DOM/压缩输入边界、长路径 package consumer、剩余 public execution detail/API 治理和 clean-clone 证据交付仍阻断 RC；Unit 的 formal API hash 失败继续保留。
- FAILED：无新增功能失败；完整 Unit 的既有 formal API hash mismatch 未被修改或隐藏。
- 回归验证：Release solution build；缓存定向 Unit `2/2`；TenantPlanCache 最终编译版本 Benchmark；PropertyAccessor ShortRun Benchmark；`git diff --check`。
- 下一步：执行 `task-finish` 后交回独立 Reviewer 进行“再次验收”；不得修改 `review.md`，不执行 commit/push。

### Round 7

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`

#### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `benchmarks/Bing.Offices.Benchmarks/MappingValidationBenchmarks.cs`
	- `tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs`
- 根因：租户淘汰基准记录 `TenantCount / 2`，但 Provider 工厂固定使用默认缓存容量 `256`，导致日志容量与实际容量不一致。
- 修复：基准改为记录实际默认容量 `256`；新增 Provider 默认缓存命中、租户隔离和超容量重建的直接回归。未扩张 Provider SPI 公共签名。
- 验证：`MappingPlan_ProviderDefaultCache_ShouldRetainThenEvictPlans` 在 net8 `1/1` 通过；Benchmark 项目和 Release 解决方案构建均通过。`git diff --check` 仍因既有 `CsvEntityPipeline.cs` EOF 空行失败，未将无关格式问题纳入本 FIX。

### Round 7 汇总

- MUST_FIX：无。
- SHOULD_FIX：`FIX-007` 已完成。
- OPTIONAL：`FIX-006` 继续 `DEFERRED`，需授权人员决定 Git 交付物纳入范围。
- PARTIAL/BLOCKED：formal API baseline、netcoreapp3.1 runtime、完整性能/资源矩阵、Failure Workbook 双 DOM、clean-clone consumer/证据交付仍为 RC No-Go 门禁。
- 回归验证：缓存命中/淘汰 Unit `1/1`；Benchmark build；Release solution build。
- 下一步：重新进行独立 Review；未执行 commit 或 push。

### Round 6

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 执行结论：`FIX-004`、`FIX-005` 的状态/证据问题已处理；`FIX-006` 为 OPTIONAL，按默认范围跳过。Round 6 用户批准的 `P0-02`、`P2-01`、`P2-02`、`P3-01`、`P3-02`、`P3-03` 已实施到当前可验证范围；RC 仍为 `BLOCKED` / `No-Go`。

#### FIX-004

- 严重程度：`HIGH`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`（本轮推荐范围内的执行/证据项）
- 修改文件：
	- `framework.props`
	- `asset/props/target.feature.props`
	- `src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj`
	- `tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj`
	- `build/ApiSnapshot/Program.cs`
	- `tests/Bing.Offices.Tests/PublicApiContractTest.cs`
	- `src/Bing.Offices.Core/...`、`src/Bing.Offices.Npoi/...`、`tests/...`、`docs/excel/nuget-migration.md`
- 根因：Review 指出的 `P2-01/P2-02/P3` 状态与实际已批准实施范围不一致，且部分跨程序集调用仍依赖已 internal 化的 Core 具体实现。
- 修复：移除受控配置中的 net5/net7；删除 Settings 和 Office exception hierarchy；将已批准的 Core plan/type-map/binding/loader/CSV concrete internal 化并改由公开 Provider SPI/接口使用；完成 NPOI Sheet executor、CSV support、cache-key 提取；同步 API 分类、迁移文档和 Benchmark/ResourceProbe 引用。
- 验证：Release solution build `0 error / 28 warning`；StreamPipeline net6/net8 各 `90/90`；Integration net6/net8 各 `15/15`；Docs net8 `11/11`；API 类型清单/分类通过，formal member hash 保持真实失败；未修改 `review.md`。

#### FIX-005

- 严重程度：`HIGH`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `progress.md`
	- `decisions.md`
	- `test-matrix.md`
	- `execution.md`
- 根因：任务运行状态已进入 review-fix Round 6，但历史进度表仍保留 `IN_PROGRESS`/TODO 语义，无法反映本轮真实执行结果。
- 修复：将已验证项更新为 `PARTIAL` 或 `VERIFIED`，补充 Round 6 回归矩阵和决策记录；通过 `task-state.mjs review-fix ... --fix-scope recommended` 确认运行时为 `active=true`、`mode=review-fix`、`reviewRound=6`、`status=implementing`，避免在交给 `task-finish` 前误标完成。
- 验证：`get_errors` 对本轮源码/测试/文档目标为 0；`review.md` 保持原始 `NEEDS_FIX`，未被修改。

#### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`OPTIONAL`
- 执行状态：`DEFERRED`
- 未执行原因：默认 `fixScope=recommended` 不包含 OPTIONAL；任务目录 Git 跟踪与 clean-clone 交付复现仍需授权人员处理，不通过本轮强行执行不可逆或共享状态操作。

### Round 6 汇总

- MUST_FIX：0 个新增未解决项；历史 `FIX-001` 至 `FIX-003` 持续 `RESOLVED`。
- SHOULD_FIX：`FIX-004`、`FIX-005` 的执行证据与状态整理完成。
- OPTIONAL：`FIX-006` `DEFERRED`。
- PARTIAL/BLOCKED：formal API hash 未获批准；netcoreapp3.1 runtime 缺失；性能预算、Failure Workbook 双 DOM 资源边界、长路径 package consumer 和 clean-clone 交付仍阻断 RC。
- 回归验证：Release solution build、StreamPipeline net6/net8、Integration net6/net8、Docs net8、API contract targeted、源码 diagnostics。
- 下一步：执行 `task-finish`，然后交由独立 Reviewer 再次验收；不执行 commit/push。

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 执行结论：本轮 `FIX-001`、`FIX-002`、`FIX-003` 均已实施并完成当前环境可执行的专项验证；发布仍为 `BLOCKED` / `No-Go`，不代表独立 Reviewer 已通过。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
	- `tests/Bing.Offices.Tests/CsvTest.cs`
- 根因：带符号数字豁免只检查符号后的首个数字或小数起始，未验证字段剩余内容。
- 修复：`IsSignedNumber()` 现在从符号开始完整扫描整数、小数和指数语法，并要求扫描结束时正好到达字段末尾；`-1+2`、`+1-2`、`-.5+1`、`-1@cmd` 不再绕过默认转义。
- 验证：`EntityPipeline_SignedNumericPrefixWithTrailingExpression_ShouldEscape` 及 CSV 相关定向回归通过；完整 CSV 输出包含转义后的表达式和未转义的纯数值，未改变 Preserve/None 合同。

#### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/api-diff.md`
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/unit-test-report.md`
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/artifacts/api-snapshot-review-fix-round1/api-snapshot-*.json`
- 根因：报告引用了旧快照值和旧 TRX 计数，未绑定 Round 1 当前 Release 产物。
- 修复：从当前 `output/release` 重新生成四个 shipped TFM 快照，记录 Release DLL SHA-256/长度/UTC 时间，并将报告绑定到 Round 1 snapshot 和 Round 1 TRX。正式 API baseline 保持不变，旧 hash mismatch 仍保留为真实失败证据。
- 当前实际 API hash：Abstractions `407B4F3C2605333A082766B13E1F1DEB704880DDE6D3E0CEDB72FF3F6281ADF0`；Core `5F68499B76921FA52D293BC851B94659FBB2AB466E6D91C8878A97F8728B4BB7`；NPOI `DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE`。
- 验证：Round 1 API compare 退出码 `1`，原因仅为临时 baseline 仍为旧期望 hash；Unit net8 `382 total / 381 passed / 1 failed`，net6 `383 total / 382 passed / 1 failed`，两者唯一失败均为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。

#### FIX-003

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs`
	- `tests/Bing.Offices.ResourceProbe/Program.cs`
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/resource-report.md`
	- `ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/artifacts/excel-resource-probe-rerun.jsonl`
- 根因：资源报告将 `resource-probe-rerun.jsonl` 的 16 个 mapping/unique 场景误写成 Excel ResourceProbe 主证据，且没有持久化七个真实 Excel child 输出。
- 修复：启用 `BING_OFFICES_RESOURCE_PROBE_ARTIFACT` 后保存七个模式的 JSONL 记录，记录输入 hash/bytes、模式、预解析维度、导入行数、状态、耗时、峰值工作集和退出码；报告现在将 Excel 与 mapping/unique 两类 workload 分节说明。
- 验证：`excel-resource-probe-rerun.jsonl` 共 7 条记录，全部 `exitCode=0`；`dom-limit` 为 `resource-limit` 且 `importedRows=100`；最大 `peakWorkingSet=53,276,672`，最大 `rows=250`；`Import_ResourceProbe_ShouldRunInIndependentProcess` 通过。报告明确不宣称任意 DOM 或 Failure Workbook 的硬内存上限。

### Round 1 汇总

- MUST_FIX：`FIX-001`、`FIX-002`、`FIX-003`
- 已完成：3
- PARTIAL：0
- BLOCKED：正式 API baseline、缺失 TFM runtime、性能预算和 Failure Workbook 双 DOM 资源边界仍阻断发布
- FAILED：0（本轮修复专项无新增功能失败；Unit 的既有 API hash 门禁失败保留）
- 回归验证：CSV 定向回归；Unit net8/net6 Round 1 TRX；API 四 TFM snapshot compare；七模式 Excel child probe；报告/artifact 一致性检查
- 下一步：交由独立 Reviewer 再次验收；不得修改 `review.md` 伪造通过

### Round 2

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 执行结论：已完成本轮纳入范围的唯一 `MUST_FIX`：`FIX-001`；未修改 `review.md`。

#### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `tests/Bing.Offices.Tests/CsvTest.cs`
- 根因：上一轮生产修复已完整扫描带符号数值，但新增回归测试只使用 `Contains/DoesNotContain`，未满足完整 CSV 输出断言要求；此前 net8 Round 1/final TRX 也未包含该新增用例。
- 修复：将 `EntityPipeline_SignedNumericPrefixWithTrailingExpression_ShouldEscape()` 改为使用 `Assert.Equal(expectedCsv, content)` 对完整 CSV 文本进行精确比较，覆盖表头、三条记录、`-1+2`、`+1-2`、`-.5+1`、`-1@cmd` 的转义，以及 `-123`、`+1.25e-2` 的合法保留。
- 验证：
	- `dotnet test ... -f net6.0 --filter "FullyQualifiedName~EntityPipeline_SignedNumericPrefixWithTrailingExpression_ShouldEscape"`：`1/1` PASS，TRX 为 `artifacts/review-fix-round2/review-fix-round2-csv-net6.trx`；
	- `dotnet test ... -f net8.0 --filter "FullyQualifiedName~EntityPipeline_SignedNumericPrefixWithTrailingExpression_ShouldEscape"`：`1/1` PASS，TRX 为 `artifacts/review-fix-round2/review-fix-round2-csv-net8.trx`；
	- net6 完整 Unit：`383 total / 382 passed / 1 failed`，CSV 新用例通过，唯一失败仍为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`；TRX 为 `artifacts/review-fix-round2/review-fix-round2-unit-net6.trx`；
	- net8 完整 Unit：`383 total / 382 passed / 1 failed`，CSV 新用例通过，唯一失败仍为 `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`；TRX 为 `artifacts/review-fix-round2/review-fix-round2-unit-net8.trx`；
	- 相关源码/测试 diagnostics：无错误；`git diff --check`：退出码 `0`；
	- net8/net6 定向测试构建成功；构建输出保留既有警告，未降低 API 断言或跳过测试。

### Round 2 汇总

- MUST_FIX：`FIX-001`
- 已完成：`FIX-001`
- PARTIAL：无
- BLOCKED：完整 Unit 的既有正式 API baseline mismatch 仍阻断发布，但不属于本轮 FIX-001 修复失败；本轮未批准或更新 baseline。
- FAILED：无
- 回归验证：CSV 完整输出断言在 net6/net8 各 `1/1` 通过；完整 Unit 在两个可用 TFM 均仅保留既有 API snapshot 失败；diagnostics 和 `git diff --check` 通过。
- 下一步：重新进行独立 Review；不得修改 `review.md` 伪造通过。

### Round 3

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 用户补充要求：`P0-03` 需要支持完全实现。
- 执行结论：已完成 P0-03 的扫描、逐符号引用矩阵和治理裁决；本轮未修改业务代码、测试代码或 `review.md`。P0-03 的 `VERIFIED` 表示证据闭环完成，不表示所有候选已删除；发布仍为 `BLOCKED` / `No-Go`。

#### P0-03 扫描范围与统计

- UTF-8 读取：使用 Python `pathlib.Path.read_text(encoding="utf-8")`；范围为 `src/`、`tests/`、`benchmarks/`、`docs/` 和根 `README.md`。
- 排除目录：`bin/`、`obj/`、`artifacts/`、`output/`，避免生成 XML、旧二进制和历史产物污染统计。
- 文件统计：生产 C# `222` 个、测试 C# `81` 个、Benchmark C# `62` 个、Markdown `14` 个。
- 特殊构造：`Obsolete` `14` 处；生产 `EditorBrowsableState.Never` `71` 处；生产 `EditorBrowsableState.Advanced` `2` 处，均为生成资源成员；生产 `NotImplementedException` `3` 处；生产 `.Result/.Wait()/Task.Run/TODO` 为 `0/0/0/0`；生产 `InternalsVisibleTo` `6` 处，三套生产程序集各仅指向 `Bing.Offices.Tests` 和 `Bing.Offices.Tests.Integration`。
- 其它分类：`catch (Exception)` `17` 处，按异常转换、清理诊断、取消过滤和边界记录逐处归类；Benchmark 的两个 `.Wait()` 和测试的三个 `Task.Run` 不属于生产执行路径。
- API 分类：Abstractions `121` 个（User API `70`、Provider SPI `8`、Execution detail `43`）；Core `61` 个（Compatibility `10`、Execution detail `51`）；NPOI `1` 个 User API；合计 `183` 个公开类型，其中 `Execution detail` `94` 个。分类账与正式 API snapshot 分离，旧 baseline hash 未更新。

#### P0-03 逐符号治理证据

- `deprecated-removal.md` 已补充全部 Obsolete/兼容候选：`ICellValueConverter`、DataTable `CsvHelper`、全局 CSV separator/quote、6 个旧 validation attributes、4 个 Office exceptions、`ExcelSetting`、`SheetSetting`、factory request overload、`AddNpoi` 和 `MaxBytes`。
- 对每个候选记录了定义文件、`src/tests/docs/benchmarks` 引用计数、生产桥接、测试/文档/Benchmark 使用、DI/反射检查、替代路径、删除风险和最终裁决。
- `deprecated-removal.md` 以完整附录列出 Public API 分类账中的 `94` 个 `Execution detail` 类型及定义证据、四类引用计数；统一裁决为等待 breaking approval 和迁移闭环，不以 `EditorBrowsable` 伪装删除。
- `AddNpoi`、`MaxBytes` 为本轮已完成删除/迁移；`NpoiFailureWorkbookWriter.CopyRow` 的三个 `NotImplementedException` 已由 NPOI 2.7.4 HSSF/XSSF 直接能力验证，属于窄范围 capability fallback，保留且不扩展为 catch-all。
- 生产 IVT 仅为 Unit/Integration 测试友元；未发现生产程序集指向 Core/NPOI 实现的友元。

#### P0-03 治理决策

- `decisions.md` D-008 至 D-011 固定扫描口径、`P0-03 VERIFIED` 的定义、公开 API 分类与正式 baseline 的分离，以及后续删除/internal 化的统一前置条件。
- 未批准的 Breaking Change 不机械删除：仍有生产桥接或外部兼容语义的 legacy converter/attributes/CSV/DataTable/exceptions/Settings/UniqueTracker 保持 `BLOCKED/CANDIDATE/PARTIAL`。
- P0-03 不替代 P2-01/P2-02 的实际 API 收敛，也不替代 P3 重构、完整 Benchmark、资源双 DOM 证明和发布门禁。

#### P0-03 验证结果

- 受控 UTF-8 静态扫描：PASS；逐候选定义/引用/裁决：PASS；API 分类覆盖：PASS；生产 IVT 合规检查：PASS。
- 未运行会自动修改源码的 formatter；本轮未执行新的业务构建或测试，因为只补充任务证据，既有 net6/net8 CSV、API、Integration、Docs、ResourceProbe、Benchmark 和 package consumer 结果保持原记录。
- `git diff --check`：待 Round 3 收口验证；`review.md` 的机器协议 `MD041` 不作为本轮问题处理。

### Round 3 汇总

- MUST_FIX：本轮用户明确要求的 P0-03 完整实现
- 已完成：P0-03 扫描统计、逐符号引用矩阵、API 分类附录、DI/反射/控制流分类、替代路径和治理裁决
- PARTIAL：P2-01/P2-02 的实际删除/internal 化、完整 Benchmark/资源/API baseline 仍未完成；这些不被伪记为 P0-03 删除完成
- BLOCKED：正式 API baseline、缺失 TFM runtime、性能预算、Failure Workbook 双 DOM 资源边界和长路径 package consumer 限制继续阻断 RC
- FAILED：无新增失败
- 回归验证：静态扫描和报告交叉检查；保留之前真实 Unit 唯一 API hash 失败及各环境限制
- 下一步：将 `execution.md` 收口为 `COMPLETED` 后运行 task finish，再交由独立 Reviewer 进行“再次验收”；不修改 `review.md`，不执行 commit/push。

### Round 4

- Review 状态：`NEEDS_FIX`
- Fix Scope：`must`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 执行结论：当前 Review 没有未解决的 `MUST_FIX`。`FIX-001`、`FIX-002`、`FIX-003` 已为 `RESOLVED`；`FIX-004`、`FIX-005` 为 `SHOULD_FIX`，`FIX-006` 为 `OPTIONAL`，均不在本轮 `must` 范围内。

#### MUST_FIX 范围核验

- 未修改业务代码、测试代码、`plan.md` 或 `review.md`。
- 未新增或重开任何 `MUST_FIX`。
- 无需运行业务回归测试；此前 `FIX-001` 的 net6/net8 定向测试与 `FIX-002`/`FIX-003` 的证据仍按 Review 记录有效。
- 由于 `SHOULD_FIX` 不在本轮范围内，不能将任务或发布状态解释为通过；现有 `BLOCKED` / `No-Go` 保持不变。

### Round 4 汇总

- MUST_FIX：无未解决项
- 已完成：当前 Review 的 MUST_FIX 范围核验
- PARTIAL：无本轮新增代码修复
- BLOCKED：`FIX-004`、`FIX-005` 未纳入本轮；正式 API baseline、缺失 TFM runtime、性能预算、资源边界和交付复现仍阻断 RC
- FAILED：无
- 回归验证：Review/执行范围核验；未改变生产行为
- 下一步：交由独立 Reviewer 再次验收；本轮不修改 `review.md`，不执行 commit/push。

### Round 5

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-PRE-RC-CLEANUP-20260901-001/review.md`
- 用户授权：当前会话明确允许删除已迁移的 `[Obsolete]` API；本轮不修改 `review.md`，不执行 Git staging/commit/push。
- 执行结论：已完成授权删除子集及 `FIX-004`/`FIX-005` 要求的证据、状态和阻塞闭环；任务仍为 `BLOCKED`，RC 仍为 `No-Go`。

#### Round 5 授权删除子集

- `ICellValueConverter`：删除接口文件；移除 `NpoiExcelImporter`/`NpoiImportRowMaterializer` 的 legacy text bridge、构造参数和测试专用 converter。
- 六个旧 validation attributes：删除 `RequiredAttribute`、`RegexAttribute`、`RangeAttribute`、`MaxLengthAttribute`、`DateTimeAttribute`、`DuplicationAttribute` 文件；生产 binding/rule/type-map/column-plan 分支、测试模型、Docs consumer 和 API contract 全部迁移到 `Excel*` attributes。
- `CsvHelper`：删除 `CsvSeparatorCharacter`、`CsvQuoteCharacter` 及五个依赖全局状态的旧隐式 DataTable 重载；保留显式 delimiter/quote API，并以 `6/6` CSV/校验专项回归确认。

#### FIX-004

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`（按 Review 要求形成逐项延期/阻塞记录）
- 修改文件：
	- `decisions.md`：新增 D-013，逐项记录剩余 P2/P3/P5 项的批准人、后续 taskId、迁移期限、waiver 和 RC 影响；当前没有具名批准、taskId、期限或 waiver 的项目明确写为“无/未分配/未指定”，不伪造批准。
	- `progress.md`：将 P2-01 更新为 `PARTIAL`，P2-02 保持 `PARTIAL`，P3-01 至 P3-04 保持 `TODO`，P5-01 保持 `PARTIAL`，P6-02 更新为 `BLOCKED`。
	- `api-diff.md`、`deprecated-removal.md`、`final-report.md`、`release-checklist.md`：同步 Round 5 删除集、剩余候选和 No-Go 影响。
- 根因：Review 要求计划偏差必须能区分“已批准延期”和“未获批准阻塞”；原报告只有泛化的 `IN_PROGRESS`/`PARTIAL`/`TODO`，缺少逐项治理字段。
- 修复：保留未获批准项目的真实 `BLOCKED/PARTIAL/TODO`，明确无批准人、无后续 taskId、无期限、无 waiver 是当前事实，而不是默认授权；只有 Round 5 用户明确授权的删除子集执行代码变更。
- 验证：独立 Review 所要求的计划项逐项记录已可在 D-013 定位；API、Unit、Integration、Docs、pack、consumer 和负向扫描结果均与当前状态一致。

#### FIX-005

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
	- `progress.md`：删除 P6-02 的“正在创建独立 review”过期表述，改为已完成 Round 5 执行、等待再次独立验收的 `BLOCKED`。
	- `execution.md`：本轮以合法终态 `BLOCKED` 收口，保留任务级 No-Go 与后续 handoff。
- 根因：runtime 已进入 review-fix Round 5，旧 progress 文案仍使用执行中的状态。
- 修复：不重新激活 runtime，不把发布阻断写成通过；将已结束执行与未完成业务门禁分离表达。
- 验证：runtime 在 task-finish 后保持 inactive/finalized；报告不再把当前 Round 5 写成 `IN_PROGRESS`。

#### Round 5 验证

- Release solution build：PASS；`get_errors`：关键 source/test/benchmark 目录 `No errors found`。
- Unit net8：`382 total / 381 passed / 1 failed`；Unit net6：`382 total / 381 passed / 1 failed`；两个唯一失败均为正式 API hash mismatch。
- Integration net8/net6：`15/15 + 15/15 = 30/30`；Docs net8：`11/11`；显式 CSV/校验专项：`6/6`。
- API snapshot：输出 `artifacts/api-snapshot-review-fix-round5`；Abstractions 新 hash `C176D71B0025C1F28F010BF05667588898A4D0EA4F847CD65658D9737D800313`，Core 新 hash `410F6A0F6CF64B41C3AB141AECFB2E1606C9B116EF2BD56C9A868FE02DC8FB68`；compare 退出码 `1` 仅由旧临时 baseline mismatch 导致，未更新正式 hash。
- Pack：`artifacts/packages-round5`；业务包 hash 已记录在 `api-diff.md` 和 `package-consumer-report.md`。
- Package-only consumer：`package-consumer-rerun2` 无 `ProjectReference`，Round 5 包源在 `C:\nupkg-cache-round5` restore/build/run 退出码 `0/0/0`，输出 `package-consumer-ok`；深路径 `MSB3106` 限制仍保留。
- 负向扫描：`src/**/*.cs` 无 `[Obsolete]`/`ObsoleteAttribute`；生产、测试、Benchmark、README 无精确旧符号；Docs 仅迁移表保留旧名称；生成 XML 仅含合法 `Excel*` 或系统属性命中。
- `git diff --check`：退出码 `0`，仅有 CRLF/LF 警告。

### Round 5 汇总

- MUST_FIX：无新增未解决项；`FIX-001` 至 `FIX-003` 继续为 `RESOLVED`。
- SHOULD_FIX：`FIX-004`、`FIX-005` 已完成本轮所需的逐项治理记录和终态修正。
- OPTIONAL：`FIX-006` 未处理，任务目录仍未 staging；按默认 `recommended` scope 跳过。
- PARTIAL/BLOCKED：正式 API baseline、缺失 TFM runtime、完整性能预算、ResourceProbe 边界、剩余 API 治理、长路径 consumer 和 clean-clone 交付仍阻断 RC。
- FAILED：无新增功能失败；Unit 的正式 API hash 门禁失败按预期保留。
- 下一步：执行 `task-finish` 后 handoff“再次验收”，由独立 Reviewer 重新复审；不得修改 `review.md`，不执行 commit/push。
