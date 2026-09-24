<!-- AI_EXECUTION_STATUS: IN_PROGRESS -->
AI_TASK_ID: BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001
AI_EXECUTION_UPDATED_AT: 2026-09-22T11:10:00+08:00

# 实施执行报告

## 执行结论

当前候选已完成约 `99.5%` 的既有整改，但本轮批准的 Performance Budget v1 精准确认重新发现五个 `100K E2E` 目标均超过 `25%`，因此执行目标恢复为 `IN_PROGRESS`，存在 `OPEN_ACTIONABLE=FIX-002C`。既有 List/合同、NPOI Entity/Template 执行链、Provider capability SPI、public-only 第三方 fixture、职责级测试、双 TFM 回归和本地资源证据仍有效；Relation 100K、c16/c64 同质长实验和完整 A-I 按批准边界未重跑。FIX-002 当前为 `FIX-002A=VERIFIED_BOUNDARY`、`FIX-002B=CLOSED`、`FIX-002C=OPEN_ACTIONABLE`；FIX-003 当前为 `FIX-003A/B=VERIFIED_BOUNDARY`、`FIX-003C=CLOSED`、`FIX-003D=LOCAL_PROFILE_EVIDENCE`、`FIX-003E=BLOCKED_EXTERNAL`。本地生产等价 profile 的 100K/500K/1M 代表性 smoke 均通过，但真实生产机和外部 CI 仍未验证；Entity/Template 1M 保持默认 `MaxWorksheetBytes=64 MiB` 的 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`。状态矩阵为 Implementation=`CLOSED`、Test=`CLOSED`、Performance=`OPEN_ACTIONABLE`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`IN_PROGRESS`；不能标记 `COMPLETED`。

## 任务信息

- Task ID：`BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001`
- 计划：`ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/plan.md`
- 分支：`feat/miniexcel-provider`
- 基线提交：`94bb52e84ffc70634067b04541857433ef7af9df`
- 当前候选 manifest：`artifacts/candidates/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/manifest-current-20260921.md`；历史 before/after 绑定仍见同目录下的 `manifest.md` 与 `before-candidate-manifest.md`，不得互相替代。
- 未执行 commit、push、tag、PR 或 NuGet 发布。

## 计划执行情况

| Phase | 状态 | 证据/说明 |
|---|---|---|
| P0 基线 | PARTIAL | 基线提交已通过 `git archive` 提取、独立 restore/build；完整六项目 clean-before、API capture 和资源基线仍未完成 |
| P1 API cleanup | PASS | `ValidateMode` rename、死 overload/未用参数清理；用户已批准新增 Entity/Template 入口，双 TFM snapshot 已更新并通过 compare/Public API 门禁 |
| P2 Provider SPI | PASS | 新增最小 capability/Entity 执行 SPI；公开包第三方 fixture 与成员级 API 合同均已双 TFM 通过 |
| P3 MiniExcel 拆责 | PASS | 已拆出 `MiniExcelSheetPlanBuilder`、`MiniExcelRowMaterializer`、`MiniExcelRelationCoordinator`；主 importer 保留读取/资源/工作簿编排，双 TFM 全量测试通过 |
| P4 Dynamic plan | PASS | Core 建图阶段限制最多一个 dynamic 属性，双 Provider 测试通过 |
| P5 List/RawDate | PASS | compiled materializer、source dictionary 直读、ICollection、按列/行 RawDate |
| P6 Relation | PASS | MiniExcel typed cache/index，MiniExcel 关系专项 `6/6`，NPOI 非 `IList` 关系合同通过，O(P+C) 热路径 |
| P7 Entity | PASS | immutable typed layout、固定 Cell、Merge anchor、List Region、多 Sheet 及 NPOI native executor 已落地；MiniExcel 明确 fail-fast |
| P8 Template | PARTIAL | NPOI Entity template 导入/导出、merge preflight、XLS/HSSF、异步、LeaveOpen、地址/区域边界、命名转换器、样式/公式/列宽/图片锚点、非精确 merge 和 10,000 行边界已覆盖；跨环境资源与发布门禁仍未完成 |
| P9 测试矩阵 | PARTIAL | Core 197/197、NPOI 566/566、MiniExcel 42/42、集成/Docs 已跑；public-only fixture 已通过，全门禁映射仍受 API/资源证据限制 |
| P10 Benchmark | OPEN_ACTIONABLE | Provider/Real IO/RawDate 3/10/30 列、Relation 1K/10K、DynamicPlan、PropertyAccessor、GenericSheetDispatch、Entity/Template after E2E、Materialization/Binding 10K/100K 分离 import/export 和 cold-plan-build tail latency 已有边界/方向性观测；Relation 100K before 为 `FIX-002A=VERIFIED_BOUNDARY`，Entity/Template before 为技术性 `NOT_APPLICABLE_BEFORE`，真实 materializer/binding 当前候选隔离 E2E 为 `FIX-002B=CLOSED`。本轮 Performance Budget v1 仅确认五个已批准的 100K E2E 目标，全部 `5/5` 且 `REGRESSION_OVER_BUDGET`；需先处理 `FIX-002C`，不把本轮当作完整 A-I 判定 |
| P11 资源门禁 | VERIFIED_BOUNDARY | 新增 Windows Job Object 2 CPU 等效/4 GiB 证据；无模板纯列表 SXSSF 的 100K/500K/1M、公开入口和 staging 局部矩阵已有证据；本轮最终候选代表性 `TempFile/excel-100k/c1` 的 100K/500K/1M smoke 均 `passed`、2 次测量、0 错误、`0/0/deleted`。Entity/Template 1M 仍保留默认 64 MiB 资源拒绝边界；完整矩阵、真实生产机和外部 CI 未验证；FIX-003A/B=`VERIFIED_BOUNDARY`、FIX-003C=`CLOSED`、FIX-003D=`LOCAL_PROFILE_EVIDENCE`、FIX-003E=`BLOCKED_EXTERNAL` |
| P12 RC closeout | BLOCKED_APPROVAL | 包、Consumer、API capture、批准后的 snapshot 和 Public API 门禁完成；Performance `FIX-002C=OPEN_ACTIONABLE`，本地 100K/500K/1M smoke 已补齐但生产/外部环境仍未闭合，Release=`BLOCKED_APPROVAL`，Goal=`IN_PROGRESS` |

## 已完成事项

- 将 `ValidateMode` 替换为 `ExcelValidationFailureMode`，同步 builder、request、Provider、测试和文档；无 obsolete wrapper。
- 删除 MiniExcel 死 overload，删除 `ValidateHeaderCount` 未使用参数。
- 修复 `ExcelXlsxZipPreflight` 在 `netstandard2.0` 下不支持的 `String.Contains(..., StringComparison)` 用法。
- Core plan build 阶段拒绝多个 `[DynamicColumn]` 属性。
- MiniExcel 使用编译构造器/setter、source dictionary 直读、`ICollection<T>` 追加；RawDate 支持日期列和有效行范围过滤。
- MiniExcel 关系绑定使用强类型 invoker 缓存和父项索引，保留 first-parent-wins、comparer、错误边界和取消语义。
- MiniExcel 导入职责已拆为工作表计划构建、行物化和关系协调三个 internal 组件；`MiniExcelExcelImporter` 仅保留工作簿/资源/工作表编排与稳定内部转发点，职责组件不扩大公共 API。
- 新增 immutable `ExcelEntityLayout<TEntity>`（固定 Cell、Merge、List Region、多 Sheet）及 `ExcelEntityTemplateOptions`；NPOI 提供实体/模板同步异步执行器，MiniExcel 在 IO 前显式报告 unsupported。
- 新增 `[EditorBrowsable(Never)]` Provider capability/Entity SPI；能力仅用于 preflight 和诊断，不引入 provider router 或运行时切换。
- 新增 Entity 布局快照、模板/HSSF、LeaveOpen、异步取消、地址/区域边界、命名转换器和 MiniExcel fail-fast 直接测试；Core 公开扩展 25/25 已有真实调用追溯。
- 新增模板样式/字体/填充、未映射公式、列宽、图片数据与 ClientAnchor 保真测试；新增三类非精确 Merge 冲突的写前拒绝和零输出测试；新增 10,000 行明细精确边界往返及越界零输出测试；新增实体文件导出预取消/中途取消的目标保护和临时文件清理测试。该条为早期测试快照（Entity 专项 `15/15`、NPOI 全量 `559/559`）；最终当前回归见下方 NPOI `566/566`。
- 新增最终 Release 二进制的 RawDate/Relation hotspot after（3/10/30 列、1K/10K/100K 行，3 次重复）和 Real IO after（24 场景、1K/10K/100K 行，3 次重复）；Provider/hotspot 记录候选程序集 SHA-256，MiniExcel/Real IO 记录同一 Release 运行环境和完整性结果。
- 从基线提交 `94bb52e84ffc70634067b04541857433ef7af9df` 提取 clean source，使用现有全局 NuGet cache 完成独立 restore 和 Release build；Provider 100K、Real IO 24 场景和 RawDate/Relation 部分 hotspot 的 before 证据已保存，详见 `benchmark-report.md` 与 `before-candidate-manifest.md`。
- 新增 cold-plan-build tail-latency before/after 观测：并发 1/4/16/64、100 次操作、5 次重复；并发 16/64 尾延迟抖动明显，预算为 `UNAPPROVED`，只作为补充证据，不构成正式门禁。
- 使用 InProcessEmitToolchain 补齐 DynamicPlan before/after：8 个组合、计划数 100/500、ShortRun 三次迭代；结果已写入 `dynamic-plan-before-after.json`，短迭代和未批准阈值使其仅为方向性证据。
- 使用 InProcessEmitToolchain 补齐 PropertyAccessor before/after：编译/反射 getter/setter 四个场景；结果已写入 `property-accessor-before-after.json`，仅作 materializer 归因 microbenchmark。
- 使用 InProcessEmitToolchain 补齐 GenericSheetDispatch before/after：NPOI 1/8 Sheet 导入/导出四个场景；结果已写入 `generic-dispatch-before-after.json`，仅作多 Sheet 分派归因 microbenchmark。
- 新增逐成员 API 审批请求草案 `api-approval-request.md`；该文件明确列出删除/重命名/新增契约分组和审批后更新 snapshot 的顺序，不构成批准。
- 本轮早先从当前 `output/release` capture 的 `artifacts/api/current-20260921` 发现旧候选 source hash 与物理 DLL 不一致；随后重新 build/pack/capture，形成 `artifacts/api/current-20260921-candidate` 与 `manifest-current-20260921.md`。旧 capture 仅保留为诊断记录，当前刷新候选的 source、Release DLL、包和 API capture 已内部一致；baseline 仍未覆盖，成员审批和完整报告迁移仍未完成。
- 新增仅 PackageReference 的 `Bing.Offices.ThirdPartyProvider.Consumer`，覆盖公开 SPI 分派、模板同步/异步、文件端点、取消、流所有权、能力 preflight、结构化 unsupported 异常；net6/net8 均输出 `third-party-public-only-provider-ok`。
- 新增 MiniExcel 与 NPOI 的非 `IList` `ICollection<T>`/`HashSet<T>` 关系导航合同测试，覆盖真实导出/导入和父子绑定断言；首次无缓存尝试受 `NU1301` 阻断，随后使用隔离缓存双 TFM 运行通过。
- 按 `chinese-comments` 规则治理本次实际修改范围内的 C# XML 参数注释。
- 前一轮已使用 `artifacts/consumer/cache-rc-final3` 隔离缓存完成测试项目编译与运行；流式路径修复后又使用全新 `artifacts/consumer/cache-rc-final7-*` 完成 PackageReference consumer 与 public-only fixture 双 TFM 编译/运行；该条为早期测试快照：MiniExcel 新增关系用例双 TFM 各 `1/1`，NPOI 关系/只读/图片回归筛选各 `4/4`，Entity 专项各 `15/15`，全量 NPOI 各 `559/559`、MiniExcel 各 `42/42`；最终全量见下方。
- Release solution build 0 warning/0 error；Provider unit/integration、Docs、PackageReference consumer、100K probes 有结果。
- 本轮续行使用隔离 NuGet 缓存完成编译，并使用当前 Release 测试 DLL 直接执行复核；该条为早期快照：NPOI net6/net8 各 `559/559`，Entity 专项各 `15/15`，MiniExcel net6/net8 各 `42/42`；最终 NPOI 全量为 `566/566`。
- 修复 NPOI Entity/Template 取消异常包装后的最终 Release build 使用当前 Benchmark DLL（物理 SHA-256 `AB369830E3456DABC08DA1F6DA6C098564B05B70FD751C7BBF339DD1FAD05CFB`）重新执行 MiniExcel 100K、Provider comparison 100K、hotspot 六场景和 Real IO 24 场景；均通过，NPOI/MiniExcel 物理 DLL 与包哈希、候选 manifest 和报告已同步当前候选，所有性能结论仍受 `UNAPPROVED` 阈值限制。
- 新增 NPOI SXSSF 流式纯列表路径和大型列表职责级测试：`Export_LargeSimpleXlsx_ShouldRoundTripRowsAndStyles`、属性级 Merge 排除、Formatter/BodyStyle 顺序和中途取消原子提交保护；该条为早期快照（双 TFM 全量 NPOI `559/559`）。修复后在 Windows Job Object 2 CPU 等效/4 GiB 下重新执行 100K/500K/1M，Job 峰值分别为 `397,295,616`、`1,474,715,648`、`2,356,469,760` bytes，错误为 `0`、输出非零且临时/遗留文件为 `0`。高级布局请求仍保守走 DOM 或显式 unsupported，完整资源矩阵仍未完成。
- 使用当前 Release 二进制在 Windows Job Object 下完成公开入口 `Export`/`ExportAsync`/`ExportToFile`/`ExportToFileAsync` 的 1K/10K/100K 三次重复矩阵 `36/36 PASS`，并完成 Setup/Sampler/Parser 三类失败注入 `3/3`；目标提交、重开和临时文件清理均有结构化记录。
- 使用当前 Release 二进制完成 `Memory`/`TempFile`/`Hybrid` × Excel/失败/模板场景 × 并发 `1/4/16/64` 的 1K 与 10K staging 结构矩阵，各 `36/36 PASS`；实际并发上限为 `1`、最大排队深度 `63`、错误和残留均为 `0`，10K runner 隔离 TEMP 终止后文件数和字节数均为 `0`。审批字段为空，产物保持 `approvalStatus=BLOCKED`；100K 原始长跑产物仍只保留部分记录，当前候选 c64 补充后合并覆盖为 `36/36 PASS`。
- 新增 `EntityProbe` 受控端到端基准入口；最终 Release DLL 完成 1K/10K after 样本和 100K/500K Job Object 受限样本。100K/500K 四种模式均通过真实导出/导入并分别导入 `100000`/`500000` 行，Job 峰值分别为 `622,891,008`/`2,396,901,376` bytes；基线不具备 Entity/Template API，before 明确记为 `NOT_APPLICABLE_BEFORE`。
- EntityProbe 加入后曾记录 Benchmark DLL 物理 SHA-256 `5C9F67A464E051FA9E36ACE68D5676E1B51BAF10239C8D23F7DEDA91E04E5CB6`；随后中间候选的 `471767025D59DD174C9F7F85BCE037570F18625785112093C54C51B06EB52A65` 也仅作历史记录，当前刷新 v2 after 证据统一使用 Benchmark DLL `6CBFFF14075B9544E83A65B95CFE9FB4D9ABCD913DB33437054B81B6CDD82DBD`，旧 benchmark 结果未自动迁移。
- Entity/Template 1M 首次受限尝试运行超过 10 分钟后无 stdout/runner JSON 进展；本轮以 `1800000ms` 上限重跑并形成 `windows-job-entity-template-1m-current-20260921.json`，子进程在 NPOI ZIP 预检阶段因默认 `ExcelResourceLimits.MaxWorksheetBytes=64 MiB` 拒绝 `xl/worksheets/sheet2.xml`，Job 峰值约 `2,049,687,552` bytes，隔离 TEMP 清理为 `0` 文件/`0` 字节。未生成 roundtrip 产物，仍明确记录为 `NOT_VERIFIED`，不填充虚构结果，也不静默放宽默认资源边界。
- 使用 `BING_OFFICES_JOB_TIMEOUT_MS=60000` 与 `BING_OFFICES_JOB_ISOLATE_TEMP=true` 对 1M streaming 场景完成一次强制终止演练；`windows-job-streaming-1m-timeout-isolated-final.json` 记录 `runner-timeout`、Job 峰值 `1,235,165,184` bytes、终止后隔离 TEMP `2` 个文件/`373,260,288` 字节及 `tempCleanupStatus=deleted`。该证据只覆盖清理链路可观测性，不把超时记为性能通过。
- 使用相同 Release 二进制完成 10K `Memory`/`TempFile`/`Hybrid` × Excel/失败/模板 × 并发 `1/4/16/64` staging 矩阵 `36/36 PASS`；实际并发 `1`、最大排队 `63`、错误和残留均为 `0`，隔离 TEMP 终止后文件数和字节数均为 `0`，审批状态按约束保持 `BLOCKED`。
- 使用相同 Release 二进制启动 100K staging 全矩阵；在 1 小时 Job Object 超时前完成 `10/36` 格，全部 `passed` 且错误/残留为 `0`，随后 runner `runner-timeout`，峰值 Job `1,574,567,936` bytes，隔离 TEMP `0` 文件/`0` 字节并成功删除。该原始矩阵证据不等价于完整 100K staging 通过；后续独立 runner 的当前候选补充最终将覆盖提升到 `36/36 PASS`，两个 Excel c64 场景使用 `7200000ms` 串行观察窗口完成，审批仍保持 `BLOCKED`。
- 为隔离长矩阵超时影响，独立完成 100K `TempFile/c1` 的 Excel、失败工作簿和模板图片样式三个场景；三项均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `396,922,880`、`1,443,917,824`、`1,247,277,056` bytes，隔离 TEMP 均为 `0` 文件/`0` 字节并成功清理。该组补充证据不单独改变完整矩阵审批状态。
- 同样独立完成 100K `Hybrid/c1` 的 Excel、失败工作簿和模板图片样式三个场景；三项均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `425,975,808`、`1,400,291,328`、`1,169,821,696` bytes，隔离 TEMP 均为 `0` 文件/`0` 字节并成功清理。该组补充证据不单独改变完整矩阵审批状态。
- 继续独立完成 100K `TempFile/c4` 的 Excel、失败工作簿和模板图片样式三个场景；三项均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `400,752,640`、`1,445,371,904`、`1,291,882,496` bytes，隔离 TEMP 均为 `0` 文件/`0` 字节并成功清理。该组补充证据不单独改变完整矩阵审批状态。
- 继续独立完成 100K `Hybrid/c4` 与 `Hybrid/c16` 的 Excel、失败工作簿和模板图片样式场景；六项均 `passed`、child exit `0`、`errorCount=0`，c4 Job 峰值分别为 `419,799,040`、`1,453,576,192`、`1,210,265,600` bytes，c16 Job 峰值分别为 `432,259,072`、`1,511,641,088`、`1,210,667,008` bytes，隔离 TEMP 均为 `0` 文件/`0` 字节并成功清理。
- 补齐 100K `Memory/template-image-style/c16` 与 `c64` 两个独立场景；均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `1,384,951,808` 与 `1,374,752,768` bytes，隔离 TEMP 均为 `0` 文件/`0` 字节并成功清理。
- 补齐 100K `TempFile/c16` 的 Excel、失败工作簿和模板图片样式三个场景；均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `398,258,176`、`1,492,705,280`、`1,451,864,064` bytes，隔离 TEMP 均为 `0` 文件/`0` 字节并成功清理。
- 继续执行 100K `Hybrid/c64` 与 `TempFile/c64` 的三个场景；六项旧记录均在 `900000ms` runner 门限超时，未形成 child 成功结果，Job 峰值分别为 Hybrid `433,729,536`、`1,508,929,536`、`1,228,423,168` bytes 和 TempFile `400,572,416`、`1,505,271,808`、`1,477,602,368` bytes；隔离 TEMP 均已删除，其中 TempFile Excel/失败/模板终止后分别观测 `3/1/1` 个文件，字节数为 `66,580,121/0/28,756,329`。这些记录保留为旧门限观测，不作为当前候选成功证据。
- 使用当前 Release ResourceProbe（NPOI 依赖 SHA-256 `D357F9A5EFA4D24327BBD8B856AAC7E4875AA0E4DEA7852FE7F50E378FFD75AD`）和 2 CPU 等效/4 GiB Job Object 重新执行六个 c64 场景；四个非 Excel 场景在 `3600000ms` 内完成 `128/128`，两个 Excel 场景先在并行 `3600000ms` 观察窗口超时，终止时分别观测到 `1/2` 个隔离 TEMP 文件并成功清理。随后使用同一当前依赖串行执行 `tempfile-excel-c64-serial-current.json` 与 `hybrid-excel-c64-serial-current.json`，观察窗口延长到 `7200000ms`，两项均完成 `128/128`、`errorCount=0`、隔离 TEMP 清理后为 `0`。因此 100K staging 当前候选覆盖为 `36/36 PASS`；旧 `-rerun` 记录因依赖哈希不一致仅保留为历史观察。
- 本轮只读审计确认 API 审批、A-I 性能和资源门禁报告之间没有状态矛盾；已澄清 `DynamicPlan` 等结果仅为方向性 microbenchmark、`36/36` 仅指公开入口及 1K/10K staging 结构矩阵，并在摘要中明确简单列表 SXSSF 的 1M 证据不包含 Entity/Template 1M。未改变任何代码、版本文件或 API baseline。

## 部分/未完成事项（历史阶段快照）

> 以下条目保留前序阶段的真实记录，不代表当前状态。当前 API 审批和 baseline 已按 Round 4-7 完成收口，FIX-001 当前为 `CLOSED`；历史条目中的 `OPEN / MUST_FIX` 不再代表当前 FIX-002/FIX-003 分类，当前状态以执行结论、Round 41/42 和最终状态/TODO 为准。

- Entity 固定 Cell、Merge、List Region、多 Sheet API 与 NPOI 执行器已实现；固定单元格命名转换器、地址/区域边界、模板流释放、样式/图片/复杂 Merge、10,000 行边界均已有双 TFM 直接证据。
- Template 独立导入/导出入口及 merge conflict preflight 已实现；当前矩阵覆盖 XSSF/HSSF、公式/样式/列宽/图片锚点和非精确 Merge，仍不替代真实受限资源、生产机器和外部 CI 门禁。
- [历史快照] capability/Entity Provider SPI 已实现最小版本；第三方 public-only fixture 已通过，发布 API 审批未完成。
- [历史快照] `build/api-snapshot-baseline.json` 保持不变；新旧 `ValidateMode` diff 等待成员级审批。
- [历史快照] 早先 API capture 与旧候选 manifest 的身份偏差已由当前刷新候选修正；`manifest-current-20260921.md` 已封存当前 source、Release DLL、包、API capture 及 v2 after 产物的绑定关系。旧 capture 仅保留为诊断记录，API compare 仍因待审批的 baseline/member 差异保持 `NOT_VERIFIED`。
- [历史快照] A-I 完整 Benchmark 仍未完成：Provider、Real IO、RawDate 3/10/30、Relation 1K/10K、DynamicPlan、PropertyAccessor、GenericSheetDispatch、Entity/Template after 和 cold-plan-build tail latency 已有观测，但 Relation 100K before、真实 materializer/binding before/after、正式阈值仍缺；Entity/Template before 因基线无该公共 API 记为 `NOT_APPLICABLE_BEFORE`，不能替代可比 before；简单列表 SXSSF 的 500K/1M 和 Entity/Template 的 500K 已在本地 Job Object 通过，Entity/Template 1M 当前重跑被默认 worksheet XML 资源限制拒绝且未完成 roundtrip，公开入口/失败注入和 1K/10K staging 结构矩阵已通过，但已批准的完整策略矩阵、生产机器/外部 CI 也未执行。
- [历史快照] hotspot/Real IO 证据已核对物理 Release DLL SHA 与 API manifest 的 canonical identity hash 语义不同，并在 candidate manifest 中并列记录；未把两种哈希混用为同一字段。
- [历史快照] 独立审查已更新 `review.md`，结论仍为 `NEEDS_FIX`；FIX-001 至 FIX-003 仍有真实未完成门禁，当前候选 100K staging c64 六项最终通过但两个 Excel 场景需要较长串行观察窗口，原 FIX-004 的测试缺口已由直接证据关闭。
- 当前状态（Round 42 更新）：FIX-001 已按用户批准、baseline、compare 和 Public API 门禁完成并标记为 `CLOSED`；FIX-002A=`VERIFIED_BOUNDARY`、FIX-002B=`CLOSED`、FIX-002C=`BLOCKED_APPROVAL`；FIX-003A/B=`VERIFIED_BOUNDARY`、FIX-003C=`CLOSED`、FIX-003D/E=`BLOCKED_EXTERNAL`。状态汇总为 Implementation=`CLOSED`、Test=`CLOSED`、Performance=`BLOCKED_APPROVAL`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`STOPPED`；没有 OPEN_ACTIONABLE，不能标记 `COMPLETED`。
- 首次无缓存的 `ICollection<T>` 关系合同尝试曾返回 `NU1301`；该阻断已由 `artifacts/consumer/cache-rc-final3` 隔离缓存编译和双 TFM 运行证据覆盖，不再是当前测试项阻断。
- 根 `artifacts/` 只读审计确认当前候选 `candidates/benchmarks/consumer/api/packages/exception-probe` 必须保留；已删除明确识别的历史 `review-round2-probe/artifacts`、旧 Consumer 输出和所有 `artifacts/**/bin|obj` 目录，删除前均确认未被 Git 跟踪且目标位于根 `artifacts/` 内。删除后扫描结果为 `nested_generated_dirs=0`。
- 关系合同复核已在隔离缓存下完成；此前的 `NU1301` 仅记录首次无缓存尝试，不再是该测试项的当前阻断。

## 修改文件

生产/API：

- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelWorkbookImportRequest.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ValidateMode.cs`（删除）
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelValidationFailureMode.cs`（新增）
- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelTypeMapFactory.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelEntityExtensions.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Entities/{ExcelEntityLayout,ExcelEntityResults,IExcelEntityImporter,IExcelEntityExporter}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Providers/ExcelProviderCapabilities.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelSheetPlanBuilder.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelRowMaterializer.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelRelationCoordinator.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Exports/MiniExcelExcelExporter.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelRawDateSerialReader.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Imports/{ExcelImportExecutionOptions,NpoiExcelImporter,NpoiImportRowMaterializer,NpoiImportSheetExecutor,NpoiWorkbookValidationPipeline}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelExporter.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Entities/NpoiEntityLayoutExecutors.cs`
- `src/provider-shared/ExcelXlsxZipPreflight.cs`
- `tests/Bing.Offices.WindowsJobRunner/{Bing.Offices.WindowsJobRunner.csproj,Program.cs}`（新增本地 Windows Job Object 资源限制工具）
- `benchmarks/Bing.Offices.Benchmarks/Program.cs`
- `benchmarks/Bing.Offices.Benchmarks/EntityProbe.cs`（新增 Entity/Template 公开 E2E 受控探针）

任务证据/迁移：

- `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md`

测试/文档：本任务 diff 中的 MiniExcel、NPOI、Integration 测试、`tests/Bing.Offices.ThirdPartyProvider.Consumer`、`review.md` 和 `docs/excel/import-validation.md`；完整映射见 `symbol-test-map.md`。

## API/数据/配置变化

- Public API Breaking：`ValidateMode` → `ExcelValidationFailureMode`；属性 `ValidateMode` → `ValidationFailureMode`；新增 Entity/Template layout、扩展入口及 capability/SPI 类型。
- 无版本号、版本属性、provider registration、runtime router 或生产程序集 IVT 变化。
- API baseline 已在 FIX-001 用户批准后更新；当前双 TFM compare 输出为空差异，历史差异见 `api-diff.md` 和 `artifacts/api/fix001-20260921-compare-before-baseline/`。

## 测试结果

- Core unit：net6/net8 各 `197/197`。
- NPOI unit：net6/net8 各 `566/566`；Entity 专项双 TFM 已通过（最终 Entity 专项 `22/22`）。
- MiniExcel unit：net6/net8 各 `42/42`；关系专项 `6/6`，Entity fail-fast `2/2`。
- NPOI integration：net6/net8 各 `30/30`。
- MiniExcel integration：net6/net8 各 `9/9`。
- Docs：net8 `10/10`。
- Public API integration：当前 net6/net8 各 `9/9 PASS`（`PublicApiContractTest` 专项）；此前 Round 3 的 `27/29` 失败已由批准后的 baseline、测试分类和成员清单同步修复。
- PackageReference consumer：使用全新 `--packages artifacts/consumer/cache-rc-final7-net6`/`final7-net8` 隔离缓存重新验证，net6/net8 restore/build/run PASS；两次输出 `package-consumer-ok`，net6 有既有 `NETSDK1138` warning；重打包内 DLL 与最新 Release 输出物理 SHA-256 一致。
- Third-party public-only fixture：使用全新 `cache-rc-final7-public-net6`/`final7-public-net8` 隔离缓存，net6/net8 build/run PASS；两次输出 `third-party-public-only-provider-ok`，net6 有既有 `NETSDK1138` warning；旧缓存和全局同号缓存结果已排除。
- 最新 Release DLL 双 TFM 复核：NPOI 全量 `566/566`、MiniExcel 全量 `42/42`；新增非 `IList` 关系合同、大型 SXSSF round-trip/样式、属性级 Merge 排除和取消提交保护测试已计入上述结果。
- MiniExcel importer 拆责后重新编译并执行双 TFM 全量测试：net6/net8 各 `42/42 PASS`；关系反射职责测试的稳定转发点保留，未发生行为漂移。
- NPOI 集成 DLL 在取消异常包装修复后最终复跑：net6/net8 各 `30/30 PASS`，本轮耗时分别约 2 分 44 秒和 2 分 35 秒。
- 本轮继续使用当前 Release DLL 新鲜复跑职责与集成合同：NPOI 单元 net6/net8 各 `566/566 PASS`，MiniExcel 单元各 `42/42 PASS`；NPOI 集成各 `30/30 PASS`，MiniExcel 集成各 `9/9 PASS`。未改变测试二进制、版本文件或 API baseline。
- 本轮修复后重新运行 NPOI 全量单元 net6/net8，各 `566/566 PASS`；NPOI 项目、Windows Job Object runner 和 Release solution build 均 `0 warning/0 error`。大型列表合并列排除、Formatter/BodyStyle 顺序和取消提交保护的直接测试均计入全量结果。
- （历史 Round 3）公共 API 集成门禁曾为 net6/net8 各 `27/29 PASS`，两个失败均为待审批 baseline/snapshot 差异；该历史失败由本轮批准后的 baseline 和测试契约同步闭合。
- 当前双 TFM API capture 目录为 `artifacts/api/fix001-20260921-capture/`，compare 目录为 `artifacts/api/fix001-20260921-compare-after-baseline/`；两者均绑定当前 immutable candidate，compare 无 unexpected/removed 差异。
- 当前刷新候选的生产回归产物位于 `artifacts/tests/candidate-20260921/`：Core `197/197`、NPOI `566/566`、MiniExcel `42/42`、NPOI integration `30/30`、MiniExcel integration `9/9` 均在 net6/net8 通过；Docs net8 `10/10` 通过；Public API 专项 net6/net8 各 `9/9 PASS`。
- Benchmark Release build：`benchmarks/Bing.Offices.Benchmarks` 0 warning/0 error；EntityProbe 1K/10K after 通过，100K Job Object 2 CPU 等效/4 GiB 通过四种模式，均导入 `100000` 行。
- 当前刷新候选使用 Benchmark DLL SHA `6CBFFF14075B9544E83A65B95CFE9FB4D9ABCD913DB33437054B81B6CDD82DBD` 重新生成 after 观察：Provider comparison NPOI/MiniExcel 各 100K、3 次重复通过；Real IO 24 场景、RawDate/Relation hotspot 六场景和 Entity/Template 10K 四模式均通过。Real IO header 已包含五个实际程序集 SHA，预算仍为 `UNAPPROVED`。
- 根 `artifacts/` 治理：历史 nested `artifacts`、`bin`、`obj` 目录已删除；深度扫描结果 `nested_generated_dirs=0`。旧候选 benchmark DLL hash 为 `5C9F67A464E051FA9E36ACE68D5676E1B51BAF10239C8D23F7DEDA91E04E5CB6`，当前刷新候选 after 证据使用 `6CBFFF14075B9544E83A65B95CFE9FB4D9ABCD913DB33437054B81B6CDD82DBD`，未把旧 benchmark 结果冒充为新候选。
- 当前工作树候选刷新：重新执行 Release solution build、四个生产包 pack 和双 TFM API capture；新 API source-scope hash 为 `E4447DEB168499218EE856357197D80A96208F625F79F4DCABDCB1E4AB101346`，并由 `manifest-current-20260921.md` 封存当前 Release DLL、包和 API capture 的身份。新包 consumer 与 public-only fixture 均使用全新缓存双 TFM 通过；旧 benchmark/resource 证据未静默迁移到刷新候选。
- （历史 Round 3）当前刷新候选对旧 API baseline 的 compare 曾按预期失败，结果保存在 `artifacts/api/compare/current-20260921-candidate`；该目录仅作诊断证据，不代表批准后的门禁结果。
- （历史 Round 3）续行复核曾将 Public API 输出写入 `artifacts/tests/continuation-20260921/`，结果为 `27/29 PASS`、`2 FAIL`；本轮已用当前测试源码重建并在 `artifacts/tests/fix001-20260921/net6.0/`、`net8.0/` 重新通过。
- FIX-001 本轮重新 build/pack/capture 后更新 `build/api-snapshot-baseline.json`，`approvedBy=user:FIX-001`、`approvedAt=2026-09-21T17:12:16+08:00`；版本文件哈希保持不变。
- 本轮 1M 资源重跑输出 `windows-job-entity-template-1m-current-20260921.json`，状态为 `child-failed`（默认 worksheet XML 64 MiB 预检拒绝），Job 峰值 `2,049,687,552` bytes，隔离 TEMP `0/0` 且 `deleted`；独立 review 已同步该证据并保持 `NEEDS_FIX`。

## Build/Typecheck/Lint/Format

- `dotnet build Bing.Offices.sln -c Release --no-restore --nologo`：PASS，0 warning/0 error。
- 续行重新执行 `dotnet build Bing.Offices.sln -c Release --no-restore --nologo`：PASS，0 warning/0 error；当前 Benchmark/NPOI/MiniExcel/Abstractions/Core 物理 SHA 与 `manifest-current-20260921.md` 保持一致。
- Round 3 后使用 `dotnet build Bing.Offices.sln -c Release --no-restore --nologo --disable-build-servers -m:1 -p:UseSharedCompilation=false` 重试：PASS，0 warning/0 error；前一次共享编译并发写入失败仅保留为环境观察。
- `dotnet build tests/Bing.Offices.ResourceProbe/Bing.Offices.ResourceProbe.csproj -c Release --no-restore`：PASS，0 warning/0 error；Windows Job Object runner 项目和 NPOI 项目 Release build 亦为 0 warning/0 error，solution build 0 warning/0 error。当前候选 SXSSF 纯列表受限 100K/500K/1M 均通过，详见 `resource-report.md`；旧 DOM OOM 证据作为历史风险保留。
- `git diff --check`：PASS。
- `git diff --stat`：已核对；版本文件 SHA-256 未变化，API baseline 仅按 FIX-001 批准更新。

## 计划偏差

本次按计划落地了最小 Entity/Template/SPI 契约。FIX-001 已完成逐项批准、当前 candidate capture、baseline 更新和双 TFM Public API 门禁；before 性能已对 Provider/Real IO/部分 hotspot 形成独立基线，但预算和受限资源仍未批准，缺失矩阵按 `NOT_VERIFIED` 处理。

## 基线问题

初始受限 restore 返回 `NU1301`；随后对基线归档目录使用外部 restore 权限和现有全局 cache 成功完成 restore/build。基线 probe 与当前候选使用不同提交但各自物理 SHA 已记录，未混用历史 artifacts；它们只用于前后观察，不代表正式阈值结论。

## 已知问题

- （历史 Round 3）API snapshot 测试因旧 `ValidateMode` 基线和未审批成员差异失败；FIX-001 已按批准流程更新 baseline 并验证通过。
- `NETSDK1138` 提示 net6 已结束支持，属于环境 warning；不改变项目目标框架。
- 本轮普通 BenchmarkDotNet `ShortRun` 的自动生成项目 restore 因 NuGet SSL/凭据错误 `NU1301` 失败；随后使用 InProcess toolchain 完成同一组 DynamicPlan before/after，普通模式的 `NA` 报告未被当作性能通过。
- GenericSheetDispatch 的 BenchmarkDotNet Import 诊断在 1/8 Sheet 分别报告 3/10 个 exception events，before/after 均存在；最小公开基准探针确认类型为 `System.MissingMethodException`，堆栈落在 NPOI `XSSFFactory.CreateDocumentPart`，导入仍分别得到 1/8 行。证据见 `generic-dispatch-exception-probe.json`；该 microbenchmark 不作为无异常发布门禁，也不归因于本任务改动。
- MiniExcel date-column inference 对 dynamic mapping 保留保守全 numeric 扫描路径，需后续基准确认进一步收窄是否安全。

## 风险与回归关注点

- `ExcelValidationFailureMode` 是 Breaking API，必须先取得成员级审批再更新双 TFM snapshot。
- MiniExcel 对 Entity/Template/Merge 保持显式 unsupported，不静默退回 List/workbook 路径；NPOI 只执行已声明的布局能力。
- 100K Provider、Real IO、RawDate 部分和 Relation 部分已有前后观察，但没有批准阈值；当前候选 SXSSF 纯列表已证明 2 CPU 等效/4 GiB 限制下 100K/500K/1M 可完成，Entity/Template 已完成 100K/500K 单次受限运行但 1M 超时无产物；这些证据不能推断并发/取消/失败提交的生产容量或替代生产/外部 CI 证据，旧 DOM OOM 仅作为历史风险记录。

## Reviewer 注意事项

- 重点复核 MiniExcel materializer、RawDate 过滤边界、relation typed cache 的错误边界和取消语义。
- 复核 `api-diff.md` 的 Breaking 成员是否批准；批准后重新 capture/compare 并更新 baseline。
- 复核所有报告是否引用同一个候选 manifest，避免混用历史 artifacts；当前 Release/API/package 候选以 `manifest-current-20260921.md` 为准，历史 before 仅以其自身 manifest 标识。

## Review 修复记录

> Round 1-40 以下记录均为历史快照；其中出现的旧状态（例如 FIX-002/FIX-003 为 `OPEN / MUST_FIX`）不代表当前状态。当前状态以执行结论、Round 41/42 和最终状态/TODO 为准。

### Round 1

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/review.md`
- 执行时间：`2026-09-21T06:31:29.3562771Z`

#### FIX-001

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：DEFERRED
- 修改文件：
  - `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/execution.md`
  - `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md`
- 根因：成员级审批不是当前工作区可以代替完成的决定；旧 baseline 仍未获批，不能通过覆盖 snapshot 消除差异。
- 修复：将执行报告和审批请求统一指向当前候选 API capture/compare 产物；保留旧 baseline、Breaking diff 和双 TFM `27/29` 失败证据，不伪造审批或更新 baseline。
- 继续处理条件：审批人对 `ValidateMode` rename、Entity/Template 和 Provider capability/SPI 逐成员给出 `APPROVED`/`REJECTED` 决定。
- 验证：
  - `artifacts/api/compare/current-20260921-candidate/api-diff.json`：存在，候选 source-scope hash 与当前 manifest 一致。
  - `build/api-snapshot-baseline.json`：未修改；双 TFM Public API 各 `27/29`，失败仅为待审批 baseline/member mismatch。

#### FIX-002

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：DEFERRED
- 修改文件：
  - `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/execution.md`
- 根因：正式性能阈值未批准，Relation 100K before 和真实 materializer/binding 同候选 before/after 尚未形成；当前 Entity/Template before 无同等旧公共 API，只能保持 `NOT_APPLICABLE_BEFORE`。
- 修复：明确当前候选 after、缺失 before 和阈值状态均为不可发布判定的证据，不把方向性 microbenchmark 或 after-only 样本升级为正式性能结论。
- 继续处理条件：批准 A-I workload、数据集、重复次数、指标和阈值，并在同一可追溯候选身份下补齐缺失 before/after。
- 验证：
  - `benchmark-report.md`：Relation 100K、真实 materializer/binding 和正式阈值继续标记 `NOT_VERIFIED`/`UNAPPROVED`。
  - `manifest-current-20260921.md`：当前 Benchmark DLL SHA-256 `6CBFFF14075B9544E83A65B95CFE9FB4D9ABCD913DB33437054B81B6CDD82DBD` 已封存。

#### FIX-003

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：DEFERRED
- 修改文件：
  - `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/execution.md`
- 根因：批准的完整 `500K/1M` 资源矩阵、生产机器和外部 CI 未提供；Entity/Template `1M` 当前在默认 `MaxWorksheetBytes=64 MiB` 预检处结构化拒绝，没有 roundtrip 产物。
- 修复：保留该拒绝作为资源限制证据，明确为 `NOT_VERIFIED`；不提高默认资源限制、不将超时或拒绝改写为通过，并保留隔离 TEMP `0/0`、`deleted` 的清理证据。
- 继续处理条件：在批准的环境和门限下完成取消、失败提交、临时文件、原子提交以及 `500K/1M` 矩阵，并补充生产机器/外部 CI 结果。
- 验证：
  - `windows-job-entity-template-1m-current-20260921.json`：`child-failed`，因 worksheet XML 超过默认 64 MiB 被拒绝，未生成 roundtrip。
  - `resource-report.md`：当前 1M、完整取消/失败矩阵、生产/外部环境继续标记 `NOT_VERIFIED`。

### Round 1 汇总

- MUST_FIX：FIX-001、FIX-002、FIX-003 均已检查并形成明确处置结论。
- 已完成：当前候选证据路径收敛；API baseline、性能阈值和资源边界保持未篡改。
- PARTIAL：本地回归和当前候选证据已存在，但外部门禁未闭合。
- BLOCKED：API 成员审批、正式性能阈值、生产/外部资源环境需要外部输入。
- FAILED：无新的实现或验证失败。
- 回归验证：Release solution build `0 warning / 0 error`；Core `197/197`、NPOI `566/566`、MiniExcel `42/42`、NPOI integration `30/30`、MiniExcel integration `9/9`（双 TFM）；`git diff --check` 通过。
- 下一步：保留本记录，交由下一轮独立 Review 重新判定；外部门禁满足前不得标记 `COMPLETED`。

### Round 2（Luna Max）

- Review 起点：Sol Medium 新增 `FIX-005`（固定 Cell 未复用 mapping/value-map/validation）和 `FIX-006`（List Region 列边界未统一校验）。
- `FIX-005`：已修复。固定 Cell 增加与 List Region 一致的 `Mapping(ExcelMappingConfiguration/ExcelMappingDocument)` 快照入口；NPOI 导入/导出通过既有 `IExcelMappingPlanFactory` 创建 `ExcelColumnPlan`，复用 converter、value-map、validation 和错误坐标语义。
- `FIX-006`：已修复。NPOI Entity executor 在导入和导出写入前统一校验 `End.Column`/行边界；空集合含表头、无表头空集合、导入列宽越界和精确边界均有直接测试。
- 回归验证：`EntityLayoutProviderTest` net6/net8 各 `20/20 PASS`；NPOI 全量 net6/net8 各 `564/564 PASS`；`git diff --check` 通过（仅保留既有 ProfileFixtures.xml 行尾提示）。
- 本轮未修改 `review.md`、API baseline、版本文件或资源限制；等待 Sol Medium 独立复审确认 `FIX-005/006`。

### Round 3（Luna Max）

- Review 起点：Sol Medium 复审确认 `FIX-005` 仍存在固定 Cell 导入先转换、后执行 raw validation 的顺序缺口。
- `FIX-005`：已修复。固定 Cell 导入现在先执行映射计划中的 raw validation；仅在通过后调用 `ConvertFrom`，再执行 converted/request-level validation；任一阶段失败均不调用 setter，并沿用既有 `IExcelValidationBinding`、错误码及 sheet/row/column/property 错误上下文。
- 职责级测试：新增不可空数值空文本 raw validation 回归，以及请求级命名校验在转换后执行且失败阻止 setter 的直接测试；既有 value-map、命名 converter、同步/异步入口测试保持不变。
- 回归验证：`EntityLayoutProviderTest` net6/net8 各 `22/22 PASS`；随后独立复跑 NPOI 全量 net6/net8 各 `566/566 PASS`。
- 本轮未修改 `review.md`、API baseline、版本文件或资源限制；未执行 commit、push、tag、PR 或 NuGet 发布。

### Round 4（FIX-001 批准后，Luna Max）

- 批准输入：用户于 2026-09-21 明确批准 `FIX-001`，覆盖 `ValidateMode` rename、Entity layout/result/template、Entity importer/exporter SPI、Provider capability SPI 及 NPOI/MiniExcel 公开实现接口形状；逐项记录见 `api-approval-request.md`。
- 当前候选重新执行 Release solution build、四个生产包 pack 和双 TFM API capture；构建为 `0 warning / 0 error`，候选 source-scope hash 为 `A1D558E350FF8819A60D7502684CDB4F84BE0B965F9CD56AA3B304C316355E89`。
- API diff 已按当前候选确认：Abstractions `111 added / 6 removed`、Core `11 / 0`、NPOI `18 / 2`、MiniExcel `8 / 2`；FIX-005 后新增的 7 个 Abstractions Cell mapping 成员也包含在批准范围内。
- `build/api-snapshot-baseline.json` 已按批准的当前候选更新，记录 `approvedBy=user:FIX-001`、`approvedAt=2026-09-21T17:12:16+08:00`；双 TFM API compare 输出无差异（`artifacts/api/fix001-20260921-compare-after-baseline/api-diff.json` 的 net6/net8 均为空对象）。
- Public API 合同测试最初暴露测试分类和期望成员清单缺口；Luna Max 仅修改 `PublicApiContractTest.cs`，补齐已批准的 Entity/Provider/Core 类型分类与成员清单，并将 `ExcelValidationFailureMode` 纳入合同。重建后 net6.0、net8.0 各 `9/9 PASS`。
- 本轮核对 `git diff --check`；版本文件 `version.props`、`version.dev.props`、`common.props`、`framework.props` 哈希保持不变；未执行 commit、push、tag、PR 或 NuGet 发布。
- FIX-001 已关闭；FIX-002（A-I 正式性能阈值）和 FIX-003（完整资源/生产/外部 CI 门禁）仍保持 `NOT_VERIFIED`，因此任务状态继续为 `PARTIAL`。

### Round 5（FIX-001 identity 修复，Luna Max）

- Sol Medium 复审发现上一轮 candidate identity 的 `breakingApprovalArtifact` 仍硬编码旧任务 `BO-RC-20260908-002`，因此上一轮 FIX-001 不能关闭；本轮按该发现继续修复，不手工伪造 baseline。
- Luna Max 仅修改 `build/ApiSnapshot/Program.cs` 和 `CandidateIdentityContractTest.cs`：新增显式 `--approval <relative-path>`，Capture/Validate 统一规范化并绑定配置路径；拒绝绝对路径、驱动器路径、越界、缺失、路径不匹配和篡改。保留旧默认路径以兼容既有 identity self-test。
- `ApiSnapshot` Release build 为 `0 warning / 0 error`；`--identity-self-test true` 通过，覆盖 LF/CRLF checkout、程序集/nupkg/批准文件篡改、默认路径兼容、自定义路径、路径不匹配和越界。
- 重新 capture 使用：`artifacts/api/fix001-20260921-capture-v2/`，显式批准文件为 `ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md`；当前 source-scope hash 为 `9E552C7553D33E768637DA76FFA51706936B14B074B82F2BA888186AD5F9A1C2`，批准文件 hash 为 `B6BD00B1C6FFF4730AD37DAE52FCADAE42DEC18C9011C6BA1F583308524667F0`。
- 基于同一当前 Release DLL/nupkg 和用户批准重新写入 `build/api-snapshot-baseline.json`；baseline 的 `approvedBy=user:FIX-001`、`approvedAt=2026-09-21T17:12:16+08:00`、candidate identity 路径/哈希均指向本任务批准文件。
- 当前路径 compare `artifacts/api/fix001-20260921-compare-after-baseline-v2/` 双 TFM 通过且 diff 为空；故意传入旧任务批准路径的负向 compare `artifacts/api/fix001-20260921-compare-wrong-approval/` 以路径不匹配退出，证明身份链不会接受旧批准。
- 重建后的 Public API 合同测试 net6.0、net8.0 各 `9/9 PASS`，TRX 位于 `artifacts/tests/fix001-20260921/net6.0-v2/` 和 `net8.0-v2/`；四个版本文件哈希保持不变。未执行 commit、push、tag、PR 或 NuGet 发布。
- 本轮完成 FIX-001 的身份链修复；FIX-002（A-I 正式性能阈值）和 FIX-003（完整资源/生产/外部 CI 门禁）仍为 `MUST_FIX`/`NOT_VERIFIED`，任务状态继续为 `PARTIAL`。

### Round 6（FIX-001 CI 接入，Luna Max）

- Sol Medium 复审发现 CI 的 API compare 未传入当前批准文件；Luna Max 仅修改 `.github/workflows/ci.yml` 的正式 compare 命令，加入 `--approval ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md`，identity self-test 保持不变。
- 核对 workflow 顺序：`release-gates` 先在干净 checkout 中重新 pack 四个生产 nupkg 到 `artifacts/packages` 并上传，`api-gate` 下载同一 release evidence 后再执行显式批准路径的 compare；因此不依赖工作区历史 package 目录。
- CI 文件验证：`git diff --check` 通过；UTF-8 解码成功、无 BOM、无 CRLF、末尾 LF；修改为 1 行替换。
- 当前 immutable candidate 使用 `artifacts/packages/fix001-20260921` 的 API compare 继续通过；本地根 `artifacts/packages` 中的历史包与当前 Release DLL 不一致，针对该目录的诊断命令按预期失败，未将其作为当前候选证据。CI 的新鲜 pack 链路已在 workflow 中固定。
- Sol Medium 最终独立复审已确认 FIX-001 `CLOSED`：工具 identity、批准文件绑定和 CI 调用链均已修复；identity self-test、正确/错误批准路径 compare、双 TFM Public API `9/9` 和 fresh pack 链路证据一致。FIX-002（A-I 正式性能阈值）和 FIX-003（完整资源/生产/外部 CI 门禁）仍为 `MUST_FIX`/`NOT_VERIFIED`。

### Round 7（最终编码与 candidate identity 收口）

- 最终字节检查确认任务内 `execution.md`、`review.md`、`api-approval-request.md` 均为 UTF-8 无 BOM、LF、真实末尾换行；清理过程中产生的字面量尾缀已移除，未改变审批正文语义。
- 当前批准文件字节与 v2 capture 一致，SHA-256 为 `B6BD00B1C6FFF4730AD37DAE52FCADAE42DEC18C9011C6BA1F583308524667F0`；最终 baseline 以 `artifacts/api/fix001-20260921-capture-v2/` 的 candidate identity 为准，source-scope hash 为 `9E552C7553D33E768637DA76FFA51706936B14B074B82F2BA888186AD5F9A1C2`。
- 重新写入 `build/api-snapshot-baseline.json` 后，`artifacts/api/fix001-20260921-compare-after-baseline-final/` 双 TFM compare 通过且 diff 为空；最终 Public API 合同测试已写入 `artifacts/tests/fix001-20260921/net6.0-final2/` 和 `net8.0-final2/`，各 `9/9 PASS`。
- v3/0858 仅保留为误加字面量尾缀期间的诊断产物，不作为当前 baseline 或发布证据。版本文件哈希、CI `--approval` 接入和未执行 commit/push/tag/PR/NuGet 约束保持不变。
- Sol Medium 已按最终 v2 证据再次独立复审：批准文件真实末尾字节为 `0A`，baseline 与 capture-v2 全字段一致，FIX-001 保持 `CLOSED`；FIX-002/FIX-003 继续 `OPEN + MUST_FIX`。

### Round 8（Relation before 受限诊断，Luna Max）

- Review 起点：FIX-002 仍缺少 Relation 100K before；旧实现为 O(P×C) 逐子项线性搜索，完整运行不能在可控窗口内结束，不能把超时或中断伪造成性能数字。
- Luna Max 仅修改 `benchmarks/Bing.Offices.Benchmarks/HotspotProbe.cs`：增加每个 worker 的可配置超时（默认 60 秒）、整棵子进程终止、已刷盘样本读取、部分结果结构化记录和非零 `not-verified` 退出码；不改变生产 API、Provider 实现或正式候选程序集。
- 诊断运行：`BING_OFFICES_HOTSPOT_TIMEOUT_MS=60000 dotnet benchmarks/Bing.Offices.Benchmarks/bin/Release/net8.0/Bing.Offices.Benchmarks.dll --hotspot-probe artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-timeout-60s-20260921.json before 3`。
- 结果：Benchmark 构建 `0 warning / 0 error`；RawDate 3/10/30 与 Relation 1K/10K 各 3 次完整样本；Relation 100K 在 `60089.3345 ms` 后终止，`completedRepetitions=0`、`status=not-verified`、worker exit `-1`，总退出码 `2`。短超时负向演练 `hotspot-before-timeout-20260921.json` 亦按预期记录 6 个 `not-verified` 场景。
- 候选绑定：该诊断修改改变了 Benchmark DLL SHA（本轮构建物理 SHA 为 `D0B53B798B4DFEE1B35C237B826F2470C78B30F5960F2596A64A64C5B4909679`），所以新 timeout artifact 不替换 `manifest-current-20260921.md` 中冻结的正式 after 证据；报告已明确其为独立诊断证据。
- FIX-002 状态：仍为 `OPEN / MUST_FIX`。安全终止和可追溯缺口已补强，但 Relation 100K before、真实 materializer/binding before/after 与正式阈值仍未完成。

### Round 9（Hotspot 超时语义修正，Luna Max）

- Sol Medium 复审发现两个 `SHOULD_FIX`：超时边界可能丢弃 worker 错误输出，且部分样本仍可能被汇总为数值中位数。
- 修复范围仅为 `benchmarks/Bing.Offices.Benchmarks/HotspotProbe.cs`：
  - `WorkerProcessResult` 记录 `killAttempted`、`killSucceeded`、`killError`、退出码及截断后的 stdout/stderr；超时记录用 `timeout-killed` 与 `timeout-with-nonzero-exit` 区分主动终止和未能确认的 worker 非零失败。
  - worker 尚未创建 JSONL 文件时按空样本处理，仍输出结构化 timeout；只要未完成要求的重复次数，summary 的 wall time、allocated、peak、index 和 LOH 五项统计全部为 `null`。
- 验证：Benchmark Release build `0 warning / 0 error`；100 ms 负向探针退出码 `2`，六个场景均为 `not-verified` 且 summary 统计为空；60 秒运行保留 RawDate 3/10/30、Relation 1K/10K 的完整样本，Relation 100K 以 `timeout-killed` 记录并退出码 `2`；9 秒边界运行中所有不完整场景 summary 统计仍为空。
- 当前诊断 DLL SHA 为 `1156E963DCF62330A2912F4D9F93EDC549E9AC6FC17A541A29CDD161CEF0B02F`，与冻结候选不同；timeout artifact 仍只作为独立诊断证据，不替代正式 A-I before/after。
- SHOULD_FIX 状态：已按要求处理，等待下一轮 Sol Medium 独立复审；FIX-002 主项与 FIX-003 仍保持 `OPEN / MUST_FIX`。

### Round 10（Hotspot 失败与竞态边界，Luna Max）

- Sol Medium 复审又指出：Kill 后等待 worker 可能无上限、非超时 worker 非零退出缺少结构化诊断，以及中间 DLL 生成的 100ms 产物会造成 identity 混淆。
- `HotspotProbe.cs` 进一步修正：Kill 后最多等待 `5000 ms`，记录 `exitObserved`；stdout/stderr 读取增加 `1000 ms` 上限并截断；非超时非零退出写入 `hotspot-failure`、summary 状态为 `failed`、探针停止后续场景并以退出码 `3` 结束；超时继续以退出码 `2` 表示 `not-verified`。
- 最终 Release Benchmark build：`0 warning / 0 error`，DLL SHA `81249E5E1247272488B10D5FFA5B7CEB5ABC42AE92AEC118D5F014CA5F28D649`。
- 最终证据：100ms before `hotspot-before-timeout-100ms-final-20260921.json` 为 `6/6 timeout-killed`、退出码 `2`；60s before `hotspot-before-timeout-60s-final-20260921.json` 保留 RawDate 3/10/30、Relation 1K/10K 的完整三次样本，Relation 100K 在 `60061.3724 ms` 超时且 summary 五项统计全为 `null`；after `hotspot-after-final-diagnostic-20260921.json` 六场景各三次通过、退出码 `0`。
- 中间 `4CBB...`/`1156...` 诊断产物不再用于本轮结论；最终诊断与冻结 `manifest-current-20260921.md` 的 `6CBFFF...` 明确隔离。
- SHOULD_FIX 状态：本轮已按审查要求处理，等待下一轮 Sol Medium 独立复审；FIX-002 主项、FIX-003 仍为 `OPEN / MUST_FIX`。

#### Round 10 Sol Medium 复审结果

- 复审结论：`NEEDS_FIX`；Round 10 的三个实现级 `SHOULD_FIX` 已关闭，没有新增 `SHOULD_FIX`。
- 已确认：Kill 后 `5000 ms` 有界等待、stdout/stderr 各 `1000 ms` 诊断上限、非零 worker 的 `hotspot-failure` 与退出码 `3`、timeout 退出码 `2`、不完整 summary 全部 `null`、最终 `81249...` identity 与中间/冻结候选分层。
- 残余风险：非零 worker、Kill 失败和 `exitObserved=false` 尚无独立故障注入产物，仅由控制流和正常/timeout 证据覆盖；这不改变 SHOULD_FIX 关闭结论，也不能替代正式性能门禁。
- FIX-002/FIX-003：继续 `OPEN / MUST_FIX`；任务保持 `PARTIAL`，进度约 `99.5%`。

### Round 11（真实 materializer/binding 可比性核验，Luna Max）

- 对 `MiniExcelRowMaterializer`、`NpoiImportRowMaterializer`、`ProviderComparisonProbe` 和基线提交 `94bb52e84ffc70634067b04541857433ef7af9df` 做了独立可比性核验；没有修改生产代码、Benchmark 代码或 API。
- 基线不包含当前新增的 `MiniExcelRowMaterializer`，旧版物化、绑定、RawDate 和关系逻辑内嵌在 `MiniExcelExcelImporter`；当前 importer 同时发生了职责拆分和多处路径重构，不能把当前 internal 方法反射调用后与基线对应。
- NPOI 的 `NpoiImportRowMaterializer` 在基线已经存在，本候选主要是 `ValidateMode` 合同迁移，未形成可归因的 materializer 算法 before/after。
- `ProviderComparisonProbe` 是完整导出→导入 E2E，`PropertyAccessor` 是 accessor microbenchmark，二者都不能替代隔离的真实 materializer/binding 热路径。未伪造数字，FIX-002 继续保留 `NOT_APPLICABLE_BEFORE/NOT_VERIFIED`。
- 若要补齐该项，需要在基线和当前候选各自使用同一固定数据集完成三类相互配套、分别可审计的证据：固定工作簿的 import-only E2E（覆盖真实读取、映射和物化）、固定行集的 export-only binding E2E（覆盖真实映射计划、绑定和写出），以及针对 materializer/binding 生产路径的归因 microbenchmark。三类证据均必须分别绑定 baseline candidate manifest 和 current candidate manifest，并保持数据集、TFM、runtime、指标和重复次数一致；否则继续保持 `NOT_APPLICABLE_BEFORE/NOT_VERIFIED`。

### Round 12（ResourceProbe 参数化与 500K 公开入口补充）

- `ResourceProbe` 的 `--staging-entrypoints` 现支持可选正整数 `rowCount`；省略参数保持 `1000/10000/100000` 默认矩阵，指定参数只运行该行数，header 的 `rowCounts`、命令参数和 `scenarios` 与请求保持一致。
- CLI 与 `StagingEntrypointMatrix.Run` 均执行 `1..1_048_575` 边界校验；`0`、负数、`1048576`、`int.MaxValue` 和溢出输入均在创建产物前返回退出码 `2`。参数化短 smoke 产物 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-1-parameter-final.json` 的 `rowCount=1` 覆盖四入口各三次，共 `12/12 PASS`，输出重开/解析和 cleanup 均通过；该文件只是本地 smoke，不是资源门禁证据。
- 在 Windows Job Object 2 CPU 等效/4 GiB、`timeoutMs=900000`、隔离 TEMP 条件下完成 `rowCount=500000` 的 `TempFile` 四公开入口各三次，共 `12/12 PASS`。runner 产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entrypoints-500k-parameter-final.json`，入口矩阵产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-500k-parameter-final.json`；`peakJobMemoryBytes=2098900992`，所有输出均可重开/解析，`leftover/temp after=0`，隔离 TEMP cleanup=`deleted`。
- 旧的 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-500k-parameter.json` 仅包含 header 和 3 条 `Export` 记录，预期 `12`、实际 `3`，应标记为 `3/12 PARTIAL, NOT_EVIDENCE`；正式证据仅使用上述带 `-final.json` 后缀的文件，glob/人工选取必须排除旧文件。
- 本轮只补充 `TempFile` 公开入口的参数化证据，不是完整 `Memory/TempFile/Hybrid × 场景 × 并发` 矩阵，也不覆盖生产环境、外部 CI 或正式性能/容量阈值；FIX-002/FIX-003 继续保持 `OPEN / MUST_FIX`，任务状态仍为 `PARTIAL`。

### Round 13（1M 公开入口边界证据）

- 使用参数化命令 `dotnet tests/Bing.Offices.ResourceProbe/bin/Release/net8.0/Bing.Offices.ResourceProbe.dll --staging-entrypoints artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-1m-parameter-final.json 1000000`，在 Windows Job Object 2 CPU 等效、4 GiB、`timeoutMs=1800000` 和隔离 TEMP 下运行；runner 产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entrypoints-1m-parameter-final.json`，入口产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-1m-parameter-final.json`。
- 入口产物完整记录四个公开入口各 3 次，共 12 条，全部为 `failed`；runner 为 `child-failed`、exit `1`。所有失败均发生在 `TryReopenWorkbook` 的 NPOI DOM 工作簿重开阶段，异常为 `OutOfMemoryException`；最大输出约 `20,261,647` bytes，`peakJobMemoryBytes=3301044224`，`leftover/temp-after=0`，隔离 TEMP cleanup=`deleted`。
- 该结果仅是 1M 入口 roundtrip 的结构化 `NOT_VERIFIED`/容量失败证据，不是 PASS，不替代既有 SXSSF 单场景 1M PASS，也不构成产品支持决策；完整策略矩阵、生产环境、外部 CI 和正式阈值仍未验证。FIX-002/FIX-003 继续保持 `OPEN / MUST_FIX`，任务状态仍为 `PARTIAL`。

### Round 14（Round 13 独立只读审查，Sol Medium）

- Sol Medium 对 Round 13 的 1M 公开入口证据执行独立只读审查，结论为无新增 `MUST_FIX` 或 `SHOULD_FIX`。
- 审查确认入口产物完整记录 12 条，即四入口各 3 次；全部失败于 NPOI DOM `TryReopenWorkbook` 的 `OutOfMemoryException`，runner 为 `child-failed`、exit `1`，`peakJobMemoryBytes=3301044224`，隔离 TEMP cleanup=`deleted`。
- `FIX-003` 仍为 `OPEN / MUST_FIX`；1M roundtrip 继续为 `NOT_VERIFIED`，该失败证据不替代 SXSSF 单场景 1M PASS，也不改变完整资源矩阵、生产环境、外部 CI 和正式阈值仍未验证的结论。

### Round 15（基线 Relation before 受限证据）

- 从基线提交 `94bb52e84ffc70634067b04541857433ef7af9df` 导出隔离副本，Benchmark Release 构建为 `0 warning/0 error`；使用 Windows JobRunner 2 CPU 等效/4 GiB、60 秒超时和隔离 TEMP。产物为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-hotspot-before-baseline-60s-final.json` 与 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-baseline-60s-final.json`。
- probe header/assembly identity 绑定 `phase=before`、`repetitions=3`、基线 numeric-cell replay 说明和 Benchmark/Abstractions/Core/MiniExcel 程序集物理 SHA；RawDate 3/10/30 与 Relation 1K/10K 各 3 次完成。Relation 100K 未完成，runner 为 `runner-timeout`/exit `-1`，隔离 TEMP 为 `0/0` 且 cleanup=`deleted`。
- Relation 100K before 现在有真实基线 Job Object 超时证据，但仍缺完整 before 数字与正式阈值，不能写成 PASS 或补填统计；FIX-002 仍为 `OPEN / MUST_FIX`，任务状态保持 `PARTIAL`。

### Round 16（基线构建日志独立审计）

- Round 15 的基线构建证据现有独立日志 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-baseline-build-final.log`；该日志绑定隔离 `.tmp-baseline-94bb52e` 和源提交 `94bb52e84ffc70634067b04541857433ef7af9df`，命令为 `dotnet build .tmp-baseline-94bb52e/benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore --nologo --no-incremental`，`ExitCode=0`，并明确记录 `0 个警告/0 个错误`。
- 该日志补齐了 Round 15 基线构建结果的独立审计链；Relation 100K before 仍只有 Job Object `runner-timeout`/exit `-1` 的边界证据，缺少完整 before 数字与正式阈值，因此不写成 `PASS`。`FIX-002`、`FIX-003` 继续保持 `OPEN / MUST_FIX` 与 `NOT_VERIFIED`，任务状态仍为 `PARTIAL`。

### Round 17（500K staging 单并发局部矩阵）

- 当前 Release `ResourceProbe` 在 Windows Job Object 2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP、500K 行、并发 `1`、预热 `1` 次/测量 `2` 次条件下，完成全部 9 个 `strategy × scenario` c1 组合；原始 runner JSON 均位于 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/`，文件名前缀为 `windows-job-staging-500k-` 且以 `-c1-final.json` 结尾。
- 三个 `excel-100k` 策略（Memory/TempFile/Hybrid）均 `PASS`，两个测量操作无错误且输出约 `424 MB`；Job 峰值分别为 `1,422,483,456`、`1,480,302,592`、`1,499,906,048` bytes。TempFile/Hybrid 临时磁盘峰值约 `212 MB`；所有 staging 残留均为 `0`，隔离 TEMP 均 `0/0/deleted`。
- 三个 `failure-double-dom` 策略均 `child-failed`，Job 峰值分别为 `3,262,435,328`、`3,254,591,488`、`3,254,464,512` bytes；每项两个 measurement 均在 NPOI 失败工作簿注释/VML 序列化阶段 OOM。三个 `template-image-style` 策略均 `child-failed`，Job 峰值分别为 `3,247,964,160`、`3,248,381,952`、`3,247,976,448` bytes；每项两个 measurement 均在 NPOI SharedStrings/ZIP 写出阶段 OOM。上述 staging 残留均为 `0`，隔离 TEMP 均成功清理；这些结果只记录受限容量失败，不作为产品支持决策。
- 本轮完整覆盖 500K c1 的 `3 策略 × 3 场景`，但不等同完整 `3 策略 × 3 场景 × 4 并发` 500K/1M 矩阵，不覆盖取消/失败提交、生产机器、外部 CI 或正式阈值。FIX-003 继续 `OPEN / MUST_FIX`、资源门禁为 `NOT_VERIFIED`，整体状态保持 `PARTIAL`。

### Round 18（Round 17 Sol Medium 独立复核）

- Sol Medium 对 Round 17 执行只读独立复核，确认 9 个 `windows-job-staging-500k-*-c1-final.json` 均满足 `requestedCpuEquivalent=2`、`requestedMemoryLimitBytes=4294967296`、`timeoutMs=1800000`、隔离 TEMP `0/0/deleted`，报告中的 Job 峰值、输出字节、临时磁盘峰值和清理后值与原始 JSON 一致。
- 复核确认三个 `excel-100k` 策略均 `PASS/0`，六个 `failure-double-dom`/`template-image-style` 组合均为 `child-failed`、退出码 `1/1`、`operationBytes=0`、`errorCount=2`；失败分别落在 NPOI 注释/VML 序列化和 SharedStrings/ZIP 写出阶段 OOM，均不是 runner timeout。
- 复核确认文档只声明 500K、并发 1、`3 策略 × 3 场景 = 9/9` 局部矩阵，未把它扩大解释为并发 4/16/64、完整 500K/1M、取消/失败提交、生产机器、外部 CI 或正式阈值证据；未新增 `MUST_FIX` 或 `SHOULD_FIX`。FIX-003 继续 `OPEN / MUST_FIX`，资源门禁继续 `NOT_VERIFIED`，整体状态保持 `PARTIAL`。

### Round 19（500K staging 请求并发 4 Excel 局部补充）

- Luna Max 使用当前 Release `ResourceProbe` 在 Windows Job Object 2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP、500K 行、请求并发 `4`、预热 `1` 次/测量 `2` 次条件下，独立完成 Memory/TempFile/Hybrid 三个 `excel-100k` 场景；三个 runner/child 均 `passed`、退出码 `0`、`errorCount=0`。
- Job 峰值分别为 `1,428,307,968`、`1,194,053,632`、`1,507,565,568` bytes；输出分别为 `1,696,533,421`、`1,696,533,420`、`1,696,533,421` bytes；TempFile/Hybrid 临时磁盘峰值均为 `212,066,678` bytes，Memory 为 `0`；所有 child 残留和 runner 隔离 TEMP 均为 `0/0/deleted`。
- 三项 `maxActualParallelism=1`，因此本轮只扩展了请求并发为 4 的局部证据，不能表述为四路同时执行。它仍不覆盖 failure/template 场景、完整 500K `3×3×4`、1M、取消/失败提交、生产机器、外部 CI 或正式阈值；FIX-003 继续 `OPEN / MUST_FIX`，资源门禁保持 `NOT_VERIFIED`。

### Round 20（Round 19 Sol Medium 独立复核）

- Sol Medium 对 Round 19 执行只读独立复核，确认 Memory/TempFile/Hybrid 三个 500K `excel-100k` c4 runner/child 均为 `passed/0/0`，Job 峰值、输出字节、临时磁盘峰值、`operationCount=8`、`errorCount=0` 和清理字段与原始 runner JSON 一致。
- 三项均满足 2 CPU 等效、4 GiB、`1800000ms`、`runnerTimedOut=false`、请求并发 `4`，且 `maxActualParallelism=1`；复核确认文档没有把请求并发 4 扩大解释为四路同时 DOM 执行，也没有把 c4 Excel `3/3` 扩大为完整 500K 矩阵。
- 复核结论为通过，无新增 `MUST_FIX` 或 `SHOULD_FIX`。Runner JSON 的 child 结果已完整嵌入 stdout，独立 child 文件未作为报告引用；三份 JSON 无末尾 LF 仅属低优先级格式观察，不影响本轮结构化证据。FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 21（500K staging 请求并发 4 失败/模板局部补充）

- Luna Max 在同一当前 Release、2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP、500K 行、请求并发 `4`、预热 `1` 次/测量 `2` 次条件下，使用正确策略枚举完成 `failure-double-dom` 与 `template-image-style` 的 Memory/TempFile/Hybrid 六项；六个 runner/child 均为 `child-failed`、退出码 `1/1`、`operationBytes=0`、`errorCount=8`，实际并行度均为 `1`。
- 六项 Job 峰值依次为 `3,270,365,184`、`3,276,005,376`、`3,277,504,512`、`3,249,311,744`、`3,249,360,896`、`3,249,115,136` bytes；TempFile/template 的临时磁盘峰值为 `2,459` bytes，其余为 `0`；所有 staging 残留和 runner 隔离 TEMP 均为 `0/0/deleted`。failure 场景错误落在注释/VML 序列化，template 场景落在 SharedStrings/ZIP 写出，均为 OOM 容量边界而非 runner timeout。
- 首次错误调用因使用小写 `memory` 策略参数而以 CLI exit `2` 结束，未纳入六项证据；正式表格仅引用正确枚举值生成的 `final2` runner JSON。当前 500K c4 局部证据为 Excel `3/3 PASS` + failure/template `6/6` 结构化 OOM；仍不等同完整 `3×3×4` 发布矩阵，不覆盖 1M、取消/失败提交、生产机器、外部 CI 或正式阈值。FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 22（Round 21 Sol Medium 独立复核）

- Sol Medium 对 Round 21 执行只读独立复核，确认六份正确枚举值生成的 `c4-final2` runner JSON 均为 `child-failed/1/1`、`runnerTimedOut=false`，满足 500K、请求并发 `4`、实际并行度 `1`、2 CPU 等效、4 GiB、`1800000ms`、预热 1 次/测量 2 次。
- 复核确认 `operationBytes=0`、`errorCount=8`、Job 峰值、临时磁盘峰值 `2459/0`、清理后 `0/0` 与报告一致；failure OOM 落在 Comments/VML，template OOM 落在 SharedStrings/ZIP。小写 `memory` 的 CLI exit `2` 调用未被计入正式证据。
- 复核确认报告只声明 c4 failure/template `6/6` 局部容量边界，未扩大为完整 500K `3×3×4`、1M、生产机器、外部 CI 或正式阈值；未新增 `MUST_FIX` 或 `SHOULD_FIX`。FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 23（500K/c16 failure/template 未形成结构化产物）

- 尝试使用同一当前 Release、500K 行、请求并发 `16`、2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP 运行第一个 failure 场景；在 runner/child 生成 JSON 前进程消失，目录未出现任何 `500k c16` runner artifact。
- 没有可靠的退出码、Job 峰值、错误数量或 cleanup 字段，原因无法由现有证据确定；未手动终止，也未继续启动其余五项。该尝试被明确排除，不作为 PASS、FAIL 或容量结论，c16 failure/template 继续 `NOT_VERIFIED`。
- 本轮只增加边界可追溯记录，不改变 500K c1/c4 局部证据或 FIX-002/FIX-003 状态；整体仍为 `PARTIAL`。

### Round 24（Round 23 Sol Medium 独立复核）

- Sol Medium 对 Round 23 执行只读独立复核，检查资源目录中不存在任何 `500k` 与 `c16` 同时匹配的 runner artifact，并确认文档没有为无产物试跑推断退出原因或容量结论。
- 复核确认 500K c1 与 c4 正式证据各 9 份、合计 18 份，均可解析且未被 c16 试跑覆盖；两份 Markdown 的 UTF-8/LF 和 `git diff --check` 均通过，未新增 `MUST_FIX` 或 `SHOULD_FIX`。
- c16 failure/template 继续 `NOT_VERIFIED`；FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 25（Release build/test 收口复核）

- 当前候选执行 `dotnet build Bing.Offices.sln -c Release --no-restore --nologo`，双 TFM 生产与测试项目均成功，`0 个警告 / 0 个错误`。
- 随后执行 `dotnet test Bing.Offices.sln -c Release --no-build --no-restore --nologo`，共 `1756/1756` 通过、`0` 失败、`0` 跳过；覆盖通用、MiniExcel、NPOI、Integration、Docs 测试的 net6/net8 目标。
- 本轮只复核现有代码与证据链，没有改变 FIX-002/FIX-003 的门禁判断；A-I 正式阈值、完整资源/生产/外部环境仍未验证，整体保持 `PARTIAL`。

### Round 26（500K/c16 60 秒有界超时诊断，16 GiB 探索性条件）

- Luna Max 使用当前 Release、500K 行、请求并发 `16`、Memory/failure-double-dom、2 CPU 等效/**16 GiB**、隔离 TEMP 和 `timeoutMs=60000` 执行有界诊断，生成 `windows-job-staging-500k-memory-failure-double-dom-c16-timeout60-final.json`；原始 runner 参数已核对为 `memoryGiB=16`、`cpuCount=2`。
- 结果为 `runner-timeout`，runner/child exit `-1/-1`、`runnerTimedOut=true`，Job 峰值 `5,372,166,144` bytes；child 未形成 JSON/stdout，但终止后隔离 TEMP 为 `0/0/deleted`。该证据只覆盖 16 GiB 条件下的超时、终止和清理链路，不是 4 GiB 容量 PASS/FAIL，也不替代 c16 完整矩阵。
- FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体状态保持 `PARTIAL`。

### Round 27（500K/c16 4 GiB 有界超时诊断修正）

- Sol 复核 Round 26 时发现原始 `timeout60-final.json` 的 runner 参数实际为 16 GiB；该产物已在报告中明确标记为探索性条件，不再误称 4 GiB。
- Luna Max 随后按正确顺序 `<runner> 4 2 ...` 重做 500K/c16 Memory/failure-double-dom `60000ms` 诊断，生成 `windows-job-staging-500k-memory-failure-double-dom-c16-timeout60-4g-final.json`。结果为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，`requestedMemoryLimitBytes=4294967296`，Job 峰值 `3,235,536,896` bytes，child 无 JSON/stdout，隔离 TEMP 为 `0/0/deleted`。
- 该修正产物只证明 4 GiB 条件下的有界超时和清理链路，不是 c16 容量 PASS/FAIL；FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 28（Round 27 Sol Medium 独立复核）

- Sol Medium 确认新 4 GiB runner JSON 与报告完全一致：`runner-timeout`、exit `-1/-1`、`runnerTimedOut=true`、`timeoutMs=60000`、内存上限 `4,294,967,296`、Job 峰值 `3,235,536,896` bytes，child 无 JSON/stdout，隔离 TEMP 为 `0/0/deleted`。
- 复核确认旧 16 GiB 探索性产物与新 4 GiB 产物已明确分层，文档没有把有界超时写成容量 PASS/FAIL，也没有宣称完成 c16 矩阵；未新增 `MUST_FIX` 或 `SHOULD_FIX`。
- FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 29（500K/c16 Excel 4 GiB 有界超时诊断）

- Luna Max 使用当前 Release、500K 行、请求并发 `16`、Memory/excel-100k、2 CPU 等效、4 GiB、隔离 TEMP 和 `timeoutMs=60000` 运行有界诊断，生成 `windows-job-staging-500k-memory-excel-100k-c16-timeout60-4g-final.json`。
- 结果为 `runner-timeout`，runner/child exit `-1/-1`、`runnerTimedOut=true`，Job 峰值 `1,399,291,904` bytes，child 无 JSON/stdout；终止后清理前为 `1` 个临时文件、`29,843,456` bytes，最终 `cleanup=deleted`。
- 该证据补充了 c16 Excel 超时期间临时文件可观测性，但不构成容量 PASS/FAIL，也不替代完整取消/失败提交矩阵；FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 30（Round 29 Sol Medium 独立复核）

- Sol Medium 确认 c16 Excel runner JSON 字段与报告一致：`runner-timeout`、exit `-1/-1`、`runnerTimedOut=true`、`timeoutMs=60000`、4 GiB、Job 峰值 `1,399,291,904` bytes、空 stdout/stderr、终止后 `1` 个临时文件/`29,843,456` bytes、最终 `deleted`。
- 复核确认 child 参数为 `Memory / excel-100k / c16 / 500000`，child JSON 不存在；文档仅声明超时和清理可观测性，未认定容量 PASS/FAIL，也未替代取消/失败提交矩阵。未新增 `MUST_FIX` 或 `SHOULD_FIX`。
- FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 31（500K/c16 TempFile/Hybrid Excel 有界超时诊断）

- Luna Max 使用当前 Release、500K 行、请求并发 `16`、TempFile/Hybrid `excel-100k`、2 CPU 等效、4 GiB、隔离 TEMP 和 `timeoutMs=60000` 分别执行有界诊断；两项均为 `runner-timeout`，runner/child exit `-1/-1`，`runnerTimedOut=true`，child 无 JSON/stdout。
- TempFile Job 峰值 `1,191,579,648` bytes，终止后清理前为 `2` 个临时文件/`75,202,560` bytes；Hybrid Job 峰值 `1,208,631,296` bytes，终止后清理前为 `1` 个临时文件/`32,239,616` bytes；两项最终隔离 TEMP 均 `cleanup=deleted`。
- 该轮只补充 c16 不同 staging 策略的超时残留和清理可观测性，不构成容量 PASS/FAIL，也不替代完整取消/失败提交矩阵；FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 32（Round 31 Sol Medium 独立复核）

- Sol Medium 确认 TempFile/Hybrid 两份 c16 Excel runner JSON 均为 `runner-timeout`、exit `-1/-1`、`runnerTimedOut=true`、`timeoutMs=60000`、4 GiB，Job 峰值分别为 `1,191,579,648` 与 `1,208,631,296` bytes。
- 复核确认终止后清理前临时文件分别为 `2/75,202,560` 与 `1/32,239,616`，最终均 `deleted`；stdout/stderr 和 child JSON 均为空/不存在。文档仅声明超时残留与清理可观测性，未认定容量 PASS/FAIL，也未替代完整取消/失败矩阵；未新增 `MUST_FIX` 或 `SHOULD_FIX`。
- FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 33（500K/c16/4 GiB failure 与 template 边界证据）

- Luna Max 使用当前 Release `ResourceProbe`，按正确入口 `childPath=dotnet`、首个 child 参数为 `ResourceProbe.dll`，在 Windows Job Object `requestedCpuEquivalent=2`、`requestedMemoryLimitBytes=4,294,967,296`、隔离 TEMP、`timeoutMs=1800000`、500K 行、请求并发 `16`、预热 `1` 次/测量 `2` 次条件下完成 `failure-double-dom` 与 `template-image-style` 六项 `c16-long-final2` 运行。三份无效的早期 DLL 直启结果保留为 `*.invalid-launch.json`，stderr 为 `System.Runtime 8.0.0.0` 加载失败，不计入正式证据。
- 三个 `failure-double-dom` 产物均为 `child-failed`、runner/child exit `1/1`、`runnerTimedOut=false`，每项 `submitted/completed/operation/errorCount=32/32/32/32`、`maxActualParallelism=1`、`maxQueuedRequests=15`、`operationBytes=0`；Job 峰值分别为 `3,271,938,048`、`3,277,131,776`、`3,277,570,048` bytes。错误均为 NPOI Comments/VML 失败工作簿序列化链路中的 `System.OutOfMemoryException`，child 临时磁盘和遗留文件为 `0`，runner 隔离 TEMP 为 `0/0/deleted`。
- 三个 `template-image-style` 产物均为 `child-failed`、runner/child exit `1/1`、`runnerTimedOut=false`，同样完成 `32/32` 请求且 `operationBytes=0`、`errorCount=32`、实际并行度 `1`、排队 `15`；Job 峰值分别为 `3,252,965,376`、`3,252,240,384`、`3,252,502,528` bytes。每项均记录 `BingOfficesExportException` 包裹的 NPOI `System.OutOfMemoryException`，临时磁盘峰值分别为 `0`、`2,461`、`0` bytes，child 遗留文件为 `0`，runner 隔离 TEMP 为 `0/0/deleted`。
- 本轮只增加 500K/c16 在 4 GiB 预算下的真实失败和清理边界证据；failure/template OOM 不是产品支持通过，且不替代完整 `3 策略 × 3 场景 × 1/4/16/64` 矩阵、取消中途提交保护、1M、生产机器、外部 CI 或批准阈值。FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体状态保持 `PARTIAL`。

### Round 34（Round 33 Sol Medium 独立复核）

- Sol Medium 对 9 份正式 `windows-job-staging-500k-*-c16-long-final2.json`（3 个 Excel 超时、3 个 failure OOM、3 个 template OOM）和 3 份 `*.invalid-launch.json` 执行只读独立复核；确认正式文件均为 `childPath=dotnet`、首个 child 参数为 `ResourceProbe.dll`、4 GiB、2 CPU、`timeoutMs=1800000`、500K/c16、Job 已分配且隔离 TEMP `cleanup=deleted`，无正式/无效产物重叠。
- 复核确认三个 Excel 为 `runner-timeout/-1/-1` 且无 child summary；failure/template 均为 `child-failed/1/1`，每项 `32/32/32/32`、最大实际并行度 `1`、最大排队 `15`，failure 错误落在 NPOI Comments/VML，template 错误落在 NPOI SharedStrings/ZIP；无效启动 stderr 为 `System.Runtime 8.0` 加载失败，未计入正式统计。复核未新增 `MUST_FIX` 或 `SHOULD_FIX`。
- 当前证据仍只是 500K/c16 局部边界和清理可观测性，不能替代完整资源矩阵、取消中途提交保护、1M、生产机器、外部 CI 或正式阈值；FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 35（Relation 100K before 长时边界诊断及 Sol Medium 复核）

- Luna Max 在隔离 `.tmp-baseline-94bb52e` Release Benchmark DLL 上使用 Windows Job Object `2 CPU`、`4 GiB`、隔离 TEMP、外层 `timeoutMs=900000`、Hotspot worker `600000ms`、`before`、3 repetitions 执行 Relation 100K before 长时诊断。runner 为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，Job 峰值 `389,754,880` bytes，隔离 TEMP `0/0/deleted`，stdout/stderr 为空。
- child 产物绑定基线 Benchmark DLL SHA-256 `65AEA685D40829C81F3221B88DE813EC4B782521EAFE4E616873B5736B63BF25`，四个程序集物理 identity 与隔离基线一致；已刷盘 15 个完整样本和 5 个 summary：RawDate 3/10/30、Relation 1K/10K 各 3 次。Relation 100K 仍为 0 样本、0 summary，输出在 Relation 10K summary 后结束，未填充伪造统计。
- Sol Medium 只读复核确认 child 身份、资源参数、样本数、超时和清理字段一致，`beforeSource` 仍为 investigative numeric-cell replay，不是 HEAD Reader baseline；未新增 `MUST_FIX` 或 `SHOULD_FIX`。长时诊断只加强 `NOT_VERIFIED` 边界证据，FIX-002 仍 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 36（当前工作树 Release build/test 复核）

- 当前工作树执行 `dotnet build Bing.Offices.sln -c Release --no-restore --nologo`，双 TFM 生产、测试、Benchmark 和工具项目均成功，`0` 个警告、`0` 个错误。
- 随后执行 `dotnet test Bing.Offices.sln -c Release --no-build --no-restore --nologo`，共 `1756/1756 PASS`、`0` 失败、`0` 跳过；Core 双 TFM 各 `197/197`、MiniExcel 双 TFM 各 `42/42`、NPOI 双 TFM 各 `566/566`、Integration 双 TFM 各 `29/29`（MiniExcel）/`30/30`（NPOI）、Docs `10/10`。
- 本轮仅验证当前代码回归和双 TFM 测试门禁，没有改变 FIX-002/FIX-003 的判断；Relation 100K before、正式阈值、完整 500K/1M 资源矩阵、生产机器和外部 CI 仍未验证，整体保持 `PARTIAL`。

### Round 37（Materialization/Binding 公开路径 smoke 及 SHOULD_FIX 复核）

- Luna Max 新增 `--materialization-binding-probe <artifact> <rows> <repetitions>`，通过公开 `IExcelExporter`/`IExcelImporter` API 和公开 NPOI/MiniExcel DI 注册，分别执行 sync/async 导出→导入往返；固定 seed `20260920`、六列显式映射，记录导出/导入耗时、分配量、GC、峰值工作集、输出字节、行数、导入错误和字段级完整性。
- 当前候选 Release build 为 `0` 个警告、`0` 个错误；smoke 产物为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/materialization-binding-smoke-v3.json`，包含 NPOI/MiniExcel × sync/async 共 `4/4` 个 `measured` 样本，均为 `3` 行导出、`3` 行导入、`0` 个导入错误、`roundTrip=true`、字段完整性通过。
- v3 产物的文档级 candidate identity 已覆盖 Benchmark、Core、Abstractions、NPOI、MiniExcel 五个 Release 程序集，并与物理 DLL SHA-256 一致；`schema=1`、`4` 个 summary、`evidenceStatus=measured`、`thresholdStatus=UNAPPROVED`。
- Sol Medium 第二次只读复核确认此前 identity 遗漏和文档措辞 `SHOULD_FIX` 已关闭，未新增 `MUST_FIX`/`SHOULD_FIX`。该证据仍是小规模、合并的 after-only smoke，不能隔离 E/F 热路径，也不能替代同候选 before/after、100K/500K/1M 扩展性或正式阈值；FIX-002/FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 38（1M staging Excel-only c1 三策略局部证据及 Sol Medium 复核）

- Luna Max 使用当前工作树 Release `WindowsJobRunner`/`ResourceProbe`，在 Windows Job Object `2 CPU` 等效、`4 GiB`、隔离 TEMP、`timeoutMs=1800000`、请求并发 `1`、预热 `1` 次/测量 `2` 次条件下完成 `1M × excel-100k × Memory/TempFile/Hybrid` 三项。三项 runner/child 均 `passed / 0 / 0`，每项 `submitted/completed/operation/errorCount=2/2/2/0`，实际并发 `1`，排队 `0`。
- 三项输出分别为 `848,270,818`、`848,270,818`、`848,270,819` bytes；Job 峰值分别为 `2,771,984,384`、`2,886,086,656`、`2,334,699,520` bytes；TempFile/Hybrid 临时峰值分别为 `424,135,409`、`424,135,410` bytes，Memory 为 `0`；三项终止后均为 `0` 文件、`0` 字节、`cleanup=deleted`。产物为 `windows-job-staging-1m-memory/tempfile/hybrid-excel-100k-c1-final.json`，child `staging-scenario` 结果嵌入各 runner `stdout`，与当前 runner 设计一致。
- 本轮运行时物理 identity 已记录在 `resource-report.md`：WindowsJobRunner、ResourceProbe 和 NPOI DLL 的 SHA-256 与本次运行文件一致；不把这些局部产物重新归属到其他冻结候选。Sol Medium 只读复核确认字段、约束、退出码、输出和清理状态一致，未新增 `MUST_FIX`/`SHOULD_FIX`。
- 该轮只补充 `1M × Excel-only × c1 × 3 strategies` 的本地 PASS/清理证据，不能替代 failure/template、并发 `4/16/64`、取消/失败提交、Entity/Template 1M、生产机器、外部 CI 或正式阈值；FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 39（1M staging failure/template c1 边界证据及 Sol Medium 复核）

- Luna Max 在与 Round 38 相同的当前 Release、Windows Job Object `2 CPU` 等效/`4 GiB`、隔离 TEMP、`timeoutMs=1800000`、请求并发 `1`、预热 `1` 次/测量 `2` 次条件下，完成 `1M × c1 × 3 strategies × 2 scenarios` 六项运行。六项 runner/child 均为 `child-failed`、退出码 `1/1`、`runnerTimedOut=false`，每项 `submitted/completed/operation/errorCount=2/2/2/2`、`operationBytes=0`、实际并发 `1`、排队 `0`、隔离 TEMP `0/0/deleted`。
- `failure-double-dom` 的 Memory/TempFile/Hybrid 三项均由 `ExcelXlsxZipPreflight` 因 `xl/worksheets/sheet1.xml` 解压大小超过默认限制而结构化拒绝，异常为 `Bing.Offices.Exceptions.BingOfficesResourceLimitException`；`template-image-style` 三项均在 NPOI 写出路径抛出 `System.OutOfMemoryException`。前者是资源限制拒绝，后者是 OOM，不能合并为容量通过结论。Job 峰值分别为 failure `303,833,088 / 309,362,688 / 301,772,800` bytes，template `3,247,915,008 / 3,248,037,888 / 3,247,992,832` bytes；六项 child 临时磁盘峰值均为 `0`，清理均成功。
- 六份原始 runner JSON 位于 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/`，child `staging-scenario` 结果嵌入 runner `stdout`。Sol Medium 对原始产物执行只读复核，确认约束、字段、异常分类和清理状态一致，未新增 `MUST_FIX`/`SHOULD_FIX`。本轮仅补充 1M failure/template 的局部边界失败证据，不覆盖其他并发、取消中途提交、Entity/Template 1M roundtrip、生产机器、外部 CI 或正式阈值；FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 40（1M staging Excel-only c4 三策略局部证据及 Sol Medium 复核）

- Luna Max 使用当前 Release `WindowsJobRunner`/`ResourceProbe`，在 Windows Job Object `2 CPU` 等效、`4 GiB`、隔离 TEMP、`timeoutMs=1800000`、请求并发 `4`、`rowCount=1000000`、预热 `1` 次/测量 `2` 次条件下完成 `Memory`、`TempFile`、`Hybrid` 三项 `excel-100k`。三项 runner/child 均为 `passed / 0 / 0`，每项 `submitted/completed/operation/errorCount=8/8/8/0`、`operationBytes` 分别为 `3,393,083,275 / 3,393,083,276 / 3,393,083,277`，`runnerTimedOut=false`。
- 三项实际 DOM 并发均为 `1`、最大排队为 `3`；Job 峰值分别为 `2,776,682,496 / 2,879,967,232 / 2,900,910,080` bytes，child 临时峰值分别为 `0 / 424,135,410 / 424,135,410` bytes；child 结束后临时文件/字节均为 `0/0`，runner 隔离 TEMP 均为 `0/0/deleted`。c4 仅表示请求并发，不能表述为四路实际并行。
- 三份原始 runner JSON 位于 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/`，child 结果嵌入各自 `stdout`。Sol Medium 只读复核确认 `childPath=dotnet`、ResourceProbe 参数、4 GiB/2 CPU、输出、清理和候选边界一致，未发现数据或执行链路问题。本轮仅补充 `1M × Excel-only × c4 × 3 strategies` 的本地 PASS/清理证据，不覆盖其他并发、failure/template、取消/失败提交、Entity/Template 1M、生产机器、外部 CI 或正式阈值；FIX-003 继续 `OPEN / MUST_FIX`，整体保持 `PARTIAL`。

### Round 41（Materialization/Binding 分离 import/export 证据）

- 当前候选执行 `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore`，退出码 `0`、`0` 个警告、`0` 个错误。新增探针仍只使用公开 `IExcelExporter`/`IExcelImporter` 与公开 Provider 注册；固定 seed `20260920`、六列 explicit mapping，并将 `export-only`、`import-only`、`roundtrip` 三种 operation 与 `sync/async` 分开计时。
- 10K 正式产物 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/materialization-binding-current-10k-v4.json`：`rowCount=10000`、`repetitions=3`、`36/36 measured`、`0` failed、`0` import errors；schema `2`，dataset identity 为 `C59A33C19CA9F35D5B20F14D599021B8BB0675885D43523D1552BF7AEA9E4B33`。
- 100K 正式产物 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/materialization-binding-current-100k-v4.json`：`rowCount=100000`、`repetitions=3`、`36/36 measured`、`0` failed、`0` import errors；schema `2`，dataset identity 为 `1F8469C53ECF45121C866368A07CFFC0DAAD01929764E8D4E5D09DB343A60FE6`。每个 Provider 的 `import-only` 重复使用同一预生成 fixture，fixture/output SHA-256、字段完整性和 XLSX workbook/sheet 结构均由 probe 自验证。
- 两个产物均明确 `baselineStatus=NOT_APPLICABLE_BEFORE` 及原因：当前工作区没有可运行且 identity 独立的 before candidate，不能用不存在的 before 或其他 List/workbook 路径补数；本轮只证明当前候选的隔离 E2E materializer/binding 路径。`thresholdStatus=UNAPPROVED`，不作性能保留/回滚或发布 PASS 判断。
- 本轮补齐了 FIX-002B 所需的真实 materialization/binding `import-only` 与 `export-only` 分离、代表性 10K/100K 行数和重复统计；当前候选隔离 E2E 为 `FIX-002B=CLOSED`，before 明确为 `NOT_APPLICABLE_BEFORE`。同候选 before/after 与正式阈值归类为 `FIX-002C=BLOCKED_APPROVAL`；Relation 100K 等基线边界归类为 `FIX-002A=VERIFIED_BOUNDARY`。资源/本地边界归类为 `FIX-003A/B=VERIFIED_BOUNDARY`，本地 staging/cleanup 收口为 `FIX-003C=CLOSED`，生产机器与外部 CI 为 `FIX-003D/E=BLOCKED_EXTERNAL`；Implementation=`CLOSED`、Test=`CLOSED`、Performance=`BLOCKED_APPROVAL`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`STOPPED`，没有 OPEN_ACTIONABLE，不能标记 `COMPLETED`。

### Round 42（Sol 分类 docs-only 收口）

- 本轮仅更新 `benchmark-report.md` 与 `execution.md` 的状态分类、当前摘要、Round 41 结论和最终 TODO；未修改源码、测试、版本文件、`review.md` 或任何性能/资源产物，未运行同质实验。
- FIX-002 现按 `FIX-002A=VERIFIED_BOUNDARY`、`FIX-002B=CLOSED`、`FIX-002C=BLOCKED_APPROVAL` 记录；其中 A-I 完整正式阈值和同候选 before/after 明确是审批阻断，不能被表述为 OPEN_ACTIONABLE。
- FIX-003 现按 `FIX-003A/B=VERIFIED_BOUNDARY`、`FIX-003C=CLOSED`、`FIX-003D/E=BLOCKED_EXTERNAL` 记录。整体状态为 Implementation=`CLOSED`、Test=`CLOSED`、Performance=`BLOCKED_APPROVAL`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`STOPPED`；没有 OPEN_ACTIONABLE，仍不能标记 `COMPLETED`，等待 Sol Medium 复审。

### Round 43（Performance Budget v1 五场景确认与最终候选 smoke）

- 用户批准 `Performance Budget v1` 后，仅执行一次五目标精准确认：NPOI async provider、MiniExcel sync provider、`excel-file-sync`、`excel-file-async`、`excel-throttled-async`，均为 100K、5 repetitions；结构化总产物为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/targeted-performance-budget-v1-5rep.json`。五项均 `5/5` 完整且 `REGRESSION_OVER_BUDGET`，中位回退分别为 `+191.15%`、`+433.80%`、`+337.08%`、`+445.69%`、`+136.12%`。本轮没有重跑完整 A-I、Relation 100K、RawDate hotspot、tail latency 或 c16/c64。
- 本轮确认把 `FIX-002C` 从审批阻断重分类为 `OPEN_ACTIONABLE`：批准预算已验证为实际回退，下一步需要围绕这五项回退定位并实施最小性能修复，不能用提高阈值或重标候选身份规避。旧 `FIX-002A=VERIFIED_BOUNDARY`、`FIX-002B=CLOSED` 保持不变。
- 重新构建当前工作树 Release solution：`dotnet build Bing.Offices.sln -c Release --no-restore --nologo --disable-build-servers -m:1 -p:UseSharedCompilation=false`，`0` warning、`0` error；WindowsJobRunner/ResourceProbe Release build 同样 `0/0`。当前目标产物绑定 Benchmark `419DC8A4F2D7907BEBFFE8351A6AD671FCDD552397CF0C4134F4B824DFC1202F`、NPOI `E87FA5777867FB1D8B22B824EB580373306C75B95F311447F554BE70390F010E`、MiniExcel `3FDDC6AF4EE936A44E783B1208CA6B33C59705419585D8A353110980C4D7DCEB`。
- 在同一 Windows Job Object 2 CPU 等效/4 GiB/隔离 TEMP profile 下，最终候选代表性 `TempFile/excel-100k/c1` smoke 的 100K、500K、1M 均 `passed`，每项 `2/2` 操作完成、错误 `0`、runner/child exit `0`、清理 `0/0/deleted`；Job 峰值分别为 `399,478,784`、`1,224,552,448`、`2,357,735,424` bytes。该结果是本地生产等价 profile 证据，不能写成真实生产机或外部 CI PASS；FIX-003D 更新为 `LOCAL_PROFILE_EVIDENCE`，FIX-003E 仍 `BLOCKED_EXTERNAL`。
- 按 FIX-003A 批准保持 `MaxWorksheetBytes=64 MiB`；Entity/Template 1M 未重跑，继续记录 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`，没有提高默认限制。

## Git 状态

当前工作树保留用户/任务变更，未执行任何 destructive git 操作。未执行 `git add`、`commit`、`push`、`tag`、PR 或发布操作。

## 进度：约 99.5%（IN_PROGRESS，非 COMPLETED）

当前状态汇总：Implementation=`CLOSED`、Test=`CLOSED`、Performance=`OPEN_ACTIONABLE`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`IN_PROGRESS`；OPEN_ACTIONABLE 仅为 `FIX-002C` 的五项已确认 100K E2E 回退。

## TODO（最终）

- [x] 获取 `ValidateMode` rename 的成员级 API approval；基于当前 immutable candidate 更新双 TFM baseline 并重跑 Public API tests，双 TFM compare 与 Public API 专项均通过。
- [x] 为已新增 Entity/Template/capability 公共成员获取成员级 API approval；当前候选新增成员已纳入审批和 baseline，Public API 专项 net6/net8 各 `9/9 PASS`。
- [x] 补齐 Entity/Template 的样式、图片、复杂 merge、边界和大明细职责级测试；Entity 专项双 TFM 已扩展至 `22/22`，包含固定 Cell mapping/validation/value-map 与 List Region 边界、文件端点预取消和中途取消时的目标保护/临时文件清理。
- [x] 编译并运行新增 MiniExcel/NPOI 非 `IList` `ICollection<T>` 关系合同测试；双 TFM 通过，NPOI 全量 `566/566`、MiniExcel 全量 `42/42`；大型 SXSSF Merge/Formatter/取消回归也已通过。
- [x] 补充当前候选公开 Materialization/Binding 证据；历史 v3 smoke 的 `4/4` 合并往返仍保留，Round 41 新增固定 seed/fixture identity 的 `export-only`、`import-only`、`roundtrip` 分离探针，10K 与 100K 各 `36/36` measured、字段完整性通过。仍明确 `NOT_APPLICABLE_BEFORE` 与 `UNAPPROVED`，不宣称正式阈值或 before/after 结论。
- [OPEN_ACTIONABLE] FIX-002C 性能回退修复：Performance Budget v1 五个批准目标均已完成 `5/5`，但 NPOI async、MiniExcel sync、`excel-file-sync`、`excel-file-async`、`excel-throttled-async` 仍超过 `25%`；需在不重跑 Relation 100K/c16/c64/完整 A-I 的前提下定位并实施最小修复，再做职责级回归和同一候选确认。FIX-002A=`VERIFIED_BOUNDARY`、FIX-002B=`CLOSED`，其余技术边界不变。
- [BLOCKED_EXTERNAL] FIX-003D/E 外部门禁：本轮已补齐最终候选在本地 2 CPU/4 GiB Job Object profile 的代表性 100K/500K/1M smoke（均 PASS、清理 `0/0/deleted`），但这不是真实生产机器，也不是外部 CI；完整 500K/1M 矩阵、取消/失败提交、Entity/Template 1M roundtrip 和外部 Build/Test/API/Pack/Consumer Gate 仍需外部环境。FIX-003A/B=`VERIFIED_BOUNDARY`、FIX-003C=`CLOSED`、FIX-003D=`LOCAL_PROFILE_EVIDENCE`、FIX-003E=`BLOCKED_EXTERNAL`。
- [x] 清理或隔离既有 `artifacts/**/artifacts`、`artifacts/**/bin`、`artifacts/**/obj` 历史产物，并重新核对根 `artifacts/` 目录治理门禁；当前扫描结果为 `nested_generated_dirs=0`。
- [x] 将刷新候选的测试、benchmark、resource、package 和 API 产物统一绑定到同一 immutable candidate；当前 build/package/API capture、Provider/Real IO/Hotspot/Entity after v2 和 100K c64 资源证据已由 `manifest-current-20260921.md` 封存，旧候选产物仍明确标记为历史记录。
- [x] 处理独立 `review.md` 的 FIX-001；用户批准、baseline、compare 和 Public API 门禁均已闭合。
- [x] 更新独立 `review.md` 的当前机器状态：顶层已改为 `AI_REVIEW_STATUS: BLOCKED`，不再保留可被机器误判的顶层 `NEEDS_FIX`；旧审查文本仅保留为历史证据。当前 `FIX-002C=OPEN_ACTIONABLE`，`FIX-003D=LOCAL_PROFILE_EVIDENCE`、`FIX-003E=BLOCKED_EXTERNAL`，性能修复和外部门禁完成前不能标记 `COMPLETED`。
- [x] 修复 Sol Medium 新增的 `FIX-005`/`FIX-006` 实现缺口；Sol Medium 已完成最终独立复审并确认 `FIX-005`/`FIX-006` 已关闭。
