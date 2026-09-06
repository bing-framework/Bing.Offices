# 执行进度

## 2026-09-06 11:26 +08:00

- 状态：Phase 0 `IN_PROGRESS`；Phase 1–6 `TODO`。
- Git：`master`，HEAD `9d78ab76e28891ac9b7e0e558f4df669721a4ea6`；执行前工作区干净。
- 修改：创建本任务计划、执行状态和 Phase 0 证据文件；未修改生产源码。
- 行为变化：无。
- 命令：`dotnet restore Bing.Offices.sln --nologo`（受控网络权限后成功）；`dotnet build Bing.Offices.sln -c Release --no-restore --nologo`；三类 `dotnet test` 基线；`rg` 调用链与弃用引用扫描。
- Build：Release 通过，0 error / 28 warnings。
- Unit：netcoreapp3.1、net6.0、net8.0 均 453 passed / 1 failed / 0 skipped；唯一失败为正式 Public API hash 与当前程序集不一致。
- Integration：net6.0、net8.0 均 15/15；Docs：net8.0 11/11。
- 未完成：完整 API 快照/批准、异常边界修复、日期导出合同、弃用删除、全量测试与性能发布证据。
- 阻塞：指定最新评审文件和 API snapshot baseline JSON 不存在；当前可继续依据源码和最近评审推进。
- 下一步：修复 `AtomicFileCommitter` 写委托误分类、图片 Try/Throw 合同并添加直接测试。

## 状态总览

| 范围 | 状态 | 证据 |
| --- | --- | --- |
| Phase 0 基线 | IN_PROGRESS | `baseline.md`、TRX、各矩阵 |
| Phase 1 正确性 | TODO | 已确认 Atomic write 误分类、TryAddPicture catch-all 等可达 |
| Phase 2 弃用/API | TODO | `deprecated-removal.md` |
| Phase 3 重构 | TODO | 尚未开始 |
| Phase 4 测试/包消费 | TODO | 仅完成修改前基线 |
| Phase 5 性能/资源 | TODO | 尚未开始 |
| Phase 6 文档/Review/门禁 | TODO | 尚未开始 |

## 2026-09-06 11:35 +08:00

- 状态：Phase 1 `IN_PROGRESS`。
- 生产变更：`AtomicFileCommitter` 区分内容生成与文件提交阶段；`TryAddPicture` 补全参数校验并仅吞明确可恢复异常；`MovePictures` 对 null/未知 Sheet 显式失败；Excel/CSV 默认以 ISO round-trip 文本导出 `DateTimeOffset`；禁止无固定 offset 的 `DateTime -> DateTimeOffset` 隐式转换。
- 定向验证：Atomic/file export 9/9；NPOI picture 4/4（XSSF/HSSF）；DateTimeOffset 4/4（XLSX/XLS/CSV/parser），均为 net8 Release。
- 未完成：多 TFM 全量回归、Observer 跨 file entry 同实例集成、未知 Sheet 可测试替身、配置 loader 决策、failure workbook 诊断与资源策略。
- 下一步：全量回归后继续配置 Loader 与 failure workbook。

## 2026-09-06 11:37 +08:00

- 回归：net8 Unit 460 passed / 1 failed / 0 skipped；唯一失败仍为基线中的 Public API hash，不是本轮行为变更新增失败。
- Integration：net6.0 与 net8.0 各 15/15，0 skipped。
- Artifact：`artifacts/test-results/unit/phase1-unit-net8.trx`、`artifacts/test-results/integration/phase1-integration_*.trx`。
- 当前判定：Phase 1 继续 `IN_PROGRESS`，不更新/放宽正式 API hash。

## 2026-09-06 12:00 +08:00

- Phase 1：默认 DI Mapping Loader 已接入 Observer；静态 Loader 保持纯解析。独立合同测试 4/4。
- Phase 1：失败工作簿 `Hidden/ZeroHeight/Collapsed` 的 `NotImplementedException` 不再静默吞掉，改发 `FailureWorkbookRowMetadataUnsupported` 结构化诊断；相关回归 4/4。
- Phase 1：`ExcelResourceLimits.MaxInputBytes` 默认 128 MiB且允许显式 null 关闭；XLS/OLE 与 XLSX 均在 DOM 前执行输入预算，XLSX 保留 ZIP/XML 专项预检；资源矩阵 31/31。
- Phase 2：已删除 DataTable `Bing.Offices.CsvHelper`、两个专属兼容测试及 API 台账条目；强类型 CSV/API 分类测试 40/40，全仓仅剩第三方 `CsvHelper.Configuration` 引用。
- 当前状态：Phase 1 `IN_PROGRESS`（跨 TZ 与未知 Sheet 直接测试仍待补）；Phase 2 `IN_PROGRESS`。
- 下一步：迁移并删除 `ExcelMapping.For<T>`、旧 builders、runtime v1 migration 与 LegacyCompatibility/v1 Benchmark。

## 2026-09-06 12:15 +08:00

- Phase 2：删除 `ExcelMapping`、`ExcelMappingBuilder<T>`、`ExcelColumnMappingBuilder<T,TProperty>`；所有有效调用迁移到方向化 Builder，Export builder 补齐实际主链需要的 `HasConverter`/`Map`。
- Phase 2：删除全部 public `MigrateV1Json/Xml`、v1 私有解析器、迁移诊断专属测试、LegacyCompatibility test 和 v1 Benchmark；JSON/XML 安全、UTF-8、converter、validation 测试改用原生 v2。
- 验证：Release solution build 0 errors；方向化 Builder 7/7；v2 配置主链 6/6；Docs 10/10；Phase 2 Integration net6/net8 各 15/15。
- 全仓扫描：弃用符号只剩迁移指南中的删除项名称，无生产、测试、Benchmark、生成 XML 引用。
- 已知失败：Phase 2 首次 Unit 全量除既有 API hash 外，发现一个 v1 形状安全测试，修复后定向通过；仍需重跑三 TFM 全量。
- 下一步：public execution-detail/SPI 治理、API 成员快照与全量 Unit。

## 2026-09-06 12:25 +08:00

- Phase 2 API：删除无消费者的 `ExpressionExtension`、`RegexConst`、`ExcelValueMap<T>`；`PropertyInfoExtensions`/`TypeExtensions` internal；六个仅由 Core 工厂构造的默认 validation rule internal。
- Phase 2 SPI：`MappingConfigurationMerger`、`MappingProfileRegistry`、`UniqueTracker`、`BingOfficesExceptionDispatcher`、`ExcelValidationRules`、`DateTimeExcelValidationRule` 明确归类 Provider SPI 并要求 `EditorBrowsable(Never)`。
- 三 TFM Unit：netcoreapp3.1/net6/net8 均 463 passed / 1 failed / 0 skipped；唯一失败为正式 API hash。
- Artifact：`artifacts/test-results/unit/phase2-api-unit_*.trx`。
- 下一步：将剩余用户可配置 DTO/Attributes/Options 从错误的 Execution detail 分类为 User API；生成候选 API snapshot 与 Breaking diff。

## 2026-09-06 12:40 +08:00

- Public 分类完成：141 User API、9 Provider User API、15 Provider SPI、0 Compatibility、0 Execution detail；分类测试通过。
- ApiSnapshot CLI 增加显式 `--dependencies`，三 TFM 候选快照成功生成；Abstractions/Core/NPOI 成员数 775/152/66。
- 历史 formal baseline JSON 缺失，正式 hash 未更新，成员级审批状态保持 BLOCKED；详见 `api-diff.md`。
- Phase 3：导入/导出 plan builder 与 Sheet executor 已按泛型类型缓存 open-instance delegate，移除每 Sheet `MethodInfo.Invoke/object[]`；异构多 Sheet与异常回归 4/4。
- 下一步：跨 TZ 日期验证、delegate cache Benchmark/测试、剩余反射路径与职责拆分。

## 2026-09-06 12:55 +08:00

- Phase 1 日期：`TZ=UTC` 与 `TZ=Pacific Standard Time` 的 DateTimeOffset 专项各 4/4，TRX 已保存。
- Phase 3：四处 per-Sheet 泛型反射替换为按类型缓存的 open-instance delegate；无每次 `MethodInfo.Invoke/object[]`，多 Sheet回归 4/4。
- Benchmark：新增 `GenericSheetDispatchBenchmarks`，公开 Import/Export 真实链 1/8 Sheet ShortRun 4/4；结果与限制记录于 `benchmark-report.md`。
- ApiSnapshot 工具：修复 `--dependencies` 显式依赖解析；清理由本任务失败调用遗留的 3 个锁文件进程。
- 下一步：剩余反射路径评估、Phase 3 职责拆分、完整 Release 回归与 PackageConsumer。

## 2026-09-06 12:34 +08:00

- Phase 3：`NpoiRelationBinder` 改为按 Workbook/Parent/Child/Key 类型缓存强类型委托；每次关系绑定不再使用 `MakeGenericMethod/Invoke/object[]`。直接缓存未命中/命中与关系异常合同测试合计 18/18。
- Phase 4 PackageConsumer：三个本轮 `2.0.0` nupkg 打包成功；独立隔离缓存恢复成功；`netstandard2.0` 合同消费者编译通过；`netcoreapp3.1`、`net6.0`、`net8.0` 均运行输出 `package-consumer-ok`。
- Package assets：Bing 依赖全部为 `type=package`，无生产 `ProjectReference`；三包均包含 DLL/XML/nuspec/LICENSE/README。详见 `package-consumer-report.md`。
- Release solution build：0 errors / 28 warnings；警告仍为既有 netcoreapp3.1 支持、nullable、obsolete、analyzer 与 Benchmark 废弃属性。
- 当前状态：Phase 1 正确性核心项 `VERIFIED`；Phase 2 删除/API 收敛 `DONE`（正式成员 baseline 缺失，API approval `BLOCKED`）；Phase 3 缓存与局部职责拆分 `DONE`；Phase 4 PackageConsumer `VERIFIED`，最终分报告仍待补。
- 下一步：运行并扩充 ResourceProbe，补齐 Unit/Integration/Docs/Resource 分报告和最终全量回归。

## 2026-09-06 12:45 +08:00

- Phase 1 Provider：新增 `MovePictures` null/未知 `ISheet` 直接失败测试，`BingOfficesUnsupportedFeatureException` 的 Code/Operation/Stage 均锁定，1/1 通过。
- Phase 2 API：`BingOfficesException` 改为 abstract；全仓无直接构造调用，直接契约测试 1/1。候选快照 canonicalizer 不编码 abstract 修饰符，hash 不变，该工具缺口已写入 `api-diff.md`。
- Phase 4 最终回归：Unit 三 TFM 各 464/465，唯一失败为既有 formal API hash；Integration net6/net8 各 15/15；Docs 更新后三章后仍 10/10；均 0 skipped。
- Phase 5 Resource：Excel 独立进程 14/14；Mapping/Unique 子进程 16/16；尾延迟 4 档完成。无批准预算、Failure Workbook 双 DOM 和完整 1k/10k/100k Excel 资源矩阵，状态保持 `PARTIAL / UNAPPROVED`。
- 报告：已生成 `unit-test-report.md`、`integration-test-report.md`、`docs-test-report.md`、`package-consumer-report.md`、`resource-report.md`。
- 下一步：最终重建/重打包并重新隔离消费，随后独立 Review、修复发现并完成 No-Go 门禁报告。

## 2026-09-06 13:08 +08:00

- 独立 Review Round 1：0 P0 / 6 P1 / 2 P2，状态 `NEEDS_FIX / No-Go`；高风险定向 26/26。
- Review Fix：Failure Workbook 无 sink 与 sink 失败均写结构化 Trace；新增直接测试 2/2。DateTimeOffset 新增 XLS/XLSX 两阶段重开往返，双 TZ 各 6/6。
- Review Fix PackageConsumer：重新 build/pack 到 `packages-reviewfix`，第三套全新缓存 restore；netstandard2.0 编译和三 runtime build/run 通过；消费者覆盖统一异常 catch、MovePictures unknown 和 TryAddPicture 参数失败；包内五个 DLL 与当前 Release 输出 hash 全部相等。
- Benchmark：新增 CSV 公开调用链 1k/10k/100k；完成 Excel 9、CSV 6、Failure Workbook 3 个 ShortRun 场景。100k Allocated：Excel Import 1,684.74 MB、Excel Export 734.51 MB、CSV Import 198.32 MB、CSV Export 1,448.26 MB、Failure Workbook 2,606.14 MB；明确不满足低 GC 描述，预算仍未批准。
- Review Fix 全量：Release build 0 errors / 28 warnings；Unit 三 TFM 各 470/471，唯一失败 formal API hash；Integration net6/net8 各 15/15；Docs 10/10；全部 0 skipped。
- 下一步：等待独立复审；仅保留正式 API 成员 baseline/审批、性能资源预算、Phase 3 更大规模职责拆分与跨 OS 基础设施等真实门禁。

## RC Final 状态总览

| 范围 | 状态 | 最终证据 |
| --- | --- | --- |
| Phase 0 基线/矩阵 | VERIFIED | baseline 与调用链/异常/日期/API/弃用台账 |
| Phase 1 正确性 | VERIFIED | Unit/Integration、双 TZ、ResourceProbe |
| Phase 2 弃用/API 收敛 | DONE / BLOCKED | 删除与分类 VERIFIED；formal API 成员 baseline/批准 BLOCKED |
| Phase 3 内部重构 | DONE | Sheet/关系委托缓存、Failure 诊断职责拆分；更大类拆分未完成 |
| Phase 4 测试/包消费 | DONE / BLOCKED | Integration/Docs/PackageConsumer VERIFIED；Unit 仅 formal API hash 失败 |
| Phase 5 性能/资源 | DONE / BLOCKED | 实测矩阵完成；历史 baseline 与批准预算 BLOCKED |
| Phase 6 文档/Review/门禁 | DONE / No-Go | 文档和独立 Review 已执行；外部门禁未解除 |

最终冻结证据：Release build 0 errors / 28 warnings；Unit 三 TFM 各 470/471；Integration 两 TFM 各 15/15；Docs 10/10；PackageConsumer 四目标通过且五 DLL hash match；candidate API 位于 `artifacts/api-snapshot/candidate-rc-final`。

## 2026-09-06 13:11 +08:00

- 独立复审：`NEEDS_FIX / No-Go`；开放 P0=0、P1=2、P2=2。FIX-001/003/004/007 已关闭。
- 开放 P1：正式 APICompat baseline/成员级审批；性能资源可比 baseline、批准预算与剩余资源矩阵。
- 开放 P2：大类职责拆分剩余；TryAddPicture 后半段可恢复失败的副作用合同。
- 本地可执行项已完成；最终报告为 `No-Go`，execution 以 `PARTIAL` 收口。
