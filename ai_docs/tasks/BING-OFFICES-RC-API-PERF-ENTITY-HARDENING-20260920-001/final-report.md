# 当前收口结论

当前候选约完成 `99.5%` 的既有整改，整体状态为 `IN_PROGRESS`（非 `COMPLETED`）。既有 `ValidateMode` Breaking rename、MiniExcel List/RawDate/Relation 优化与职责拆分、immutable Entity layout、NPOI Entity/Template 执行链、最小 Provider capability SPI、公开包第三方 fixture、双 TFM 回归和专项测试均保持；本轮又补齐了批准的 Performance Budget v1 五场景 100K/5-repetition 精准确认和最终候选本地生产等价 profile 的 100K/500K/1M smoke。确认结果显示五个目标均超过 25% budget，当前有一个明确的 `FIX-002C OPEN_ACTIONABLE`，不能把状态写成已停止。

当前状态汇总：Implementation=`CLOSED`、Test=`CLOSED`、Performance=`OPEN_ACTIONABLE`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`IN_PROGRESS`、Open Actionable=`1`。FIX-002A=`VERIFIED_BOUNDARY`，FIX-002B=`CLOSED`，FIX-002C=`OPEN_ACTIONABLE`；FIX-003A/B=`VERIFIED_BOUNDARY`、FIX-003C=`CLOSED`、FIX-003D=`LOCAL_PROFILE_EVIDENCE`、FIX-003E=`BLOCKED_EXTERNAL`。Entity/Template 1M 仍按默认 `MaxWorksheetBytes=64 MiB` 记录为 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`；真实生产机和外部 CI 仍 `NOT_VERIFIED`，因此不能标记 `COMPLETED`。旧 review 文本中的顶层状态只作历史快照，当前机器状态以后续独立复审为准。未执行 commit、push、tag、PR 或 NuGet 发布。

## Round 43 新增证据与 TODO

- Performance Budget v1 artifact：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/targeted-performance-budget-v1-5rep.json`。NPOI async、MiniExcel sync、`excel-file-sync`、`excel-file-async`、`excel-throttled-async` 均 `5/5`，中位回退分别 `+191.15%`、`+433.80%`、`+337.08%`、`+445.69%`、`+136.12%`，均为 `REGRESSION_OVER_BUDGET`。
- 本地生产等价 profile runner artifacts：`windows-job-production-equivalent-smoke-100k.json`、`windows-job-production-equivalent-smoke-500k.json`、`windows-job-production-equivalent-smoke-1m.json` 及对应 child JSON。三项均 `passed`、2 次测量、0 错误、TEMP `0/0/deleted`；Job 峰值分别约 `399 MB`、`1.22 GB`、`2.36 GB`。这不是真实生产机或外部 CI 结果。
- [OPEN_ACTIONABLE] 定位并修复五个已确认的 100K E2E 回退，保持用户批准的 25% budget，不重跑 Relation 100K、c16/c64 或完整 A-I；修复后只补职责级回归和同一 candidate identity 的精准复测。
- [BLOCKED_EXTERNAL] 执行一次 Frozen Candidate 外部 CI 完整 Build/Test/API/Pack/Consumer Gate，并取得真实生产等价环境的外部证据；本地 smoke 不能替代该门禁。

# 生产符号到测试追溯

| 生产符号/行为 | 直接测试 |
|---|---|
| `ExcelValidationFailureMode`、`ExcelSheetImportBuilder.Validate` | `Bing.Offices.Npoi.Tests.ExcelP0RegressionTest.ValidationFailureMode_ShouldContinueCollectingErrors`、`Bing.Offices.Npoi.Tests.StreamPipelineTest.Import_ContinueValidation_ShouldCollectAllErrors` |
| `ExcelTypeMapFactory` 动态列最多一个 | `Bing.Offices.Npoi.Tests.StreamPipelineTest.TypeMap_MultipleDynamicColumnProperties_ShouldThrowArgumentException` |
| `MiniExcelRawDateSerialReader` 行列范围 | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest.RawDateSerialReader_ShouldIndexNumericCellsAndRestorePosition`、`DataRowStartIndex_ShouldAlignRawDateSerials`、`DataRowStartIndexAsync_ShouldAlignRawDateSerials` |
| MiniExcel materialization / `ICollection<T>` | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest` list/dynamic round-trip；`Bing.Offices.MiniExcel.Tests.Integration.MiniExcelProviderContractTest` |
| NPOI relation/navigation collection filtering and read-only scalar guard | `Bing.Offices.Npoi.Tests.ExcelWorkbookRequestTest.Import_RelationWithNonListCollectionNavigation_ShouldBindChildren`、`Bing.Offices.Npoi.Tests.StreamPipelineTest.Import_ReadOnlyMappedProperty_ShouldThrowConfigurationException`、`Bing.Offices.Npoi.Tests.ExcelP0RegressionTest.Import_FixedImageColumnMultiplicity_ShouldApplyConfiguredPolicy` |
| MiniExcel relation cached invoker/index | `Relations_ShouldHonorCustomKeyComparer`、`Relations_ShouldPreserveFirstMatchingParent`、`Relations_WithNoChildren_ShouldNotEvaluateParentKeys`、`Relations_ShouldStopAfterFirstMatchingParentBeforeEvaluatingLaterParents`、`Relations_ShouldPreserveRepeatedParentKeyEvaluationAndErrorBoundary` |
| `ExcelXlsxZipPreflight` netstandard 字符串兼容性 | `Bing.Offices.Npoi.Tests.NpoiXlsxZipPreflightTest` 及完整 NPOI 测试矩阵 |
| `ExcelEntityLayout<TEntity>` 固定 Cell/Merge/List、多 Sheet | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ShouldRoundTripFixedCellsAndListRegion`、`EntityLayout_ShouldSnapshotBuilderAndMapping`、`EntityLayout_ShouldRejectInvalidAddressesAndOverlappingRegions` |
| Entity Template、merge preflight、HSSF、异步取消、流释放、样式/图片/复杂 merge/大明细边界 | `EntityTemplate_ShouldPreserveMergeAndValidateTemplateShape`、`EntityTemplate_Hssf_ShouldRoundTrip`、`EntityTemplate_ShouldDisposeTemplateWhenLeaveOpenIsFalse`、`EntityAsync_ShouldRoundTripAndHonorCancellation`、`EntityTemplate_ShouldPreserveStylePictureAnchorAndOriginalTemplate`、`EntityTemplate_ShouldRejectNonExactMergeConflictsBeforeWriting`、`EntityListRegion_LargeDetail_ShouldRoundTripAtExactBoundary` |
| Entity 固定字段命名转换器 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.EntityCell_ShouldUseNamedConverter` |
| Core Entity 扩展 25/25 直接调用追溯 | `Bing.Offices.Tests.PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`、`EntityExtensions_ShouldRejectNullProvidersThroughDirectCalls` |
| MiniExcel Entity/Template unsupported fail-fast | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest` Entity fail-fast tests |
| 非 `IList` `ICollection<T>`/`HashSet<T>` 关系导航合同 | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest.Relations_ShouldBindNonListCollectionNavigation`、`Bing.Offices.Npoi.Tests.ExcelWorkbookRequestTest.Import_RelationWithNonListCollectionNavigation_ShouldBindChildren`；双 TFM 均通过 |
| public-only Provider SPI、能力 preflight、取消、流和文件端点 | `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs`；net6/net8 输出 `third-party-public-only-provider-ok` |

# 测试与构建结果

| 项目 | TFM | 结果 |
|---|---|---|
| `Bing.Offices.Tests` | net6.0/net8.0 | `197/197 PASS` 各一轮 |
| `Bing.Offices.Npoi.Tests` | net6.0/net8.0 | `566/566 PASS` 各一轮；Entity 专项双 TFM 已通过 |
| `Bing.Offices.MiniExcel.Tests` | net6.0/net8.0 | `42/42 PASS` 各一轮 |
| `Bing.Offices.Npoi.Tests.Integration` | net6.0/net8.0 | `30/30 PASS` 各一轮 |
| `Bing.Offices.MiniExcel.Tests.Integration` | net6.0/net8.0 | `9/9 PASS` 各一轮 |
| `Bing.Offices.Docs.Tests` | net8.0 | `10/10 PASS` |
| `Bing.Offices.Tests.Integration` | net6.0/net8.0 | `39/39 PASS` 各一轮 |

API 分类、生产程序集 IVT、NPOI exact-member 和 Core 公开扩展覆盖门禁均通过。Release solution build 为 0 warning/0 error；版本文件未修改，`build/api-snapshot-baseline.json` 已按 FIX-001 用户批准更新。PackageReference consumer 与第三方 public-only fixture 均使用全新 `--packages artifacts/consumer/cache-rc-final7-*` 验证修复后的重打包 `artifacts/packages`，且包内 DLL 与当前 Release 输出物理 SHA-256 一致，未混用旧或用户全局同号缓存。`artifacts/tests/` 保存本轮 TRX。

新增的非 `IList` `ICollection<T>`/`HashSet<T>` 关系合同测试已写入 MiniExcel 与 NPOI 测试项目，并在隔离缓存下双 TFM 通过；同时修复 NPOI 对非图片可枚举关系集合的错误固定列预检，保留只读标量异常和图片集合导入合同。

修复后的 NPOI 集成 DLL 已重新执行，net6/net8 均通过；NPOI 全量单元为 `566/566 PASS`，MiniExcel 全量单元为 `42/42 PASS`。新增大型列表 round-trip、属性级 Merge 排除和取消提交保护测试双 TFM 均通过。

资源补充：新增 Windows Job Object runner，宿主机 22 逻辑处理器下以 `910/10000` CPU rate units 近似 2 CPU，并设置 4 GiB 作业内存上限；当前候选无模板纯列表 SXSSF 场景的 100K/500K/1M 均通过，Entity/Template 100K/500K 单次运行也通过。Entity/Template 1M 当前在默认 `ExcelResourceLimits.MaxWorksheetBytes=64 MiB` 的 ZIP 预检阶段因 `sheet2.xml` 超限退出，未生成 roundtrip 产物，仍为 `NOT_VERIFIED`。1M staging Excel-only 的 c1 与 c4 三策略均已通过，其中 c4 的请求并发为 `4`、实际 DOM 并发为 `1`、队列为 `3`；failure/template c1 仅形成资源限制拒绝和 OOM 边界证据，不能按容量通过处理。公开入口、失败注入及 1K/10K staging 结构矩阵已有结构化证据；完整 500K/1M 策略矩阵、取消/失败提交、生产机器和外部 CI 仍未验证。

Benchmark 补充：Provider 100K、Real IO 24 场景、RawDate 3/10/30 列、Relation 1K/10K、DynamicPlan（8 组合）、PropertyAccessor（4 组合）、GenericSheetDispatch（4 组合）、Entity/Template after E2E 和 cold-plan-build tail latency（并发 1/4/16/64、5 次重复）已有方向性 before/after 观测；这些 microbenchmark 不等于正式 BenchmarkDotNet 通过或已批准阈值。旧 after 证据使用历史候选 SHA；Round 43 当前候选使用 Benchmark DLL SHA `419DC8A4F2D7907BEBFFE8351A6AD671FCDD552397CF0C4134F4B824DFC1202F`，并新增 `targeted-performance-budget-v1-5rep.json`。该精准确认的五个 100K E2E 目标全部超过批准的 25% budget，当前 `FIX-002C=OPEN_ACTIONABLE`；完整 A-I 仍未重跑。Relation 100K before 被安全停止，作为 `FIX-002A=VERIFIED_BOUNDARY` 边界证据；Entity/Template before 因基线缺少同一公共 API 记为 `NOT_APPLICABLE_BEFORE`。Entity/Template 1M 当前重跑在默认 worksheet XML 资源限制处拒绝且未生成 roundtrip 产物，不能填充性能数字。DynamicPlan 的 BenchmarkDotNet ShortRun 在自动生成项目 restore 阶段因 `NU1301` 失败，后续 InProcess 结果只作方向性归因，不能宣称发布性能门禁通过。GenericSheetDispatch Import 的精确探针记录了 NPOI `XSSFFactory.CreateDocumentPart` 触发并被捕获的 `MissingMethodException`，before/after 计数一致且导入结果正确，不作为无异常门禁。完整数据见 `benchmark-report.md`、`before-candidate-manifest.md`、`targeted-performance-budget-v1-5rep.json` 及既有方向性产物。

本轮尝试的 `DynamicPlanBenchmarks` BenchmarkDotNet `ShortRun` 在自动生成项目 restore 阶段遇到 NuGet SSL/凭据错误 `NU1301`，结果为 `NA`，因此 dynamic plan 性能项仍为 `NOT_VERIFIED`。Entity/Template after 探针使用受控公开 API 路径完成，但基线缺少相同公共 API，before 明确为 `NOT_APPLICABLE_BEFORE`。

# API 差异与未完成门禁

当前 API diff 包含 `ExcelValidationFailureMode` rename，以及 `ExcelEntity*` layout/result/template、`ExcelEntityExtensions`、`IExcelEntity*` 和 `ExcelProviderCapabilities`。FIX-001 已按用户批准更新 baseline 并通过双 TFM compare/Public API 门禁；独立审查见 `review.md`，其历史内容不覆盖当前 Round 43 分类。当前 FIX-002A/B/C 分别为 `VERIFIED_BOUNDARY`/`CLOSED`/`OPEN_ACTIONABLE`，FIX-003A/B/C/D/E 分别为 `VERIFIED_BOUNDARY`/`VERIFIED_BOUNDARY`/`CLOSED`/`LOCAL_PROFILE_EVIDENCE`/`BLOCKED_EXTERNAL`。

剩余 TODO：

- [x] 获取全部 Breaking/API 新增成员的成员级审批，更新双 TFM snapshot 并重跑公共 API 门禁；FIX-001 双 TFM compare 与 Public API 专项通过。
- [x] 补齐 Entity/Template 样式、图片、复杂 merge、边界、大明细和文件取消/原子提交职责级矩阵；本轮独立 review 已复核证据。
- [x] 编译并运行新增 MiniExcel/NPOI 非 `IList` `ICollection<T>` 关系合同测试；双 TFM 通过，NPOI 全量 `566/566`、MiniExcel 全量 `42/42`；大型 SXSSF Merge/Formatter/取消回归也已通过。
- [OPEN_ACTIONABLE] FIX-002C：定位并修复五个已确认的 100K E2E budget 回退，然后只对职责级回归和同一 candidate identity 做精准复测；Relation 100K、c16/c64 和完整 A-I 仍按批准边界不重跑。FIX-002A=`VERIFIED_BOUNDARY`、FIX-002B=`CLOSED`，Entity/Template before 保持 `NOT_APPLICABLE_BEFORE`。
- [BLOCKED_EXTERNAL] FIX-003D/E：本地 2 CPU/4 GiB Job Object profile 的 100K/500K/1M 代表性 smoke 已通过，但完整 500K/1M 矩阵、取消/失败提交、生产机器和外部 CI 仍未验证；FIX-003A/B=`VERIFIED_BOUNDARY`、FIX-003C=`CLOSED`、FIX-003D=`LOCAL_PROFILE_EVIDENCE`、FIX-003E=`BLOCKED_EXTERNAL`。Entity/Template 1M 继续 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`。
- [x] 清理或隔离既有 `artifacts/**/artifacts`、`artifacts/**/bin`、`artifacts/**/obj` 历史产物，并重新核对根 `artifacts/` 目录治理门禁；当前扫描结果为 `nested_generated_dirs=0`。
- [x] `review.md` 的顶层机器状态已在独立 Round 43 复审后更新为 `AI_REVIEW_STATUS: BLOCKED`，不再保留顶层 `NEEDS_FIX`；历史审查文本保留为证据。FIX-001 已关闭，五项性能回退和外部门禁未收口前不能标记 `COMPLETED`。

最新资源矩阵补充：在原始 100K staging runner 的 `10/36` 记录之后，当前候选使用最新 ResourceProbe 依赖独立重跑六个 c64 场景，四项在 `3600000ms` 并行窗口完成，两个 Excel 场景的并行窗口 timeout 记录随后由 `7200000ms` 串行当前候选 runner 替代，最终六项均完成 `128/128`、child exit `0`、错误为 `0`、隔离 TEMP 清理后为 `0`。当前合并覆盖为 `36/36 PASS`；旧 `-rerun` 证据因依赖哈希不一致仅保留为门限观察。审批字段仍为空，因此该结果是本地结构/清理证据，不替代批准的 500K/1M、生产和外部 CI 门禁。

API 身份复核补充：早先 `artifacts/api/current-20260921` capture 与旧候选 manifest 的 source hash、物理 Release DLL 不一致；本轮已重新 build/pack/capture，当前刷新记录位于 `artifacts/api/current-20260921-candidate` 和 `manifest-current-20260921.md`，source、Release DLL、包与 API capture 已内部一致。当前刷新候选已补充双 TFM 测试、PackageReference、public-only、Provider/Real IO/Hotspot/Entity after 和 100K c64 资源关联；旧候选历史 baseline 未被用于替代当前批准 baseline，仍需完整 A-I/500K/1M 证据和外部环境门禁。
