# 生产符号到测试追溯

| 生产符号/行为 | 直接测试 |
|---|---|
| `ExcelValidationFailureMode`、`ExcelSheetImportBuilder.Validate` | `Bing.Offices.Npoi.Tests.ExcelP0RegressionTest.ValidationFailureMode_ShouldContinueCollectingErrors`、`Bing.Offices.Npoi.Tests.StreamPipelineTest.Import_ContinueValidation_ShouldCollectAllErrors` |
| `ExcelTypeMapFactory` 动态列最多一个 | `Bing.Offices.Npoi.Tests.StreamPipelineTest.TypeMap_MultipleDynamicColumnProperties_ShouldThrowArgumentException` |
| `MiniExcelRawDateSerialReader` 行列范围 | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest.RawDateSerialReader_ShouldIndexNumericCellsAndRestorePosition`、`DataRowStartIndex_ShouldAlignRawDateSerials`、`DataRowStartIndexAsync_ShouldAlignRawDateSerials` |
| MiniExcel materialization / `ICollection<T>` | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest` list/dynamic round-trip tests；`Bing.Offices.MiniExcel.Tests.Integration.MiniExcelProviderContractTest` |
| MiniExcel relation cached invoker/index | `Relations_ShouldHonorCustomKeyComparer`、`Relations_ShouldPreserveFirstMatchingParent`、`Relations_WithNoChildren_ShouldNotEvaluateParentKeys`、`Relations_ShouldStopAfterFirstMatchingParentBeforeEvaluatingLaterParents`、`Relations_ShouldPreserveRepeatedParentKeyEvaluationAndErrorBoundary` |
| `ExcelXlsxZipPreflight` netstandard 字符串兼容性 | `Bing.Offices.Npoi.Tests.NpoiXlsxZipPreflightTest` 及完整 NPOI 测试矩阵 |
| `ExcelEntityLayout<TEntity>` 固定 Cell/Merge/List、多 Sheet | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.Npoi_EntityLayout_ShouldRoundTripFixedCellsAndListRegion`、`EntityLayout_ShouldSnapshotBuilderAndMapping`、`EntityLayout_ShouldRejectInvalidAddressesAndOverlappingRegions` |
| NPOI SXSSF 大型纯列表流式写入、样式、属性级 Merge 排除和取消提交保护 | `Bing.Offices.Npoi.Tests.ExcelWorkbookRequestTest.Export_LargeSimpleXlsx_ShouldRoundTripRowsAndStyles`、`Export_LargeMergedList_ShouldPreserveMergedRegion`、`Export_LargeSimpleXlsx_CancellationShouldPreserveTargetAndCleanTemporaryFile`；当前候选 100K/500K/1M 受限资源 JSON 见 `artifacts/resource/.../windows-job-streaming-*-final.json`，强制终止隔离 TEMP 证据见 `windows-job-streaming-1m-timeout-isolated-final.json` |
| Entity Template、merge preflight、HSSF、异步取消、流释放、样式/图片/复杂 merge/大明细边界 | `EntityTemplate_ShouldPreserveMergeAndValidateTemplateShape`、`EntityTemplate_Hssf_ShouldRoundTrip`、`EntityTemplate_ShouldDisposeTemplateWhenLeaveOpenIsFalse`、`EntityAsync_ShouldRoundTripAndHonorCancellation`、`EntityTemplate_ShouldPreserveStylePictureAnchorAndOriginalTemplate`、`EntityTemplate_ShouldRejectNonExactMergeConflictsBeforeWriting`、`EntityListRegion_LargeDetail_ShouldRoundTripAtExactBoundary` |
| Entity 固定字段命名转换器 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.EntityCell_ShouldUseNamedConverter` |
| Entity 固定 Cell raw/converted 校验顺序与完整错误合同 | `Bing.Offices.Npoi.Tests.EntityLayoutProviderTest.EntityCell_ShouldValidateRequiredBeforeNumericConversion`、`EntityCell_ShouldRunRequestValidationAfterConversion` |
| Core Entity 扩展 25/25 直接调用 | `Bing.Offices.Tests.PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`、`EntityExtensions_ShouldRejectNullProvidersThroughDirectCalls` |
| Provider capability/Entity SPI 声明 | `Bing.Offices.Tests.Integration.PublicApiContractTest.PublicApi_ExportedTypes_ShouldHaveGovernedClassification`、`PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies`、NPOI/MiniExcel Entity tests |
| MiniExcel Entity/Template unsupported fail-fast | `Bing.Offices.MiniExcel.Tests.MiniExcelProviderTest` Entity fail-fast import/export tests |
| public-only Provider SPI、能力 preflight、取消、流和文件端点 | `tests/Bing.Offices.ThirdPartyProvider.Consumer/Program.cs`；net6/net8 `third-party-public-only-provider-ok` |
| Entity/Template 公共 E2E（固定 Cell、Merge、List Region、模板公式/图片/明细、同步/异步） | `benchmarks/Bing.Offices.Benchmarks.EntityProbe`；`entity-template-after-current.json`、`entity-template-after-10k-current.json`、`entity-template-100k-final.json`、`entity-template-500k-final.json`；100K/500K Job Object 受限证据 `windows-job-entity-template-100k-final.json`、`windows-job-entity-template-500k-final.json` |

public-only Provider fixture 已有独立 PackageReference 验证；样式/图片/复杂 Merge/大明细/文件取消保护职责级矩阵已补齐并由 NPOI Entity 专项双 TFM `22/22` 覆盖。早期快照的 `15/15` 不作为当前结果使用。A-I 同候选 before/after 性能追溯仍缺；当前简单列表 SXSSF 的 100K/500K/1M 和 Entity/Template 的 100K/500K 单次运行已有 Job Object 受限资源 PASS，Entity/Template 1M、并发/取消/失败提交、生产/外部资源门禁仍列入执行报告 TODO。
