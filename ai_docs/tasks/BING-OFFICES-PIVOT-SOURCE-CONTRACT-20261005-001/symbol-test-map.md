# 生产符号到直接测试追溯

本阶段不新增公共成员。下表以最终生产调用链为入口，测试项目均为 `Bing.Offices.ProviderContract.Tests`，测试类型为 `PivotTableSourceContractTest`。

| 生产符号 | 关键行为 | 直接测试方法 |
| --- | --- | --- |
| `ClosedXmlPivotSourceValidator.Validate` | 来源任意字段公式在读取和缓存创建前拒绝；支持/未知函数及模板缓存状态一致 | `FormulaInSource_ShouldRejectBeforeEvaluation` |
| `ClosedXmlPivotSourceValidator.Validate` | 空金额/空分组允许，完整来源和全部缓存字段及记录一致；区域外公式保留 | `PureValues_ShouldWriteCompleteSourceAndCache` |
| `ClosedXmlPivotSourceValidator.Validate` / `ClosedXmlReportWriter.ApplyPivotTables` | 数字、日期、布尔及文本分组来源/缓存类型保留，字段名称忽略大小写匹配 | `TypedGroups_ShouldPreserveSourceAndCacheTypes` |
| `ClosedXmlPivotSourceValidator.Validate` / `ClosedXmlReportWriter.ApplyPivotTables` | 小写字段选择原始表头，正确行/值字段索引及完整类型化缓存 | `LowercaseFieldNames_ShouldSelectTypedSourceAndCache` |
| `ClosedXmlPivotSourceValidator.Validate` | 文本分组忽略大小写的完整来源缓存与精确行边界 | `CaseInsensitiveGroups_ShouldKeepCompleteSourceCacheAtExactBoundary` |
| `ClosedXmlPivotSourceValidator.Validate` / `InvalidSource` | 空/纯空白/大小写重复表头、空行、非数值及所有来源字段错误单元格的结构化原始坐标 | `InvalidSource_ShouldReportCoordinateAndPreserveOutput` |
| `ClosedXmlPivotSourceValidator.Validate` / `ClosedXmlReportWriter.ApplyPivotTables` | 带首尾空格字段按原文绑定，完整来源与全部缓存一致 | `PaddedHeaderFieldNames_ShouldBindExactlyAndPreserveCache` |
| `ClosedXmlPivotSourceValidator.Validate` / `InvalidSource` | 标题不自动修剪；缺失字段定位来源首个表头并保留目标流 | `UnpaddedFieldNames_ShouldNotMatchPaddedHeaders` |
| `ClosedXmlPivotSourceValidator.Validate` | long 展开预算及越界来源位置 | `ExpandedOutputBudget_ShouldReportSourceCoordinate` |
| `ClosedXmlReportWriter.ApplyPivotTables` / `ClosedXmlExcelExporter.ExportToFile` / `ExportToFileAsync` | 文件来源预检失败保全原目标、无残留暂存、观察一次 | `FileExport_FormulaSourceFailure_ShouldPreserveTargetAndObserveOnce` |
| `ClosedXmlReportWriter.ApplyPivotTables` / `ClosedXmlExcelExporter.Export` / 文件入口 | 后一个透视来源失败时整次输出保留旧流与旧文件、临时文件清理并观察一次 | `LaterPivotSourceFailure_ShouldPreserveStreamAndFile` |
| `ClosedXmlExcelExporter.Export` / `ExportAsync` | 预取消保留来源和调用方流 | `PreCanceledSourceExport_ShouldPreserveStreamAndTemplate` |
| `ClosedXmlExcelExporter.Export` / `ExportAsync` / 文件入口 | 明细生成中取消及旧目标保护 | `SourceGenerationCancellation_ShouldPreserveStreamAndFile` |
| `ClosedXmlExcelExporter.ExportToFile` / `ExportToFileAsync` | 暂存实际写入中取消、旧目标及临时文件清理 | `FileWriteCancellation_ShouldPreserveTargetAndCleanTemporaryFile` |

既有 `PivotTableContractTest`、`PivotTableResourcesTest`、`PivotTableLookupCompositionTest` 继续直接覆盖完整透视结构、请求快照、格式边界、辅助 Sheet 分配和提交失败合同。双 TFM 验证产物见 execution 的最终门禁记录。

`Bing.Offices.Consumer.Net6` 与 `Bing.Offices.Consumer.Net8` 从本阶段生成的 NuGet 包验证完整纯值来源及缓存、字典名称顺延，以及同步/异步公式来源拒绝和调用方流保护。
